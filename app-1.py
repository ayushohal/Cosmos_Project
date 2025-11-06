from flask import Flask, request, jsonify, render_template, send_file
from flask_cors import CORS
import os
import json
import logging
from werkzeug.utils import secure_filename
from io import StringIO
import sys
import warnings
warnings.filterwarnings("ignore", category=UserWarning)

import fitz  # PyMuPDF
import re
from typing import List, Dict, Tuple
import itertools
import openpyxl
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# Configuration
UPLOAD_FOLDER = 'Uploads'
ALLOWED_EXTENSIONS = {'pdf'}
RESULTS_FILE = 'optimization_results.json'

# Constants
MIN_PANEL_SIZE = 100
MAX_PANEL_SIZE = 600
STANDARD_PANEL_SIZES = [100, 200, 300, 400, 500, 600]

# IC/EC Constants - Hardcoded shape requirements
SHAPE_IC_EC_REQUIREMENTS = {
    'column': {'IC': 0, 'EC': 4},
    'l-shape': {'IC': 1, 'EC': 5},
    'e-shape': {'IC': 4, 'EC': 8},
    'u-shape': {'IC': 2, 'EC': 6},
    'lift': {'IC': 2, 'EC': 6},
    't-shape': {'IC': 2, 'EC': 6},
    'i-shape': {'IC': 4, 'EC': 8}
}

# Cache for storing previously computed panel combinations
panel_combinations_cache = {}

# Setup
app = Flask(__name__)
CORS(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Logger setup
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# --- TEXT-BASED EXTRACTION FUNCTION (from text_extraction.py) ---

def extract_castings_from_pdf(pdf_path):
    """
    Extract casting data from PDF using text-based extraction.
    Returns structured casting data compatible with optimization logic.
    """
    doc = fitz.open(pdf_path)
    all_text = "\n".join(page.get_text("text") for page in doc)
    doc.close()

    # Split lines and remove empty lines
    lines = [line.strip() for line in all_text.splitlines() if line.strip()]

    castings = []
    current_casting = None
    current_wall = None

    for line in lines:
        # Detect casting header
        if re.search(r'(?i)casting[-\s]*\d+', line):
            if current_casting:
                castings.append(current_casting)
            current_casting = {"casting_number": line.strip(), "equipment_groups": []}
            current_wall = None
            continue

        # Any line containing letters -> wall ID
        elif re.search(r'[a-zA-Z]', line):
            current_wall = {"id": line.strip(), "sides": []}
            if current_casting:
                current_casting["equipment_groups"].append(current_wall)
            continue

        # Any line containing only numbers -> dimension
        elif re.match(r'^\d+(\.\d+)?$', line):
            if current_wall:
                current_wall["sides"].append(float(line))
            continue

        # Ignore other lines
        else:
            continue

    # Append last casting
    if current_casting:
        castings.append(current_casting)

    # Convert to format expected by optimization logic
    castings_data = []
    for casting in castings:
        casting_name = casting["casting_number"]
        shapes_data = {}
        
        for wall in casting["equipment_groups"]:
            wall_id = wall["id"]
            sides = wall["sides"]
            
            # Create shape with all sides
            shape_name = f"Shape_{wall_id}"
            sides_dict = {}
            for i, side in enumerate(sides, 1):
                sides_dict[f"side{i}"] = int(side)
            
            shapes_data[shape_name] = sides_dict
        
        if shapes_data:
            castings_data.append({
                "name": casting_name,
                "shapes": shapes_data
            })
    
    return castings_data

# --- HELPER FUNCTIONS for SHAPE CLASSIFICATION ---

def classify_shape_type(shape_name: str) -> str:
    """
    Classify shape type based on name patterns for IC/EC calculations.
    Returns the shape type key for SHAPE_IC_EC_REQUIREMENTS lookup.
    """
    shape_lower = shape_name.lower()
    
    # Check for specific patterns
    if 'sw' in shape_lower or 'column' in shape_lower:
        return 'column'
    elif 'lift' in shape_lower or 'lsw' in shape_lower:
        return 'lift'
    elif 'l' in shape_lower and ('shape' in shape_lower or 'lw' in shape_lower):
        return 'l-shape'
    elif 'e' in shape_lower and ('shape' in shape_lower or len(shape_lower) <= 3):
        return 'e-shape'
    elif 'u' in shape_lower and ('shape' in shape_lower or len(shape_lower) <= 3):
        return 'u-shape'
    elif 't' in shape_lower and ('shape' in shape_lower or len(shape_lower) <= 3):
        return 't-shape'
    elif 'i' in shape_lower and ('shape' in shape_lower or len(shape_lower) <= 3):
        return 'i-shape'
    
    # Default to column if no specific pattern matches
    return 'column'

# --- PANEL OPTIMIZATION CLASSES AND FUNCTIONS ---

class Shape:
    def __init__(self, name: str, sides: List[int]):
        self.name = name
        self.sides = sides
        self.panel_layout = [[] for _ in range(len(sides))]
        self.shape_type = classify_shape_type(name)
        self.ic_requirement = SHAPE_IC_EC_REQUIREMENTS[self.shape_type]['IC']
        self.ec_requirement = SHAPE_IC_EC_REQUIREMENTS[self.shape_type]['EC']
    
    def get_total_length(self) -> int:
        return sum(self.sides)
    
    def __str__(self) -> str:
        return f"Shape: {self.name}, Sides: {self.sides}, Type: {self.shape_type}, IC: {self.ic_requirement}, EC: {self.ec_requirement}"

class Casting:
    def __init__(self, name: str):
        self.name = name
        self.shapes = []
    
    def add_shape(self, shape: Shape) -> None:
        self.shapes.append(shape)
    
    def get_total_length(self) -> int:
        return sum(shape.get_total_length() for shape in self.shapes)
    
    def get_total_ic_requirement(self) -> int:
        return sum(shape.ic_requirement for shape in self.shapes)
    
    def get_total_ec_requirement(self) -> int:
        return sum(shape.ec_requirement for shape in self.shapes)
    
    def __str__(self) -> str:
        return f"Casting: {self.name}, Shapes: {len(self.shapes)}, IC: {self.get_total_ic_requirement()}, EC: {self.get_total_ec_requirement()}"

def analyze_castings(castings: List[Casting]) -> Dict:
    """
    Analyze all castings to identify common dimensions and optimal panel sizes.
    Returns information about preferred panel sizes for optimization.
    """
    length_counts = {}
    common_divisors = {}
    
    all_lengths = []
    for casting in castings:
        for shape in casting.shapes:
            all_lengths.extend(shape.sides)
    
    for length in all_lengths:
        length_counts[length] = length_counts.get(length, 0) + 1
    
    for panel_size in sorted(STANDARD_PANEL_SIZES, reverse=True):
        divisible_count = 0
        total_panels = 0
        
        for length in all_lengths:
            if length % panel_size == 0:
                divisible_count += 1
                total_panels += length // panel_size
            elif length % panel_size <= MIN_PANEL_SIZE and length >= panel_size:
                divisible_count += 0.5
                total_panels += length // panel_size
        
        efficiency = divisible_count / len(all_lengths) if all_lengths else 0
        common_divisors[panel_size] = {
            "efficiency": efficiency,
            "divisible_count": divisible_count,
            "total_panels": total_panels
        }
    
    panel_efficiency = sorted(
        [(size, data["efficiency"], data["total_panels"]) 
         for size, data in common_divisors.items()],
        key=lambda x: (-x[1], x[2])
    )
    
    preferred_sizes = [size for size, efficiency, _ in panel_efficiency if efficiency > 0.3]
    
    if len(preferred_sizes) < 2:
        for size in sorted(STANDARD_PANEL_SIZES, reverse=True):
            if size not in preferred_sizes:
                preferred_sizes.append(size)
            if len(preferred_sizes) >= 3:
                break
    
    return {
        "length_counts": length_counts,
        "preferred_sizes": preferred_sizes,
        "panel_efficiency": panel_efficiency
    }

def get_possible_panels(length: int) -> List[List[int]]:
    """
    Generate optimal panel combinations for a given length.
    Ensures all panels are within the valid size range (MIN_PANEL_SIZE to MAX_PANEL_SIZE).
    """
    if length in panel_combinations_cache:
        return panel_combinations_cache[length]
    
    valid_panels = []
    standard_sizes = sorted(STANDARD_PANEL_SIZES, reverse=True)
    
    if length < MIN_PANEL_SIZE:
        valid_panels.append([MIN_PANEL_SIZE])
        panel_combinations_cache[length] = valid_panels
        return valid_panels
    
    if length >= MAX_PANEL_SIZE:
        num_max_panels = length // MAX_PANEL_SIZE
        remaining = length - (num_max_panels * MAX_PANEL_SIZE)
        
        if remaining == 0:
            valid_panels.append([MAX_PANEL_SIZE] * num_max_panels)
        elif remaining >= MIN_PANEL_SIZE:
            valid_panels.append([MAX_PANEL_SIZE] * num_max_panels + [remaining])
        else:
            adjusted_length = MAX_PANEL_SIZE + remaining
            for r in range(1, 3):
                for combo in itertools.combinations_with_replacement(standard_sizes, r):
                    if sum(combo) == adjusted_length and all(p >= MIN_PANEL_SIZE for p in combo):
                        valid_panels.append([MAX_PANEL_SIZE] * (num_max_panels - 1) + list(combo))
            if not any(p for p in valid_panels if sum(p) == length) and adjusted_length >= MIN_PANEL_SIZE:
                valid_panels.append([MAX_PANEL_SIZE] * (num_max_panels - 1) + [adjusted_length])
    
    for r in range(1, min(8, length // MIN_PANEL_SIZE + 1)):
        for combo in itertools.combinations_with_replacement(standard_sizes, r):
            if sum(combo) == length:
                valid_panels.append(list(combo))
    
    for size1 in standard_sizes:
        if size1 <= length:
            max_count1 = min(length // size1, 8)
            for count1 in range(1, max_count1 + 1):
                remaining1 = length - (size1 * count1)
                if remaining1 == 0:
                    valid_panels.append([size1] * count1)
                elif remaining1 >= MIN_PANEL_SIZE:
                    for size2 in standard_sizes:
                        if size2 <= remaining1:
                            if remaining1 % size2 == 0:
                                count2 = remaining1 // size2
                                valid_panels.append([size1] * count1 + [size2] * count2)
                            elif remaining1 > size2 and (remaining1 % size2) >= MIN_PANEL_SIZE:
                                count2 = remaining1 // size2
                                last_panel = remaining1 - (size2 * count2)
                                if last_panel >= MIN_PANEL_SIZE:
                                    valid_panels.append([size1] * count1 + [size2] * count2 + [last_panel])
    
    if length <= MAX_PANEL_SIZE and length >= MIN_PANEL_SIZE:
        if length in STANDARD_PANEL_SIZES:
            valid_panels.append([length])
        else:
            valid_panels.append([length])
    
    valid_panels = [
        combo for combo in valid_panels 
        if all(MIN_PANEL_SIZE <= p <= MAX_PANEL_SIZE for p in combo) and sum(combo) == length
    ]
    
    verified_panels = []
    seen = set()
    for combo in valid_panels:
        combo_tuple = tuple(sorted(combo))
        if combo_tuple not in seen:
            verified_panels.append(combo)
            seen.add(combo_tuple)
    
    sorted_panels = sorted(
        verified_panels,
        key=lambda x: (
            sum(0 if p in STANDARD_PANEL_SIZES else 1 for p in x),
            len(x),
            -sum(p for p in x) / len(x) if x else 0,
            -sum(1 for p in x if p == MAX_PANEL_SIZE)
        )
    )
    
    if not sorted_panels and length > 0:
        if length < MIN_PANEL_SIZE:
            sorted_panels = [[MIN_PANEL_SIZE]]
        elif length <= MAX_PANEL_SIZE:
            sorted_panels = [[length]]
        else:
            max_count = length // MAX_PANEL_SIZE
            remaining = length % MAX_PANEL_SIZE
            if remaining >= MIN_PANEL_SIZE:
                sorted_panels = [[MAX_PANEL_SIZE] * max_count + [remaining]]
            else:
                adjusted = [MAX_PANEL_SIZE] * (max_count - 1)
                remaining_length = MAX_PANEL_SIZE + remaining
                for size in sorted(standard_sizes, reverse=True):
                    if remaining_length >= size + MIN_PANEL_SIZE:
                        final_remainder = remaining_length - size
                        if final_remainder >= MIN_PANEL_SIZE:
                            sorted_panels = [adjusted + [size, final_remainder]]
                            break
                if not sorted_panels:
                    if (max_count-1) * MAX_PANEL_SIZE >= length - MIN_PANEL_SIZE:
                        remaining = length - (max_count-1) * MAX_PANEL_SIZE
                        sorted_panels = [[MAX_PANEL_SIZE] * (max_count-1) + [remaining]]
                    else:
                        min_count = (length + MIN_PANEL_SIZE - 1) // MIN_PANEL_SIZE
                        sorted_panels = [[MIN_PANEL_SIZE] * min_count]
    
    panel_combinations_cache[length] = sorted_panels
    return sorted_panels

def optimize_panels_and_accessories(castings: List[Casting], primary_idx: int) -> Dict:
    """
    Optimize panel layout and IC/EC accessories with cumulative reuse: 
    panels and accessories from all previous castings are available for reuse in each subsequent casting.
    """
    print("\nOptimizing panel layouts and accessories with cumulative reuse...")
    
    ordered_castings = castings
    panel_inventory = {}
    ic_inventory = 0
    ec_inventory = 0
    
    new_panels_added = {}
    new_ic_added = 0
    new_ec_added = 0
    
    panels_used_per_casting = []
    ic_used_per_casting = []
    ec_used_per_casting = []

    for cast_idx, casting in enumerate(ordered_castings):
        casting_panels_used = {}
        casting_ic_used = casting.get_total_ic_requirement()
        casting_ec_used = casting.get_total_ec_requirement()
        
        # Handle panel requirements
        for shape in casting.shapes:
            for side_idx, side_length in enumerate(shape.sides):
                layouts = get_possible_panels(side_length)
                if not layouts:
                    continue
                layout = layouts[0]
                shape.panel_layout[side_idx] = layout.copy()
                for panel in layout:
                    if panel_inventory.get(panel, 0) > 0:
                        panel_inventory[panel] -= 1
                    else:
                        new_panels_added[panel] = new_panels_added.get(panel, 0) + 1
                    casting_panels_used[panel] = casting_panels_used.get(panel, 0) + 1
        
        # Handle IC requirements
        if ic_inventory >= casting_ic_used:
            ic_inventory -= casting_ic_used
        else:
            ic_needed = casting_ic_used - ic_inventory
            new_ic_added += ic_needed
            ic_inventory = 0
        
        # Handle EC requirements
        if ec_inventory >= casting_ec_used:
            ec_inventory -= casting_ec_used
        else:
            ec_needed = casting_ec_used - ec_inventory
            new_ec_added += ec_needed
            ec_inventory = 0
        
        # Add current casting's resources to inventory
        for panel, count in casting_panels_used.items():
            panel_inventory[panel] = panel_inventory.get(panel, 0) + count
        ic_inventory += casting_ic_used
        ec_inventory += casting_ec_used
        
        # Store usage per casting
        panels_used_per_casting.append(casting_panels_used)
        ic_used_per_casting.append(casting_ic_used)
        ec_used_per_casting.append(casting_ec_used)

    print(f"\nOptimization completed.")
    
    return {
        "new_panels_added": new_panels_added,
        "new_ic_added": new_ic_added,
        "new_ec_added": new_ec_added,
        "panels_used_per_casting": panels_used_per_casting,
        "ic_used_per_casting": ic_used_per_casting,
        "ec_used_per_casting": ec_used_per_casting,
        "panel_inventory": panel_inventory,
        "final_ic_inventory": ic_inventory,
        "final_ec_inventory": ec_inventory
    }

def print_results_with_accessories(castings: List[Casting], primary_idx: int, optimization_results: Dict) -> None:
    """Print the optimized panel layouts and IC/EC requirements for all castings with detailed cumulative reuse analysis."""
    print(f"\nResults (Primary Casting: {castings[primary_idx].name})\n")

    panels_used_per_casting = optimization_results["panels_used_per_casting"]
    ic_used_per_casting = optimization_results["ic_used_per_casting"]
    ec_used_per_casting = optimization_results["ec_used_per_casting"]
    new_panels_added = optimization_results["new_panels_added"]
    new_ic_added = optimization_results["new_ic_added"]
    new_ec_added = optimization_results["new_ec_added"]

    for i, casting in enumerate(castings):
        print(f"{'*' * 20} {casting.name} {'*' * 20}")
        print("PRIMARY" if i == primary_idx else f"SECONDARY #{i if i > primary_idx else i+1}")
        
        for shape in casting.shapes:
            print(f"\n  Shape: {shape.name} (Type: {shape.shape_type})")
            print(f"    IC Requirement: {shape.ic_requirement}, EC Requirement: {shape.ec_requirement}")
            for side_idx, side_length in enumerate(shape.sides):
                panels = shape.panel_layout[side_idx]
                print(f"    Side {side_idx+1} (Length: {side_length}): {panels}")
        
        print(f"  Panels used in this casting: {panels_used_per_casting[i]}")
        print(f"  IC used in this casting: {ic_used_per_casting[i]}")
        print(f"  EC used in this casting: {ec_used_per_casting[i]}")

    print("\n" + "=" * 50)
    print("PANEL & ACCESSORIES USAGE SUMMARY")
    print("=" * 50)
    print(f"Total unique panel types used: {len(optimization_results['panel_inventory'])}")
    print(f"Panel inventory after all castings: {optimization_results['panel_inventory']}")
    print(f"IC inventory after all castings: {optimization_results['final_ic_inventory']}")
    print(f"EC inventory after all castings: {optimization_results['final_ec_inventory']}")
    print(f"Total new panels needed: {sum(new_panels_added.values())}")
    print(f"Total new IC needed: {new_ic_added}")
    print(f"Total new EC needed: {new_ec_added}")
    print(f"Breakdown of new panels needed: {new_panels_added}")
    
    print("\n" + "=" * 50)
    print("REQUIREMENTS BY CASTING")
    print("=" * 50)
    
    total_new_panels = 0
    total_new_ic = 0
    total_new_ec = 0
    
    # Primary casting
    primary_panels = sum(panels_used_per_casting[primary_idx].values())
    primary_ic = ic_used_per_casting[primary_idx]
    primary_ec = ec_used_per_casting[primary_idx]
    
    print(f"{castings[primary_idx].name} (PRIMARY): {primary_panels} panels, {primary_ic} IC, {primary_ec} EC needed")
    total_new_panels += primary_panels
    total_new_ic += primary_ic
    total_new_ec += primary_ec
    
    # Secondary castings
    cumulative_panels = panels_used_per_casting[primary_idx].copy()
    cumulative_ic = primary_ic
    cumulative_ec = primary_ec
    
    for i in range(len(castings)):
        if i == primary_idx:
            continue
            
        used_panels = panels_used_per_casting[i]
        used_ic = ic_used_per_casting[i]
        used_ec = ec_used_per_casting[i]
        
        # Calculate panel reuse
        new_panels_needed = 0
        temp_inventory = cumulative_panels.copy()
        for panel, count in used_panels.items():
            reused_count = min(temp_inventory.get(panel, 0), count)
            new_panels_needed += count - reused_count
            temp_inventory[panel] = temp_inventory.get(panel, 0) - reused_count
        
        # Calculate IC/EC reuse
        new_ic_needed = max(0, used_ic - cumulative_ic)
        new_ec_needed = max(0, used_ec - cumulative_ec)
        
        print(f"{castings[i].name} (SECONDARY): {new_panels_needed} panels, {new_ic_needed} IC, {new_ec_needed} EC needed")
        total_new_panels += new_panels_needed
        total_new_ic += new_ic_needed
        total_new_ec += new_ec_needed
        
        # Update cumulative inventory
        for panel, count in used_panels.items():
            cumulative_panels[panel] = cumulative_panels.get(panel, 0) + count
        cumulative_ic += used_ic
        cumulative_ec += used_ec
    
    print(f"\nTotal new resources needed for entire project:")
    print(f"  Panels: {total_new_panels}")
    print(f"  IC: {total_new_ic}")
    print(f"  EC: {total_new_ec}")

    # Calculate overall efficiency
    total_panels_used = sum(sum(casting_panels.values()) for casting_panels in panels_used_per_casting)
    total_ic_used = sum(ic_used_per_casting)
    total_ec_used = sum(ec_used_per_casting)
    
    total_panels_reused = total_panels_used - total_new_panels
    total_ic_reused = total_ic_used - total_new_ic
    total_ec_reused = total_ec_used - total_new_ec
    
    total_resources_used = total_panels_used + total_ic_used + total_ec_used
    total_resources_reused = total_panels_reused + total_ic_reused + total_ec_reused
    
    if total_resources_used > 0:
        overall_efficiency = (total_resources_reused / total_resources_used) * 100
        panel_efficiency = (total_panels_reused / total_panels_used) * 100 if total_panels_used > 0 else 0
        ic_efficiency = (total_ic_reused / total_ic_used) * 100 if total_ic_used > 0 else 0
        ec_efficiency = (total_ec_reused / total_ec_used) * 100 if total_ec_used > 0 else 0
        
        print(f"\nOverall resource reuse efficiency: {overall_efficiency:.1f}%")
        print(f"  Panel reuse efficiency: {panel_efficiency:.1f}% ({total_panels_reused} of {total_panels_used} panels reused)")
        print(f"  IC reuse efficiency: {ic_efficiency:.1f}% ({total_ic_reused} of {total_ic_used} IC reused)")
        print(f"  EC reuse efficiency: {ec_efficiency:.1f}% ({total_ec_reused} of {total_ec_used} EC reused)")
    else:
        print("\nResource reuse efficiency: N/A (no castings)")

def convert_extracted_data_to_castings(castings_data: List[Dict]) -> List[Casting]:
    """Convert extracted casting data to Casting objects for optimization."""
    castings = []
    
    for casting_data in castings_data:
        casting = Casting(casting_data["name"])
        for shape_name, sides_data in casting_data["shapes"].items():
            sides = [length for _, length in sides_data.items()]
            shape = Shape(shape_name, sides)
            casting.add_shape(shape)
        castings.append(casting)
    
    return castings

def export_to_excel_with_accessories(panel_data, ic_count, ec_count):
    """Export panels and accessories to Excel with IC/EC requirements included."""
    sorted_panels = sorted(panel_data, key=lambda x: x["width"], reverse=True)
    wb = Workbook()
    ws = wb.active

    ws.merge_cells('A1:H1')
    ws['A1'] = "COSMOS CONSTRUCTION MACHINERIES & EQUIPMENT PVT.LTD."
    ws['A1'].font = Font(size=14, bold=True)
    ws['A1'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A2:H2')
    ws['A2'] = "Plot No.E/307/5, Gat No.307, Nanekarwadi, Next to MINDA, Chakan, Pune. Phone -+91 86008 84281"
    ws['A2'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A3:H3')
    ws['A3'] = "Email: aluformservice@cosmossales.com"
    ws['A3'].alignment = Alignment(horizontal='center')

    ws.merge_cells('A5:H5')
    ws['A5'] = "TOTAL CONSOLIDATED MATERIAL QUANTITY"
    ws['A5'].alignment = Alignment(horizontal='center')
    ws['A5'].font = Font(bold=True)

    ws.merge_cells('A6:H6')
    ws['A6'] = "SHUTTERING AREA STATEMENT"
    ws['A6'].alignment = Alignment(horizontal='center')
    ws['A6'].font = Font(bold=True)

    ws.merge_cells('A7:G7')
    ws['A7'] = "Project Name"
    ws['A7'].font = Font(bold=True)
    ws['A7'].alignment = Alignment(horizontal='center')

    ws['H7'] = datetime.now().strftime("%d.%m.%Y")
    ws['H7'].font = Font(bold=True)
    ws['H7'].alignment = Alignment(horizontal='right')

    table_headers = [
        "Sr. No.", "Panel Size (mm)", "Width (mm)", "Code", "Length (mm)",
        "Total No. Of Panel Quantity", "Unit Area of Panel in Sq.m", "Total Area in Sq.m of all Panels"
    ]
    ws.append(table_headers)
    for col in range(1, len(table_headers) + 1):
        ws.cell(row=8, column=col).font = Font(bold=True)
        ws.cell(row=8, column=col).alignment = Alignment(horizontal='center')

    total_area_sum = 0
    total_qty_sum = 0
    row_num = 9
    
    # Add panels
    for idx, panel in enumerate(sorted_panels, start=1):
        width = panel["width"]
        length = panel["length"]
        code = panel["code"]
        qty = panel["quantity"]
        panel_size_str = f"{width}{code}{length}"
        unit_area = round((width * length) / 1_000_000, 3)
        total_area = round(unit_area * qty, 2)
        total_area_sum += total_area
        total_qty_sum += qty
        ws.append([
            idx,
            panel_size_str,
            width,
            code,
            length,
            qty,
            unit_area,
            total_area
        ])
        row_num += 1
    
    # Add IC accessories if needed
    if ic_count > 0:
        ws.append([
            row_num - 8,
            "IC (Internal Corner)",
            "-",
            "IC",
            "-",
            ic_count,
            "-",
            "-"
        ])
        row_num += 1
        total_qty_sum += ic_count
    
    # Add EC accessories if needed
    if ec_count > 0:
        ws.append([
            row_num - 8,
            "EC (External Corner)",
            "-",
            "EC",
            "-",
            ec_count,
            "-",
            "-"
        ])
        row_num += 1
        total_qty_sum += ec_count

    ws.append(["", "", "", "", "", total_qty_sum])
    ws.cell(row=ws.max_row, column=6).font = Font(bold=True)
    ws.append(["", "", "", "", "", "", "Total Shuttering Area in Sq.m", round(total_area_sum, 2)])
    ws.cell(row=ws.max_row, column=7).font = Font(bold=True)
    ws.cell(row=ws.max_row, column=8).font = Font(bold=True)

    for col in range(1, 9):
        ws.column_dimensions[get_column_letter(col)].width = 20

    wb.save("consolidated_list.xlsx")
    print("✅ Excel file 'consolidated_list.xlsx' generated successfully with IC/EC requirements.")

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/extract-castings', methods=['POST'])
def extract_castings_route():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['pdf']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            file.save(file_path)
            
            # Extract castings using text-based extraction
            castings_data = extract_castings_from_pdf(file_path)
            
            # Display extracted data for verification
            print("\n" + "=" * 80)
            print("EXTRACTED CASTING DATA (Text-Based Extraction)")
            print("=" * 80)
            print(json.dumps(castings_data, indent=2))
            print("=" * 80 + "\n")
            
            os.remove(file_path)
            return jsonify(castings_data), 200
        except Exception as e:
            logging.exception("Failed to process PDF")
            if os.path.exists(file_path):
                os.remove(file_path)
            return jsonify({'error': f"Failed to process PDF: {str(e)}"}), 500
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/api/optimize', methods=['POST']) 
def optimize():
    try:
        data = request.get_json()
        if not data or 'castings' not in data or 'primaryIdx' not in data:
            return jsonify({'error': 'Invalid request data'}), 400

        # Display received casting data for verification
        print("\n" + "=" * 80)
        print("RECEIVED CASTING DATA FOR OPTIMIZATION")
        print("=" * 80)
        print(json.dumps(data['castings'], indent=2))
        print("=" * 80 + "\n")

        castings = convert_extracted_data_to_castings(data['castings'])
        primary_idx = int(data['primaryIdx'])

        if primary_idx < 0 or primary_idx >= len(castings):
            return jsonify({'error': f"Invalid primaryIdx: {primary_idx}"}), 400

        # Analyze castings
        analysis_results = analyze_castings(castings)

        # Perform optimization with accessories
        optimization_results = optimize_panels_and_accessories(castings, primary_idx)

        # Capture print_results output
        original_stdout = sys.stdout
        sys.stdout = buffer = StringIO()
        print_results_with_accessories(castings, primary_idx, optimization_results)
        sys.stdout = original_stdout
        text_summary = buffer.getvalue()

        # Prepare panel_data for Excel export
        new_panels_added = optimization_results["new_panels_added"]
        new_ic_added = optimization_results["new_ic_added"]
        new_ec_added = optimization_results["new_ec_added"]
        
        panel_data = [
            {
                "width": panel_size,
                "length": 2400,
                "code": "WS",
                "quantity": quantity
            }
            for panel_size, quantity in new_panels_added.items()
        ]

        # Export to Excel with accessories
        export_to_excel_with_accessories(panel_data, new_ic_added, new_ec_added)

        # Calculate overall efficiency including accessories
        total_panels_used = sum(sum(casting_panels.values()) for casting_panels in optimization_results["panels_used_per_casting"])
        total_ic_used = sum(optimization_results["ic_used_per_casting"])
        total_ec_used = sum(optimization_results["ec_used_per_casting"])
        total_resources_used = total_panels_used + total_ic_used + total_ec_used
        
        total_new_panels = sum(new_panels_added.values())
        total_panels_reused = total_panels_used - total_new_panels
        total_ic_reused = total_ic_used - new_ic_added
        total_ec_reused = total_ec_used - new_ec_added
        total_resources_reused = total_panels_reused + total_ic_reused + total_ec_reused
        
        overall_efficiency = (total_resources_reused / total_resources_used * 100) if total_resources_used > 0 else 0
        panel_efficiency = (total_panels_reused / total_panels_used * 100) if total_panels_used > 0 else 0
        ic_efficiency = (total_ic_reused / total_ic_used * 100) if total_ic_used > 0 else 0
        ec_efficiency = (total_ec_reused / total_ec_used * 100) if total_ec_used > 0 else 0

        # Prepare results for JSON
        results = {
            'castings': [],
            'optimization_summary': {
                'total_castings': len(castings),
                'primary_casting': castings[primary_idx].name,
                'panel_sizes_used': list(set(
                    panel for casting in castings
                    for shape in casting.shapes
                    for panels in shape.panel_layout
                    for panel in panels
                )),
                'total_new_panels_needed': total_new_panels,
                'total_new_ic_needed': new_ic_added,
                'total_new_ec_needed': new_ec_added,
                'overall_efficiency': {
                    'total_efficiency': round(overall_efficiency, 1),
                    'panel_efficiency': round(panel_efficiency, 1),
                    'ic_efficiency': round(ic_efficiency, 1),
                    'ec_efficiency': round(ec_efficiency, 1),
                    'total_resources_used': total_resources_used,
                    'total_resources_reused': total_resources_reused
                },
                'panel_reuse_analysis': {
                    'primary_casting': {
                        'name': castings[primary_idx].name,
                        'panels_used': optimization_results["panels_used_per_casting"][primary_idx],
                        'ic_used': optimization_results["ic_used_per_casting"][primary_idx],
                        'ec_used': optimization_results["ec_used_per_casting"][primary_idx],
                        'new_panels_needed': optimization_results["panels_used_per_casting"][primary_idx],
                        'new_ic_needed': optimization_results["ic_used_per_casting"][primary_idx],
                        'new_ec_needed': optimization_results["ec_used_per_casting"][primary_idx]
                    },
                    'secondary_castings': []
                },
                'panel_analysis': {
                    'length_counts': analysis_results['length_counts'],
                    'preferred_sizes': analysis_results['preferred_sizes'],
                    'panel_efficiency': analysis_results['panel_efficiency']
                },
                'text_summary': text_summary
            }
        }

        # Add secondary casting analysis with IC/EC
        panels_used_per_casting = optimization_results["panels_used_per_casting"]
        ic_used_per_casting = optimization_results["ic_used_per_casting"]
        ec_used_per_casting = optimization_results["ec_used_per_casting"]
        
        cumulative_panels = panels_used_per_casting[primary_idx].copy()
        cumulative_ic = ic_used_per_casting[primary_idx]
        cumulative_ec = ec_used_per_casting[primary_idx]
        
        for i in range(len(castings)):
            if i == primary_idx:
                continue
                
            used_panels = panels_used_per_casting[i]
            used_ic = ic_used_per_casting[i]
            used_ec = ec_used_per_casting[i]
            
            # Calculate panel reuse
            reused_panels = {}
            new_panels_needed = {}
            temp_inventory = cumulative_panels.copy()
            
            for panel, count in used_panels.items():
                reused_count = min(temp_inventory.get(panel, 0), count)
                reused_panels[panel] = reused_count
                new_panels_needed[panel] = count - reused_count
                temp_inventory[panel] = temp_inventory.get(panel, 0) - reused_count
            
            # Calculate IC/EC reuse
            reused_ic = min(cumulative_ic, used_ic)
            new_ic_needed = max(0, used_ic - cumulative_ic)
            reused_ec = min(cumulative_ec, used_ec)
            new_ec_needed = max(0, used_ec - cumulative_ec)
            
            results["optimization_summary"]["panel_reuse_analysis"]["secondary_castings"].append({
                "name": castings[i].name,
                "panels_used": used_panels,
                "ic_used": used_ic,
                "ec_used": used_ec,
                "reused_panels_from_previous": reused_panels,
                "reused_ic_from_previous": reused_ic,
                "reused_ec_from_previous": reused_ec,
                "new_panels_needed": new_panels_needed,
                "new_ic_needed": new_ic_needed,
                "new_ec_needed": new_ec_needed
            })
            
            # Update cumulative inventory
            for panel, count in used_panels.items():
                cumulative_panels[panel] = cumulative_panels.get(panel, 0) + count
            cumulative_ic += used_ic
            cumulative_ec += used_ec

        # Add casting details with IC/EC information
        for casting in castings:
            casting_data = {
                'name': casting.name,
                'total_ic_requirement': casting.get_total_ic_requirement(),
                'total_ec_requirement': casting.get_total_ec_requirement(),
                'shapes': []
            }
            for shape in casting.shapes:
                shape_data = {
                    'name': shape.name,
                    'shape_type': shape.shape_type,
                    'sides': shape.sides,
                    'panel_layouts': shape.panel_layout,
                    'ic_requirement': shape.ic_requirement,
                    'ec_requirement': shape.ec_requirement
                }
                casting_data['shapes'].append(shape_data)
            results['castings'].append(casting_data)

        # Save results to file
        with open(RESULTS_FILE, 'w') as f:
            json.dump(results, f, indent=2)

        return jsonify(results), 200

    except ValueError as ve:
        logging.error(f"ValueError: {str(ve)}")
        return jsonify({'error': f"Invalid data: {str(ve)}"}), 400
    except Exception as e:
        logging.exception("Optimization failed")
        return jsonify({'error': f"Optimization failed: {str(e)}"}), 500

@app.route('/api/download-excel', methods=['GET'])
def download_excel():
    excel_file = "consolidated_list.xlsx"
    try:
        if os.path.exists(excel_file):
            return send_file(excel_file, as_attachment=True, download_name='consolidated_list.xlsx')
        else:
            return jsonify({'error': 'Excel file not found. Please run optimization first.'}), 404
    except Exception as e:
        logging.exception("Failed to download Excel file")
        return jsonify({'error': f"Failed to download Excel file: {str(e)}"}), 500

if __name__ == '__main__':
    logging.info("Starting Flask server on http://localhost:5000")
    app.run(debug=True, port=5000)