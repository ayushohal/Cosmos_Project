import fitz
import re
import json

def extract_castings_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    all_text = "\n".join(page.get_text("text") for page in doc)
    doc.close()

    # Split lines and remove empty lines
    lines = [line.strip() for line in all_text.splitlines() if line.strip()]

    castings = []
    current_casting = None
    current_wall = None

    for line in lines:
        # Detect casting header (keep this as it was)
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

    return {"castings": castings}


# --- MAIN ---
if __name__ == "__main__":
    pdf_path = input("Enter PDF filename: ").strip()
    result = extract_castings_from_pdf(pdf_path)

    # Save JSON
    with open("output.json", "w") as f:
        json.dump(result, f, indent=2)

    # Print preview
    print(json.dumps(result, indent=2))
    print("\n✅ Extraction complete. JSON saved as output.json")
