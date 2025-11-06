import requests
import base64
import json
import os
import time
import dotenv
import fitz  # PyMuPDF
import easyocr
import re

dotenv.load_dotenv()
API_KEY = os.getenv("OPENROUTER_API_KEY")
MODEL = "meta-llama/llama-4-maverick:free"
API_URL = "https://openrouter.ai/api/v1/chat/completions"
RETRY_DELAY = 10
MAX_RETRIES = 20

# Initialize OCR reader
reader = easyocr.Reader(["en"])

# Updated prompt for condensed, grouped JSON format
PROMPT = """
You are an expert structural engineer analyzing a casting diagram image. 
Extract all castings from the page and return a single JSON object.

Requirements:
1. Identify each casting (Casting-1st, Casting-2nd, etc.) from the image.
   Everything below a casting heading belongs to that casting until the next heading.

2. Group all equipment in each casting by type (SW, FSW, LSW, OTHER).

3. For Lift/Shear Wall or custom shapes:
   - Capture all dimensions for all sides and group them as one shape.
   - Read all sides/dimensions visible in the diagram.
   - Club them together as one shape in the JSON.
   - Do NOT stop at the first few dimensions; ensure all are captured.
 

4. Ensure no duplication or mixing between castings.

5. Return JSON in this format:
{
  "castings": [
    {
      "casting_number": "CASTING-1st",
      "equipment_groups": {...},
      "detailed_diagrams": {...}
    },
    {
      "casting_number": "CASTING-2nd",
      ...
    }
  ]
}

6. IMPORTANT: Return ONLY valid JSON starting with { and ending with }. 
Do NOT wrap in markdown or add explanations.
""" 

# --- FUNCTIONS ---

def pdf_to_images(pdf_path, dpi=300):
    """Convert PDF pages to high-resolution images using PyMuPDF."""
    images = []
    try:
        doc = fitz.open(pdf_path)
        print(f"📄 PDF has {len(doc)} pages")
        for i in range(len(doc)):
            page = doc[i]
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            img_bytes = pix.tobytes("png")
            images.append(img_bytes)
        doc.close()
        return images
    except Exception as e:
        print(f"❌ Error converting PDF: {e}")
        return []

def get_custom_shape_sides_hint(shape_img):
    """Use OCR to detect number of dimension lines in a custom shape."""
    ocr_results = reader.readtext(shape_img)
    dimension_patterns = [r"\d+X\d+", r"\d+"]  # Extend if needed
    dimensions_found = set()
    for (_, text, _) in ocr_results:
        for pattern in dimension_patterns:
            matches = re.findall(pattern, text)
            for m in matches:
                dimensions_found.add(m)
    num_sides = len(dimensions_found)
    if num_sides > 0:
        return f"This custom shape has {num_sides} sides with dimensions. Ensure all are captured."
    return ""

def analyze_casting_page(image_bytes, page_number):
    """Send image to OpenRouter API for analysis with retry logic."""
    headers = {
        "Authorization": f"Bearer {API_KEY}",
        "Content-Type": "application/json"
    }

    # Encode image to base64
    image_b64 = base64.b64encode(image_bytes).decode("utf-8")

    # Get hint for custom shapes
    hint_text = get_custom_shape_sides_hint(image_bytes)
    full_prompt = PROMPT
    if hint_text:
        full_prompt += f"\n{hint_text}"

    payload = {
        "model": MODEL,
        "messages": [
            {"role": "system", "content": "You are a structural engineering expert."},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": full_prompt},
                    {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{image_b64}"}}
                ]
            }
        ],
        "max_tokens": 4000,
        "temperature": 0.1
    }

    print(f"🔄 Analyzing page {page_number}...")
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.post(API_URL, headers=headers, json=payload, timeout=60)

            if response.status_code == 429:
                print(f"⏰ Rate limited (attempt {attempt}), waiting {RETRY_DELAY}s...")
                time.sleep(RETRY_DELAY)
                continue
            if response.status_code != 200:
                print(f"⚠️ HTTP {response.status_code}: {response.text}")
                time.sleep(RETRY_DELAY)
                continue

            result = response.json()
            raw_output = result["choices"][0]["message"]["content"].strip()

            # Clean markdown
            if raw_output.startswith("```"):
                raw_output = raw_output.strip("`").replace("json","").strip()

            try:
                parsed_json = json.loads(raw_output)
                # Save individual page JSON
                output_file = f"page_{page_number}_analysis.json"
                with open(output_file, "w", encoding="utf-8") as f:
                    json.dump(parsed_json, f, indent=2, ensure_ascii=False)
                print(f"✅ Page {page_number} saved: {output_file}")
                return parsed_json
            except json.JSONDecodeError:
                print(f"⚠️ Invalid JSON, skipping page {page_number}")
                return None

        except Exception as e:
            print(f"❌ Error on attempt {attempt}: {e}")
            time.sleep(RETRY_DELAY)
            continue

    print(f"❌ Failed page {page_number} after {MAX_RETRIES} retries")
    return None

def validate_setup():
    if not API_KEY:
        print("❌ Missing API key in .env")
        return False
    return True

# --- MAIN ---

def main():
    print("🏗️ Casting Diagram Analyzer")
    print("="*40)

    if not validate_setup():
        return

    pdf_path = input("📁 Enter path to PDF: ").strip().strip('"\'')
    if not os.path.exists(pdf_path):
        print(f"❌ File not found: {pdf_path}")
        return

    images = pdf_to_images(pdf_path)
    if not images:
        print("❌ Failed to convert PDF to images")
        return

    master_json = {"castings": []}
    for i, img_bytes in enumerate(images, start=1):
        result = analyze_casting_page(img_bytes, i)
        if result:
            master_json["castings"].append(result)
        time.sleep(2)

    out_file = "all_castings.json"
    with open(out_file, "w", encoding="utf-8") as f:
        json.dump(master_json, f, indent=2, ensure_ascii=False)

    print(f"\n📊 Completed. All castings saved in {out_file}")

if __name__ == "__main__":
    main()
