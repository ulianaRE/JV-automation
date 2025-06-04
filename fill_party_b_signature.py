
"""
fill_party_b_signature.py

This script fills a Word document paragraph with a cleaned-up entity name and
a full U.S. state name, formatted for legal use, based on a JSON input.
"""

import json
import sys
from docx import Document
from docx.shared import Pt
from pathlib import Path

INPUT_DOCX_PATH = "working_agreement.docx"
OUTPUT_DOCX_PATH = "working_agreement.docx"

# Map of U.S. state abbreviations to full names
US_STATE_ABBREVIATIONS = {
    'AL': 'Alabama', 'AK': 'Alaska', 'AZ': 'Arizona', 'AR': 'Arkansas',
    'CA': 'California', 'CO': 'Colorado', 'CT': 'Connecticut', 'DE': 'Delaware',
    'FL': 'Florida', 'GA': 'Georgia', 'HI': 'Hawaii', 'ID': 'Idaho',
    'IL': 'Illinois', 'IN': 'Indiana', 'IA': 'Iowa', 'KS': 'Kansas',
    'KY': 'Kentucky', 'LA': 'Louisiana', 'ME': 'Maine', 'MD': 'Maryland',
    'MA': 'Massachusetts', 'MI': 'Michigan', 'MN': 'Minnesota', 'MS': 'Mississippi',
    'MO': 'Missouri', 'MT': 'Montana', 'NE': 'Nebraska', 'NV': 'Nevada',
    'NH': 'New Hampshire', 'NJ': 'New Jersey', 'NM': 'New Mexico', 'NY': 'New York',
    'NC': 'North Carolina', 'ND': 'North Dakota', 'OH': 'Ohio', 'OK': 'Oklahoma',
    'OR': 'Oregon', 'PA': 'Pennsylvania', 'RI': 'Rhode Island', 'SC': 'South Carolina',
    'SD': 'South Dakota', 'TN': 'Tennessee', 'TX': 'Texas', 'UT': 'Utah',
    'VT': 'Vermont', 'VA': 'Virginia', 'WA': 'Washington', 'WV': 'West Virginia',
    'WI': 'Wisconsin', 'WY': 'Wyoming', 'DC': 'District of Columbia'
}

def get_article(word):
    """Return 'an' if the word starts with a vowel sound, else 'a'."""
    return "an" if word[0].lower() in 'aeiou' else "a"

print()  # Empty line for visual separation
print(f"🔄 Script started: fill_party_b_signature.py")

# Load JSON data
try:
    with open("extracted_values.json", "r") as f:
        data = json.load(f)
    state_abbr = data["funding_partner1_state"]
    raw_entity = data["funding_partner1_entity"]
except KeyError as e:
    print(f"⚠️  Missing required key in JSON: {e}")
    sys.exit(1)
except Exception as e:
    print(f"❌ Error loading JSON: {e}")
    sys.exit(1)

# Convert and validate state
state_full = US_STATE_ABBREVIATIONS.get(state_abbr.upper())
if not state_full:
    print(f"⚠️  Invalid state abbreviation: {state_abbr}")
    sys.exit(1)

# Clean entity name by removing commas
entity_name = raw_entity.replace(",", "")

# Decide on article
article = get_article(state_full)

# Load Word document
doc_path = Path(INPUT_DOCX_PATH)
if not doc_path.exists():
    print("❌ Word document not found.")
    sys.exit(1)

doc = Document(doc_path)

# Find the target paragraph
target_label = "NAME, LLC, "
target_para = None
for para in doc.paragraphs:
    if para.text.strip().startswith(target_label):
        target_para = para
        break

if not target_para:
    print("⚠️  Paragraph starting with label not found.")
    sys.exit(0)

print("✅ Label paragraph found.")

# Extract last run's style
last_run = target_para.runs[-1] if target_para.runs else None
font_name = last_run.font.name if last_run and last_run.font.name else "Calibri"
font_size = last_run.font.size.pt if last_run and last_run.font.size else 11

# Compose new paragraph
new_text = f"{entity_name}, {article} {state_full} limited liability company"

# Replace paragraph content while preserving formatting
target_para.clear()
run = target_para.add_run(new_text)
run.font.name = font_name
run.font.size = Pt(font_size)

# Save document
doc.save(OUTPUT_DOCX_PATH)

# Final debug output
print(f"💾 Inserted value: {new_text}")
print(f"🖋️  Font: {font_name}, Size: {font_size}pt")
print("✅ Script finished successfully.")
