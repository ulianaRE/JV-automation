import shutil
from pathlib import Path
import subprocess
import sys

# === CONSTANTS ===
TEMPLATE_DOCX = "template.docx"
WORKING_DOCX = "working_agreement.docx"
OUTPUT_DOCX = "filled_agreement.docx"

# Step 1: Create a working copy of the template
if Path(WORKING_DOCX).exists():
    Path(WORKING_DOCX).unlink()
shutil.copy(TEMPLATE_DOCX, WORKING_DOCX)

# Step 2: Run each filler script
filler_scripts = [
    "extract_values.py",
    "fill_property.py",
    "fill_lender_name.py",
    "fill_lender_address.py",   
    "fill_lender_email.py", 
    "fill_lender_phone.py",
    "fill_coe_date.py",
    "fill_tiltle_entity.py",
    "fill_title_phone.py",
    "fill_escrow_agent.py",
    "fill_party_a_funding.py",
    "fill_party_b_funding.py",
    "fill_party_b_amount_plus_roi.py",
    "fill_funds_released_at_coe.py",
    "fill_maturity_date.py",
    "fill_party_b_late_fees.py",
    "fill_party_b_signature.py",
    # Add more scripts here
]

for script in filler_scripts:
    result = subprocess.run([sys.executable, script], capture_output=True, text=True)
    if result.returncode != 0:
        print(f"❌ Error running {script}")
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)
        raise RuntimeError(f"{script} failed")

# Step 3: Finalize the document
shutil.copy(WORKING_DOCX, OUTPUT_DOCX)
print(f"✅ All done! Final agreement saved as {OUTPUT_DOCX}")
