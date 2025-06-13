import shutil
import os
import subprocess
import sys
from pathlib import Path

TEMPLATE = "template.docx"
WORKING = "working_agreement.docx"
OUTPUT = "filled_agreement.docx"

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
    "fill_remedies_on_default.py",
    "fill_party_b_signature.py",
]

def main():
    # Copy template to working file
    shutil.copy(TEMPLATE, WORKING)
    print(f"📝 Working file created: {WORKING}")

    errors = []
    total = len(filler_scripts)

    # Run each filler script
    for i, script in enumerate(filler_scripts, start=1):
        print(f"[{i}/{total}] ▶️ Running {script}")
        result = subprocess.run([sys.executable, script], capture_output=True, text=True)
        if result.returncode != 0:
            print(f"❌ Error in {script}")
            print(result.stdout.strip())
            print(result.stderr.strip())
            errors.append(script)
        else:
            print(f"✅ Completed {script}")

    # Rename working file to output
    if Path(WORKING).exists():
        os.replace(WORKING, OUTPUT)
        print(f"✅ Final agreement saved as: {OUTPUT}")

    # Exit message
    if errors:
        print("\n⚠️ Some scripts had errors:")
        for s in errors:
            print(f" - {s}")
    else:
        print("\n🎉 All scripts ran successfully!")

if __name__ == "__main__":
    main()
