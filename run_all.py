import shutil
from pathlib import Path
import subprocess
import sys

# === CONSTANTS ===
TEMPLATE_DOCX = "template.docx"
WORKING_DOCX = "working_agreement.docx"
OUTPUT_DOCX = "filled_agreement.docx"
LOG_FILE = "run_all.log"  # Path for execution log file

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
    "fill_remedies_on_default.py",
    "fill_party_b_signature.py",
    # Add more scripts here
]

errors = []

# Clear previous log file and add header
with open(LOG_FILE, "w") as log:
    log.write("📄 JV Automation Execution Log\n\n")

for i, script in enumerate(filler_scripts, start=1):
    print(f"[{i}/{len(filler_scripts)}] ▶️ Running {script} ...")
    result = subprocess.run([sys.executable, script], capture_output=True, text=True)

    # Write output to log
    with open(LOG_FILE, "a") as log:
        log.write(f"▶️ Running {script}...\n")
        log.write(result.stdout)
        log.write(result.stderr)

        if result.returncode != 0:
            print(f"❌ Error in {script}")
            print("STDOUT:", result.stdout.strip())
            print("STDERR:", result.stderr.strip())
            log.write(f"❌ Error in {script}\n\n")
            errors.append(script)
        else:
            print(f"✅ {script} completed")
            log.write(f"✅ {script} completed successfully\n\n")

# Step 3: Finalize the document
shutil.copy(WORKING_DOCX, OUTPUT_DOCX)
print(f"\n✅ Final agreement saved as {OUTPUT_DOCX}")

# Step 4: Report any errors
with open(LOG_FILE, "a") as log:
    if errors:
        print("\n⚠️ The following scripts had errors and may need review:")
        log.write("\n⚠️ The following scripts failed:\n")
        for err in errors:
            print(f" - {err}")
            log.write(f" - {err}\n")
    else:
        print("\n🎉 All scripts ran successfully!")
        log.write("\n🎉 All scripts ran successfully!\n")
