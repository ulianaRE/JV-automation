# run_all.py
# Master runner script for executing modular document automation steps

import subprocess
import os
import sys

def run_script(script_name):
    """Run an individual Python script and report its output."""
    print(f"▶️ Running {script_name}...")
    result = subprocess.run([sys.executable, script_name], capture_output=True, text=True)
    if result.stdout:
        print(result.stdout)
    if result.stderr:
        print(f"⚠️ Error in {script_name}: {result.stderr}")

def main():
    print("\n📄 Starting full document automation process...")

    scripts = [
        "extract_values.py",
        "fill_property.py",
        "fill_party_a_funding.py",
        # Add more scripts here as needed
    ]

    for script in scripts:
        if os.path.exists(script):
            run_script(script)
        else:
            print(f"❌ Script not found: {script}")

    # Cleanup: delete the extracted_values.json file
    json_path = "extracted_values.json"
    if os.path.exists(json_path):
        try:
            os.remove(json_path)
            print(f"🧹 Removed temporary file: {json_path}")
        except Exception as e:
            print(f"⚠️ Could not delete {json_path}: {e}")

    print("\n✅ All scripts executed. Check the generated document(s).")

if __name__ == "__main__":
    main()