import os
import subprocess

# Path to the source Excel file
excel_path = "/home/pritom/Desktop/Packing List Extraction/Demo"
output_dir = excel_path  # Output PDF will be saved in the same directory

# Loop through all Excel files in the directory
for filename in os.listdir(excel_path):
    if filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
        if "INV" in filename.upper():
            full_input_path = os.path.join(excel_path, filename)

            try:
                # Run LibreOffice in headless mode to convert to PDF
                subprocess.run([
                    "soffice", 
                    "--headless", 
                    "--convert-to", "pdf", 
                    "--outdir", output_dir, 
                    full_input_path
                ], check=True)

                print(f"✅ Converted: {filename}")
            except subprocess.CalledProcessError as e:
                print(f"❌ Failed to convert {filename}: {e}")
        else:
            print(f"⚠️ Skipped (not an invoice): {filename}")