import os
import subprocess
import tempfile
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import pdfplumber
import re

def convert_excel_sheets_to_pdf(excel_path):
    """
    Convert specific sheets from Excel files to PDF and merge them
    Sheets must have "PACKING SLIP" as first non-empty words
    """
    
    with tempfile.TemporaryDirectory() as temp_dir:
        pdf_files = []
        
        for filename in os.listdir(excel_path):
            if (filename.lower().endswith((".xls", ".xlsx", ".xlsm")) and 
                "INV" not in filename.upper() and 
                "BCR" not in filename.upper()):
                
                full_input_path = os.path.join(excel_path, filename)
                print(f"🔍 Processing: {filename}")
                
                try:
                    # Convert entire Excel file to PDF
                    subprocess.run([
                        "soffice", 
                        "--headless", 
                        "--convert-to", "pdf", 
                        "--outdir", temp_dir, 
                        full_input_path
                    ], check=True, capture_output=True)
                    
                    converted_pdf = os.path.join(temp_dir, f"{os.path.splitext(filename)[0]}.pdf")
                    
                    if os.path.exists(converted_pdf):
                        # Filter this individual PDF first to keep only packing slip pages
                        filtered_pdf = filter_individual_pdf(converted_pdf, temp_dir)
                        
                        if filtered_pdf:
                            pdf_files.append(filtered_pdf)
                            print(f"✅ Found and filtered packing slip in: {filename}")
                        else:
                            print(f"⚠️ No packing slip pages found in: {filename}")
                            os.remove(converted_pdf)
                    
                except subprocess.CalledProcessError as e:
                    print(f"❌ Failed to convert {filename}: {e}")
        
        # Merge all filtered PDFs
        if pdf_files:
            output_pdf = os.path.join(excel_path, "combined_packing_slips.pdf")
            merger = PdfMerger()
            
            for pdf_file in pdf_files:
                merger.append(pdf_file)
            
            merger.write(output_pdf)
            merger.close()
            
            print(f"📄 Combined {len(pdf_files)} filtered packing slips into: {output_pdf}")
        else:
            print("❌ No packing slips found to combine")

def filter_individual_pdf(input_pdf_path, temp_dir):
    """
    Filter individual PDF to keep only pages with PACKING SLIP
    """
    try:
        with pdfplumber.open(input_pdf_path) as pdf:
            pdf_writer = PdfWriter()
            pages_kept = 0
            
            for page_num, page in enumerate(pdf.pages):
                try:
                    # Extract text with better configuration
                    text = page.extract_text(
                        x_tolerance=1,
                        y_tolerance=1,
                        keep_blank_chars=False
                    )
                    
                    if is_packing_slip_page(text):
                        # Add this page to the output PDF
                        pdf_reader = PdfReader(input_pdf_path)
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                        pages_kept += 1
                        
                except Exception as e:
                    print(f"❌ Error processing page {page_num + 1} in {os.path.basename(input_pdf_path)}: {e}")
                    continue
            
            # Save filtered PDF if we kept any pages
            if pages_kept > 0:
                filtered_pdf_path = os.path.join(temp_dir, f"filtered_{os.path.basename(input_pdf_path)}")
                with open(filtered_pdf_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                return filtered_pdf_path
        
        return None
        
    except Exception as e:
        print(f"❌ Error filtering PDF {os.path.basename(input_pdf_path)}: {e}")
        return None

def is_packing_slip_page(text):
    """
    Check if the text represents a packing slip page
    by looking for 'PACKING SLIP' in the first meaningful content
    """
    if not text:
        return False
    
    # Clean and normalize text
    text = re.sub(r'\s+', ' ', text.upper().strip())
    
    # Split into words and look for "PACKING SLIP" in the first 20 words
    words = text.split()[:20]
    text_start = ' '.join(words)
    
    # Check if "PACKING SLIP" appears in the beginning of the content
    if "PACKING SLIP" in text_start:
        return True
    
    return False

def main():
    excel_path = "/home/pritom/Desktop/Packing List Extraction/Demo"
    
    convert_excel_sheets_to_pdf(excel_path)

if __name__ == "__main__":
    main()