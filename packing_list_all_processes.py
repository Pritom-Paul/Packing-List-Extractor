import os
import sys
import subprocess
import importlib.util

def run_script_1(excel_path):
    """
    Run the first script (Excel to PDF conversion for INV files)
    """
    print("=" * 60)
    print("RUNNING SCRIPT 1: Excel to PDF conversion for INV files")
    print("=" * 60)
    
    # Path to the source Excel file
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

def run_script_2(excel_path):
    """
    Run the second script (Packing slip extraction and PDF merging)
    """
    print("\n" + "=" * 60)
    print("RUNNING SCRIPT 2: Packing slip extraction and PDF merging")
    print("=" * 60)
    
    # Import required modules
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
                    "B255" not in filename.upper() and 
                    "CCI" not in filename.upper()):
                    
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

    convert_excel_sheets_to_pdf(excel_path)

def run_script_3(excel_path):
    """
    Run the third script (Excel data extraction and JSON output)
    """
    print("\n" + "=" * 60)
    print("RUNNING SCRIPT 3: Excel data extraction and JSON output")
    print("=" * 60)
    
    # Import required modules
    import xlrd
    import pandas as pd

    def extract_and_print_xls_data(directory):
        # List all .xls files in the directory
        xls_files = [f for f in os.listdir(directory) if f.lower().endswith((".xls", ".xlsx", ".xlsm")) ]

        if not xls_files:
            print("No .xls files found in the directory.")
            return

        # Create a list to store all DataFrames
        all_dfs = []

        for filename in xls_files:
            file_path = os.path.join(directory, filename)
            print(f"\n==== Reading file: {filename} ====")

            try:
                workbook = xlrd.open_workbook(file_path)

                # Flag to track if we've found a packing slip in this file
                packing_slip_found = False
                
                for sheet in workbook.sheets():
                    print(f"\n-- Sheet: {sheet.name} --")
                    
                    # Skip if we already found a packing slip in this file
                    if packing_slip_found:
                        print("Skipping sheet - already found a packing slip in this file")
                        continue
                    
                    # Check if this sheet has 'PACKING SLIP' as the first non-empty cell
                    has_packing_slip = False
                    for row_idx in range(min(5, sheet.nrows)):  # Check first 5 rows
                        row_values = [str(cell.value).strip() for cell in sheet.row(row_idx)]
                        if 'PACKING SLIP' in row_values:
                            has_packing_slip = True
                            break
                    
                    if not has_packing_slip:
                        print("Skipping sheet - 'PACKING SLIP' not found in header")
                        continue
                    
                    # Mark that we found a packing slip in this file
                    packing_slip_found = True
                    
                    # Extract PO number using the specific pattern from your example
                    po_number = None
                    all_rows = []
                    colors_list = []  # List to store colors list from SUB TOTAL rows
                    cartons_list = []  # List to store cartons from SUB TOTAL rows
                    pieces_list = []  # List to store pieces from SUB TOTAL rows
                    total_gross_weight_list = []  # List to store total gross weight from SUB TOTAL rows
                    
                    # Find column indices for cartons, pieces, and total gross weight
                    cartons_col_index = None
                    pieces_col_index = None
                    total_gross_weight_col_index = None
                    
                    for row_idx in range(sheet.nrows):
                        row_values = [cell.value for cell in sheet.row(row_idx)]
                        all_rows.append(row_values)
                        
                        # Look for the specific header row pattern
                        if (len(row_values) > 2 and 
                            'PO' in str(row_values) and 
                            'STYLE' in str(row_values) and 
                            'COLOR' in str(row_values)):
                            
                            # Check next row for the actual data
                            if row_idx + 1 < sheet.nrows:
                                next_row = [cell.value for cell in sheet.row(row_idx + 1)]
                                
                                # Find PO column index
                                po_col_index = None
                                for i, cell_value in enumerate(row_values):
                                    if str(cell_value).strip() == 'PO':
                                        po_col_index = i
                                        break
                                
                                # Extract PO number from next row
                                if po_col_index is not None and po_col_index < len(next_row):
                                    po_candidate = next_row[po_col_index]
                                    if po_candidate and str(po_candidate).strip():
                                        po_number = str(po_candidate).strip()
                                        print(f"*** FOUND PO NUMBER: {po_number} ***")
                        
                        # Find column indices for cartons, pieces, and total gross weight
                        if ('# CARTONS' in str(row_values) and 
                            'TOTAL PIECES' in str(row_values) and 
                            'TOTAL G.W(kg)' in str(row_values)):
                            
                            for i, cell_value in enumerate(row_values):
                                cell_str = str(cell_value).strip()
                                if cell_str == '# CARTONS':
                                    cartons_col_index = i
                                elif cell_str == 'TOTAL PIECES':
                                    pieces_col_index = i
                                elif cell_str == 'TOTAL G.W(kg)':
                                    total_gross_weight_col_index = i
                            
                            print(f"*** FOUND COLUMN INDICES - Cartons: {cartons_col_index}, Pieces: {pieces_col_index}, Total GW: {total_gross_weight_col_index} ***")
                        
                        # Extract data from rows starting with 'SUB TOTAL'
                        if (len(row_values) > 2 and 
                            str(row_values[0]).strip() == 'SUB TOTAL' and 
                            row_values[2] and str(row_values[2]).strip()):
                            
                            color = str(row_values[2]).strip()
                            colors_list.append(color)
                            print(f"*** FOUND COLOR: {color} ***")
                            
                            # Extract cartons
                            if cartons_col_index is not None and cartons_col_index < len(row_values):
                                cartons_value = row_values[cartons_col_index]
                                if cartons_value and str(cartons_value).strip():
                                    # Convert to integer
                                    cartons_int = int(float(cartons_value))
                                    cartons_list.append(cartons_int)
                                    print(f"*** FOUND CARTONS: {cartons_int} ***")
                            
                            # Extract pieces
                            if pieces_col_index is not None and pieces_col_index < len(row_values):
                                pieces_value = row_values[pieces_col_index]
                                if pieces_value and str(pieces_value).strip():
                                    # Convert to integer
                                    pieces_int = int(float(pieces_value))
                                    pieces_list.append(pieces_int)
                                    print(f"*** FOUND PIECES: {pieces_int} ***")
                            
                            # Extract total gross weight
                            if total_gross_weight_col_index is not None and total_gross_weight_col_index < len(row_values):
                                total_gross_weight_value = row_values[total_gross_weight_col_index]
                                if total_gross_weight_value and str(total_gross_weight_value).strip():
                                    # Format to 3 decimal places
                                    formatted_weight = f"{float(total_gross_weight_value):.3f}"
                                    total_gross_weight_list.append(formatted_weight)
                                    print(f"*** FOUND TOTAL GROSS WEIGHT: {formatted_weight} ***")
                    
                    # Print all rows
                    for row in all_rows:
                        print(row)
                    
                    if po_number:
                        print(f"\n*** EXTRACTED PO NUMBER: {po_number} ***")
                    else:
                        print("\n*** PO NUMBER NOT FOUND ***")
                    
                    # Print extracted colors_list
                    if colors_list:
                        colors_str = ', '.join(colors_list)
                        print(f"*** EXTRACTED COLORS: [{colors_str}] ***")
                    else:
                        print("*** NO COLORS FOUND ***")
                    
                    # Print extracted cartons
                    if cartons_list:
                        cartons_str = ', '.join([str(c) for c in cartons_list])
                        print(f"*** EXTRACTED CARTONS: [{cartons_str}] ***")
                    else:
                        print("*** NO CARTONS FOUND ***")
                    
                    # Print extracted pieces
                    if pieces_list:
                        pieces_str = ', '.join([str(p) for p in pieces_list])
                        print(f"*** EXTRACTED PIECES: [{pieces_str}] ***")
                    else:
                        print("*** NO PIECES FOUND ***")
                    
                    # Print extracted total gross weight
                    if total_gross_weight_list:
                        total_gross_weight_str = ', '.join(total_gross_weight_list)
                        print(f"*** EXTRACTED TOTAL GROSS WEIGHT: [{total_gross_weight_str}] ***")
                    else:
                        print("*** NO TOTAL GROSS WEIGHT FOUND ***")

                    # Create DataFrame for this packing list sheet (ONE ROW PER FILE)
                    if po_number:  # Only create DataFrame if we found a PO number
                        df_data = {
                            'PO_Number': [po_number],
                            'Colors': [colors_list],
                            'Cartons': [cartons_list],
                            'Pieces': [pieces_list],
                            'Total_Gross_Weight': [total_gross_weight_list]
                        }
                        
                        df = pd.DataFrame(df_data)
                        all_dfs.append(df)
                        print(f"\n*** CREATED DATAFRAME FOR {filename} - {sheet.name} ***")
                        print(df)

            except Exception as e:
                print(f"Error reading '{filename}': {e}")

        # Combine all DataFrames into one master DataFrame
        if all_dfs:
            master_df = pd.concat(all_dfs, ignore_index=True)
            print(f"\n{'='*50}")
            print("MASTER DATAFRAME SUMMARY:")
            print(f"{'='*50}")
            print(f"Total files processed: {len(xls_files)}")
            print(f"Total packing lists found: {len(all_dfs)}")
            print(f"Master DataFrame shape: {master_df.shape}")
            print(f"\nMaster DataFrame:")
            print(master_df)

            # Export to JSON and print
            json_output = master_df.to_json(orient='records', indent=2)
            print(json_output)
            
            # # Save to CSV
            # output_csv = os.path.join(directory, "packing_lists_summary.csv")
            # master_df.to_csv(output_csv, index=False)
            # print(f"\n*** Master DataFrame saved to: {output_csv} ***")
        else:
            print("\n*** No packing list data found to create DataFrame ***")

    extract_and_print_xls_data(excel_path)

def main():
    """
    Main function to run all three scripts sequentially
    """
    # Set your input directory here
    excel_path = "/home/pritom/Desktop/Packing List Extraction/Demo"
    
    print("🚀 STARTING ALL SCRIPTS")
    print(f"📁 Input Directory: {excel_path}")
    
    try:
        # Run Script 1
        run_script_1(excel_path)
        
        # Run Script 2  
        run_script_2(excel_path)
        
        # Run Script 3
        run_script_3(excel_path)
        
        print("\n" + "=" * 60)
        print("✅ ALL SCRIPTS COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()