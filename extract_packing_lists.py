import xlrd # type: ignore
import os
import pandas as pd # type: ignore

def extract_and_print_xls_data(directory):
    # List all .xls files in the directory
    xls_files = [f for f in os.listdir(directory) if f.lower().endswith('.xls')]

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

if __name__ == "__main__":
    directory_path = "/media/pritom/Products/New/Packing List Extraction/Demo"
    extract_and_print_xls_data(directory_path)