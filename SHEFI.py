# SHEFI_BULK.py - Bulk processing multiple SHEFI files
import pandas as pd
import re
import os
import glob
from pathlib import Path

def process_shefi_file(input_file_path, output_folder=None):
    """
    Process single SHEFI Excel file and convert to GATI format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output Excel file (optional)
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read PO# from A2 ---
        po_value = pd.read_excel(
            input_file_path,
            header=None, engine='openpyxl'
        ).iloc[1, 0]

        # --- Step 2: Read data starting row 11 ---
        df = pd.read_excel(
            input_file_path,
            skiprows=10
        )

        # --- Step 3: Select required columns ---
        selected_columns = [
            'VendorStyle#', 'QTY', 'MetalType', 'Color', 'PD#',
            'Description', 'Shefi#', 'SHEFIPO#', 'CODE'
        ]
        
        # Check if all required columns exist
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            return False, None, f"Missing columns: {', '.join(missing_columns)}", None
        
        df_selected = df[selected_columns].copy()

        # --- Step 4: Drop fully blank rows ---
        df_selected.dropna(how='all', inplace=True)

        # --- Step 5: Keep only rows where StyleCode or ItemRefNo exists ---
        df_selected = df_selected[df_selected['VendorStyle#'].notna() | df_selected['PD#'].notna()].copy()

        # --- Step 6: Clean Description ---
        df_selected['Description'] = df_selected['Description'].astype(str).str.replace('\n', ' ', regex=True)

        # --- Step 7: Rename columns ---
        df_selected.rename(columns={
            'VendorStyle#': 'StyleCode',
            'QTY': 'OrderQty',
            'MetalType': 'MetalType',
            'Color': 'Tone',
            'PD#': 'ItemRefNo',
            'Description': 'CustomerProductionInstruction',
            'Shefi#': 'SKUNo',
            'SHEFIPO#': 'SHEFIPO#',
            'CODE': 'DIA GRADE'
        }, inplace=True)

        # --- Step 8: Clean and standardize StyleCode ---
        def clean_stylecode(value, item_ref):
            val = str(value).strip().upper()
            if val in ["N.A.", "NAN", "", "NONE"]:
                return str(item_ref).strip()
            return val.split('-')[0] if '-' in val else val

        df_selected['StyleCode'] = df_selected.apply(
            lambda r: clean_stylecode(r['StyleCode'], r['ItemRefNo']), axis=1
        )

        # --- Step 9: Drop rows where StyleCode or MetalType missing ---
        df_cleaned = df_selected.dropna(subset=['StyleCode', 'MetalType']).copy()

        # --- Step 10: Add SrNo ---
        df_cleaned.insert(0, 'SrNo', range(1, len(df_cleaned) + 1))

        # --- Step 11: Add ItemSize ---
        df_cleaned.insert(2, 'ItemSize', '')

        # --- Step 12: Add OrderItemPcs ---
        order_qty_index = df_cleaned.columns.get_loc('OrderQty')
        df_cleaned.insert(order_qty_index + 1, 'OrderItemPcs', 1)

        # --- Step 13: Add StockType & MakeType ---
        item_ref_index = df_cleaned.columns.get_loc('ItemRefNo')
        df_cleaned.insert(item_ref_index + 1, 'StockType', '')
        df_cleaned.insert(item_ref_index + 2, 'MakeType', '')

        # --- Step 14: Add OrderGroup & Certificate before SKUNo ---
        sku_index = df_cleaned.columns.get_loc('SKUNo')
        df_cleaned.insert(sku_index, 'OrderGroup', 'SHEFI')
        df_cleaned.insert(sku_index + 1, 'Certificate', '')

        # --- Step 15: Add additional columns after SKUNo ---
        sku_index = df_cleaned.columns.get_loc('SKUNo')
        new_columns = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for i, col in enumerate(new_columns):
            df_cleaned.insert(sku_index + 1 + i, col, '')

        # --- Step 16: Add ItemPoNo ---
        tone_index = df_cleaned.columns.get_loc('Tone')
        df_cleaned.insert(tone_index + 1, 'ItemPoNo', po_value)

        # --- Step 17: Create Metal column ---
        def create_metal(metaltype, tone):
            num = re.sub(r'\D', '', str(metaltype))
            return 'G' + num + str(tone) if num else 'G' + str(tone)

        df_cleaned['Metal'] = df_cleaned.apply(lambda r: create_metal(r['MetalType'], r['Tone']), axis=1)

        # Replace MetalType with Metal
        metal_type_index = df_cleaned.columns.get_loc('MetalType')
        df_cleaned.drop(columns=['MetalType'], inplace=True)
        metal_col = df_cleaned.pop('Metal')
        df_cleaned.insert(metal_type_index, 'Metal', metal_col)

        # --- Step 18: Add SpecialRemarks ---
        df_cleaned['SpecialRemarks'] = df_cleaned.apply(
            lambda r: f"PD#, {r['ItemRefNo']}, SHEFI # {r['SKUNo']}, SHEFI PO# ,{r['SHEFIPO#']} ,{r['Metal']}, DIA QLTY {r['DIA GRADE']}",
            axis=1
        )

        # --- Step 19: Add DesignProductionInstruction & StampInstruction ---
        dpi_index = df_cleaned.columns.get_loc('CustomerProductionInstruction')
        df_cleaned.insert(dpi_index + 1, 'SpecialRemarks', df_cleaned.pop('SpecialRemarks'))
        df_cleaned.insert(dpi_index + 2, 'DesignProductionInstruction', '')

        # --- Step 20: StampInstruction logic ---
        def get_stamp_instruction(metal):
            if metal.startswith("G14"):
                return "14K & DP2 LOGO"
            elif metal.startswith("G10"):
                return "10K & DP2 LOGO"
            elif metal.startswith("G18"):
                return "18K & DP2 LOGO"
            elif metal == "PC95":
                return "PT950 & DP2 LOGO"
            elif metal == "A4YUP342-":
                return "ALLOY & DP2 LOGO"
            elif metal == "AG925":
                return "KT & DP2 LOGO"
            else:
                return "0 & DP2 LOGO"

        df_cleaned['StampInstruction'] = df_cleaned['Metal'].apply(get_stamp_instruction)
        df_cleaned.insert(dpi_index + 3, 'StampInstruction', df_cleaned.pop('StampInstruction'))

        # AG925 rule: Tone is blank for silver metal
        df_cleaned.loc[df_cleaned['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        # --- Step 21: Remove any residual NaN (fill blanks) ---
        df_cleaned = df_cleaned.fillna('')

        # --- Step 22: Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"GATI_FORMAT_{input_filename}.xlsx")
        else:
            output_path = f"GATI_FORMAT_{input_filename}.xlsx"
        
        # --- Step 23: Export ---
        df_cleaned.to_excel(output_path, index=False)
        
        return True, output_path, None, df_cleaned
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output Excel files (optional)
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_shefi_file(file_path, output_folder)
        
        results.append({
            'input_file': os.path.basename(file_path),
            'success': success,
            'output_file': os.path.basename(output_path) if output_path else None,
            'output_path': output_path,
            'error': error,
            'row_count': len(df) if df is not None else 0
        })
    
    return results

def main():
    """Main function for command line usage"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Bulk process SHEFI Excel files to GATI format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        results = process_multiple_files(args.input, args.output)
        
        print("\n" + "="*50)
        print("BATCH PROCESSING RESULTS")
        print("="*50)
        
        success_count = 0
        for result in results:
            status = "✅ SUCCESS" if result['success'] else "❌ FAILED"
            print(f"{status}: {result['input_file']} -> {result['output_file'] or 'N/A'}")
            print(f"   Rows: {result['row_count']}")
            if result['error']:
                print(f"   Error: {result['error']}")
            print()
            
            if result['success']:
                success_count += 1
        
        print(f"Processed: {len(results)} files | Successful: {success_count} | Failed: {len(results) - success_count}")
        
    else:
        # Process single file
        if not os.path.isfile(args.input):
            print("❌ Input path must be a file when not using --batch")
            return
        
        success, output_path, error, df = process_shefi_file(args.input, args.output)
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print("\nFirst 5 rows preview:")
            print(df.head())
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()