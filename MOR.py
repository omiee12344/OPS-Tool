# MOR_EXCEL_BULK.py - Bulk processing multiple MOR Excel files
import pandas as pd
import re
import os
import glob
from pathlib import Path


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('RG0002939A', 'EU56', 'W') -> 'RG0002939A-56WG'
    """
    base = str(base).strip() if base else ''
    item_size = str(item_size).strip() if item_size else ''
    tone = str(tone).strip().upper() if tone else ''
    if not base or base.upper() == 'NAN':
        return ''
    # Detect INCH before stripping — insert IN in suffix (e.g. 7 INCH+W -> 7INWG)
    has_inch = bool(re.search(r'\bINCH\b', item_size, flags=re.IGNORECASE))
    size_num = re.sub(r'^(?:UP|US|EU|IT|UT|TS|IS)\s*', '', item_size, flags=re.IGNORECASE).strip()
    size_num = re.sub(r'\s*INCH\s*$', '', size_num, flags=re.IGNORECASE).strip()
    try:
        f = float(size_num)
        size_num = str(int(f)) if f.is_integer() else str(f)
    except (ValueError, TypeError):
        pass
    # Normalize multi-char tones: 'YV' -> 'Y', 'WG' -> 'W', etc. Keep 'PT' as-is
    if tone and len(tone) > 1 and tone != 'PT':
        first = tone[0]
        if first in ('W', 'Y', 'P', 'R'):
            tone = first
    in_part = 'IN' if has_inch else ''
    if tone == 'PT':
        suffix = f"{size_num}{in_part}PT" if size_num else (f"{in_part}PT" if in_part else 'PT')
    elif tone in ('W', 'Y', 'P', 'R'):
        suffix = f"{size_num}{in_part}{tone}G" if size_num else f"{in_part}{tone}G"
    elif tone == 'AG':
        suffix = f"{size_num}{in_part}AG" if size_num else (f"{in_part}AG" if in_part else 'AG')
    else:
        suffix = size_num
    return f"{base}-{suffix}" if suffix else base


_ITEM_SIZE_LOOKUP = None


def _get_size_lookup():
    global _ITEM_SIZE_LOOKUP
    if _ITEM_SIZE_LOOKUP is not None:
        return _ITEM_SIZE_LOOKUP
    _ITEM_SIZE_LOOKUP = {}
    try:
        mst = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ItemSize_Mst.xlsx')
        _df_mst = pd.read_excel(mst)
        for val in _df_mst['Item Size Code'].dropna():
            vs = str(val).strip()
            if vs and vs.upper() != 'NAN':
                k = _normalize_size_key(vs)
                if k:
                    _ITEM_SIZE_LOOKUP[k] = vs
    except Exception:
        pass
    return _ITEM_SIZE_LOOKUP


def _normalize_size_key(s):
    s = str(s).strip()
    if not s or s.upper() == 'NAN':
        return ''
    m = re.match(r'^(\d+(?:\.\d+)?)\s*INCH$', s, re.IGNORECASE)
    if m:
        return f"{float(m.group(1)):.2f}inch"
    return re.sub(r'\s+', '', s).lower()


def _map_item_size(raw):
    """Map raw ItemSize to its canonical form from ItemSize_Mst.xlsx."""
    if not raw or str(raw).strip().upper() in ('', 'NAN'):
        return raw
    lookup = _get_size_lookup()
    key = _normalize_size_key(str(raw).strip())
    return lookup.get(key, raw)

def process_mor_file(input_file_path, output_folder=None, item_po_no=None, priority_value=None):
    """
    Process single MOR Excel file and convert to standardized format
    
    Parameters:
    input_file_path (str): Path to input Excel file
    output_folder (str): Folder for output Excel file (optional)
    item_po_no (str): ItemPoNo value
    priority_value (str): Priority value
    
    Returns:
    tuple: (success_status, output_path, error_message, dataframe)
    """
    try:
        # --- Step 1: Read Excel and skip top 4 rows ---
        df = pd.read_excel(input_file_path, skiprows=4)

        # --- Step 2: Select relevant columns ---
        selected_columns = ['SAP CODE', 'Shimayra #', 'LABEL DESCRIPTION', 'Diamond Quality', 'ORDER QTY']
        
        # Check if all required columns exist
        missing_columns = [col for col in selected_columns if col not in df.columns]
        if missing_columns:
            return False, None, f"Missing columns: {', '.join(missing_columns)}", None
        
        df = df[selected_columns].copy()

        # --- Step 3: Rename columns ---
        df.rename(columns={
            'SAP CODE': 'SKUNo',
            'Shimayra #': 'StyleCode',
            'LABEL DESCRIPTION': 'CustomerProductionInstruction',
            'ORDER QTY': 'OrderQty'
        }, inplace=True)

        # --- Step 4: Extract StyleCode and ItemSize ---
        def extract_style_and_size(style_text):
            """
            Extract StyleCode (before '-') and ItemSize (EU + number) if present.
            Example:
                RG0002939A-EU58WG → StyleCode: RG0002939A, ItemSize: EU58
                RG0002939A → StyleCode: RG0002939A, ItemSize: ""
            """
            if pd.isna(style_text):
                return "", ""
            
            style_text = str(style_text).strip()
            
            # Try pattern like "RG0002939A-EU58WG"
            match = re.match(r'^([A-Za-z0-9]+)-?(EU\d+)?', style_text)
            if match:
                style_code = match.group(1).strip()
                item_size = match.group(2).strip() if match.group(2) else ""
            else:
                style_code = style_text.split('-')[0].strip()
                item_size = ""
            return style_code, item_size

        # Apply extraction
        df['StyleCode'], df['ItemSize'] = zip(*df['StyleCode'].map(extract_style_and_size))

        # --- Helper Functions ---
        def map_metal_from_instruction(instruction):
            """Map metal type from CustomerProductionInstruction"""
            if pd.isna(instruction):
                return ""
            instruction = str(instruction).upper()
            if "OR585 WHITE GOLD" in instruction:
                return "G585W"
            elif "OR585 YELLOW GOLD" in instruction:
                return "G585Y"
            elif "OR585 PINK GOLD" in instruction:
                return "G585P"
            elif "OR587 WHITE GOLD" in instruction:
                return "G587W"
            elif "OR587 YELLOW GOLD" in instruction:
                return "G587Y"
            elif "OR587 PINK GOLD" in instruction:
                return "G587P"
            elif "OR750 WHITE GOLD" in instruction:
                return "G750W"
            elif "OR750 YELLOW GOLD" in instruction:
                return "G750Y"
            elif "OR750 PINK GOLD" in instruction:
                return "G750P"
            elif "PT" in instruction or "PLATINUM" in instruction:
                return "PC95"
            else:
                return ""

        def extract_tone_from_metal(metal):
            """Extract tone from metal code"""
            if pd.isna(metal) or metal == "":
                return ""
            metal = str(metal)
            if metal.endswith("W"):
                return "W"
            elif metal.endswith("Y"):
                return "Y"
            elif metal.endswith("P"):
                return "P"
            elif metal == "PC95":
                return "PT"
            else:
                return ""

        def create_special_remarks(row):
            """Create SpecialRemarks from SKUNo, Metal, Tone, and Diamond Quality"""
            sku = str(row.get('SKUNo', "")) if pd.notna(row.get('SKUNo', "")) else ""
            metal = str(row.get('Metal', "")) if pd.notna(row.get('Metal', "")) else ""
            tone = str(row.get('Tone', "")) if pd.notna(row.get('Tone', "")) else ""
            diamond_quality = str(row.get('Diamond Quality', "")) if pd.notna(row.get('Diamond Quality', "")) else ""
            metal_num = ""
            if metal:
                numbers = re.findall(r'\d+', metal)
                if numbers:
                    metal_num = numbers[0]
            tone_map = {
                "W": "WHITE GOLD",
                "Y": "YELLOW GOLD",
                "P": "PINK GOLD",
                "PT": "PLATINUM"
            }
            tone_full = tone_map.get(tone, tone)
            parts = []
            if sku:
                parts.append(f"S.{sku}")
            if metal_num:
                parts.append(metal_num)
            if tone_full:
                parts.append(tone_full)
            if diamond_quality:
                parts.append(f"DIA QLTY: {diamond_quality}")
            return ", ".join(parts)

        def get_design_production_instruction(tone):
            """Get design production instruction based on tone"""
            if pd.isna(tone) or tone == "":
                return ""
            tone = str(tone)
            if tone == "W":
                return "WHITE RODIUM"
            else:
                return "NO RODIUM"

        def get_stamp_instruction(metal):
            """Get stamp instruction based on metal"""
            if pd.isna(metal) or metal == "":
                return ""
            metal = str(metal)
            numbers = re.findall(r'\d+', metal)
            if numbers:
                metal_num = numbers[0]
                return f"GOLD TITLE ({metal_num} with frame) + SJ LOGO + CHR with \"ROUND FRAME\" LOGO"
            else:
                return "GOLD TITLE + SJ LOGO + CHR with \"ROUND FRAME\" LOGO"

        # --- Get user inputs if not provided ---
        if item_po_no is None:
            item_po_no = input(f"Enter ItemPoNo for {Path(input_file_path).name}: ").strip()
        
        if priority_value is None:
            priority_value = input(f"Enter Priority for {Path(input_file_path).name}: ").strip()

        # --- Create the final DataFrame ---
        result_df = pd.DataFrame()
        result_df['SrNo'] = range(1, len(df) + 1)
        result_df['StyleCode'] = df['StyleCode']
        result_df['ItemSize'] = df['ItemSize']
        result_df['OrderQty'] = df['OrderQty']
        result_df['OrderItemPcs'] = 1
        result_df['Metal'] = df['CustomerProductionInstruction'].apply(map_metal_from_instruction)
        result_df['Tone'] = result_df['Metal'].apply(extract_tone_from_metal)
        result_df['ItemPoNo'] = item_po_no
        result_df['ItemRefNo'] = ""
        result_df['StockType'] = ""
        result_df['Priority'] = priority_value
        result_df['MakeType'] = ""
        result_df['CustomerProductionInstruction'] = df['CustomerProductionInstruction']
        result_df['SKUNo'] = df['SKUNo']
        result_df['Diamond Quality'] = df['Diamond Quality']
        result_df['SpecialRemarks'] = result_df.apply(create_special_remarks, axis=1)
        result_df['DesignProductionInstruction'] = result_df['Tone'].apply(get_design_production_instruction)
        result_df['StampInstruction'] = result_df['Metal'].apply(get_stamp_instruction)
        result_df['OrderGroup'] = ""
        result_df['Certificate'] = ""
        result_df['ItemSize'] = result_df['ItemSize'].apply(_map_item_size)
        result_df['StyleCode'] = result_df.apply(
            lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
        )
        result_df.loc[result_df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''

        blank_columns = [
            'Basestoneminwt', 'Basestonemaxwt', 'Basemetalminwt', 'Basemetalmaxwt',
            'Productiondeliverydate', 'Expecteddeliverydate', 'SetPrice', 'StoneQuality'
        ]
        for col in blank_columns:
            result_df[col] = ""

        # --- Generate output filename ---
        input_filename = Path(input_file_path).stem
        if output_folder:
            output_path = os.path.join(output_folder, f"MOR_FORMAT_{input_filename}.xlsx")
        else:
            output_path = f"MOR_FORMAT_{input_filename}.xlsx"
        
        # --- Save to Excel ---
        result_df.to_excel(output_path, index=False)
        
        return True, output_path, None, result_df
        
    except Exception as e:
        return False, None, str(e), None

def process_multiple_files(input_folder, output_folder=None, item_po_no=None, priority_value=None):
    """
    Process all Excel files in a folder
    
    Parameters:
    input_folder (str): Folder containing input Excel files
    output_folder (str): Folder for output Excel files (optional)
    item_po_no (str): Common ItemPoNo for all files
    priority_value (str): Common Priority value for all files
    
    Returns:
    list: Results for each file processed
    """
    if output_folder and not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel_files = glob.glob(os.path.join(input_folder, "*.xlsx")) + glob.glob(os.path.join(input_folder, "*.xls"))
    results = []
    
    for file_path in excel_files:
        print(f"Processing: {os.path.basename(file_path)}")
        success, output_path, error, df = process_mor_file(
            file_path, output_folder, item_po_no, priority_value
        )
        
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
    
    parser = argparse.ArgumentParser(description='Bulk process MOR Excel files to standardized format')
    parser.add_argument('--input', '-i', required=True, help='Input file or folder path')
    parser.add_argument('--output', '-o', help='Output folder path (optional)')
    parser.add_argument('--batch', '-b', action='store_true', help='Process all files in input folder')
    parser.add_argument('--item-po-no', '-p', help='ItemPoNo value')
    parser.add_argument('--priority', '-r', help='Priority value')
    
    args = parser.parse_args()
    
    if args.batch:
        # Process all files in folder
        if not os.path.isdir(args.input):
            print("❌ Input path must be a folder when using --batch")
            return
        
        # Get common values if not provided
        if not args.item_po_no:
            args.item_po_no = input("Enter ItemPoNo value for all files: ").strip()
        
        if not args.priority:
            args.priority = input("Enter Priority value for all files: ").strip()
        
        results = process_multiple_files(args.input, args.output, args.item_po_no, args.priority)
        
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
        
        success, output_path, error, df = process_mor_file(args.input, args.output, args.item_po_no, args.priority)
        
        if success:
            print(f"✅ SUCCESS: Processed {args.input}")
            print(f"📊 Output: {output_path}")
            print(f"📋 Rows processed: {len(df)}")
            print(f"🏷️ ItemPoNo: {df['ItemPoNo'].iloc[0]}")
            print(f"⏱️ Priority: {df['Priority'].iloc[0]}")
            print("\nSample Preview (StyleCode, ItemSize, Tone, Metal):")
            print(df[['SrNo', 'StyleCode', 'ItemSize', 'Tone', 'Metal']].head(10).to_string(index=False))
        else:
            print(f"❌ FAILED: {error}")

if __name__ == "__main__":
    main()