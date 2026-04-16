import pandas as pd
import re
import os

INPUT_FILE = r"C:\Users\Admin\Desktop\SHEFI_PO_DHAVAL\Book2.xlsx"

FINAL_COLUMNS = [
    "SrNo",
    "StyleCode",
    "ItemSize",
    "OrderQty",
    "OrderItemPcs",
    "Metal",
    "Tone",
    "ItemPoNo",
    "ItemRefNo",
    "StockType",
    "Priority",
    "MakeType",
    "CustomerProductionInstruction",
    "SpecialRemarks",
    "DesignProductionInstruction",
    "StampInstruction",
    "OrderGroup",
    "Certificate",
    "SKUNo",
    "Basestoneminwt",
    "Basestonemaxwt",
    "Basemetalminwt",
    "Basemetalmaxwt",
    "Productiondeliverydate",
    "Expecteddeliverydate",
    "SetPrice",
    "StoneQuality",
]


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def map_style_and_size(style_code, original_size):
    """
    Drop the last '-' segment from StyleCode entirely.
    ItemSize = last segment with trailing '0' stripped.
    e.g. 'FR1500JE-SM-US070' → StyleCode='FR1500JE-SM', ItemSize='US07'
    """
    if not isinstance(style_code, str):
        return style_code, original_size

    parts = style_code.split("-")
    if len(parts) > 1:
        size_raw = parts[-1]                          # e.g. 'US070'
        base     = "-".join(parts[:-1])               # e.g. 'FR1500JE-SM'
    else:
        size_raw = parts[0]
        base     = parts[0]

    # Strip trailing '0' from size (e.g. 'US070' → 'US07')
    size = size_raw[:-1] if size_raw.endswith("0") and len(size_raw) > 1 else size_raw
    return base, size


def parse_metal_tone(mitemcode):
    """
    Parse MItemCode into (Metal, Tone).
    Examples:
        'AG925'     → ('AG925', '')
        'PC95'      → ('PC95',  '')
        'G14W'      → ('G14',   'W')
        'G10W/P/Y'  → ('G10',   'W/P/Y')
        'G18Y'      → ('G18',   'Y')
    Rule: leading letters+digits form the Metal; remaining letters/slashes form Tone.
    """
    if not isinstance(mitemcode, str) or not mitemcode.strip():
        return ("", "")

    mitemcode = mitemcode.strip()
    m = re.match(r'^([A-Z]+\d+)([A-Z/]*)$', mitemcode, re.IGNORECASE)
    if m:
        metal = m.group(1).upper()
        tone  = m.group(2).upper().strip("/")
        return (metal, tone)

    return (mitemcode, "")


def extract_sku(special_remarks):
    """
    Extract 'SKU#XXXXXXX' from SpecialRemarks if present, else return ''.
    Stops at comma, space, or end of string.
    """
    if not isinstance(special_remarks, str):
        return ""
    m = re.search(r'SKU#[^,\s]+', special_remarks, re.IGNORECASE)
    return m.group(0).strip() if m else ""


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    if not os.path.exists(INPUT_FILE):
        print(f"Error: '{INPUT_FILE}' not found in the current directory.")
        return

    # User inputs
    item_po_no = input("Enter ItemPoNo : ").strip()
    priority   = input("Enter Priority  : ").strip()

    print(f"\nReading '{INPUT_FILE}' ...")
    df = pd.read_excel(INPUT_FILE)

    # Keep only the first row of each StyleCode group (where original SrNo == 1)
    df = df[df["SrNo"] == 1].copy().reset_index(drop=True)
    print(f"  {len(df)} StyleCode group(s) found (rows where SrNo = 1)")

    out = pd.DataFrame()

    # SrNo — sequential starting from 1 for output
    out["SrNo"] = range(1, len(df) + 1)

    # StyleCode + ItemSize (derived together)
    style_mapped, size_mapped = zip(
        *df.apply(
            lambda r: map_style_and_size(r.get("StyleCode"), r.get("ItemSize")),
            axis=1,
        )
    )
    out["StyleCode"] = list(style_mapped)
    out["ItemSize"]  = list(size_mapped)

    # Quantities (renamed)
    out["OrderQty"]      = df.get("InwardQty", "")
    out["OrderItemPcs"]  = df.get("ItemPcs", "")

    # Metal + Tone from MItemCode
    metal_tone = df["MItemCode"].apply(parse_metal_tone)
    out["Metal"] = metal_tone.apply(lambda x: x[0])
    out["Tone"]  = metal_tone.apply(lambda x: x[1])

    # ItemPoNo — user input applied to all rows
    out["ItemPoNo"] = item_po_no

    # ItemRefNo — not in source, leave blank
    out["ItemRefNo"] = ""

    out["StockType"] = df.get("StockType", "")
    out["Priority"]  = priority
    out["MakeType"]  = df.get("MakeType", "")
    out["CustomerProductionInstruction"] = df.get("CustomerProductionInstruction", "")
    out["SpecialRemarks"]                = df.get("SpecialRemarks", "")
    out["DesignProductionInstruction"]   = df.get("DesignProductionInstruction", "")
    out["StampInstruction"]              = df.get("StampingInstruction", "")

    # Fixed value
    out["OrderGroup"] = "STERLING JEWELERS(OUTLET)"

    # Blank columns
    out["Certificate"] = ""

    # SKUNo extracted from SpecialRemarks
    out["SKUNo"] = df["SpecialRemarks"].apply(extract_sku)

    for col in ["Basestoneminwt", "Basestonemaxwt", "Basemetalminwt", "Basemetalmaxwt",
                "Productiondeliverydate", "Expecteddeliverydate", "SetPrice", "StoneQuality"]:
        out[col] = ""

    # Enforce final column order
    out = out[FINAL_COLUMNS]

    output_file = os.path.splitext(INPUT_FILE)[0] + "_Output.xlsx"
    out.to_excel(output_file, index=False)

    print(f"\nDone! {len(out)} rows × {len(out.columns)} columns saved to '{output_file}'")


if __name__ == "__main__":
    main()
