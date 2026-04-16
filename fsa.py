import os
import re
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'EU58', 'W') -> 'VR1943EEA-58WG'
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


def extract_full_text_from_pdf(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text:
                full_text += text + "\n"
    return full_text


def find_po_number(text: str) -> str:
    m = re.search(r"Order\s+(\d+)", text)
    return m.group(1) if m else ""


# Patterns and parser ported from notebook (non-interactive)
HEADER_PAT_A = re.compile(r"^([A-Z0-9\-]{3,})\s+(?:([YWR]G)\s*?750|([YWR]G750))?\s*(\d{8})\s+(STA|\d+)\s+(\d+)\s+(\d+)\s+Piece\b")
HEADER_PAT_B = re.compile(r"^([A-Z][A-Z0-9\-]{2,})\s+(STA|\d+)\s+(\d+)\s+(\d+)\s+Piece\b")
HEADER_PAT_C = re.compile(r"^(\d{8})\s+(STA|\d+)\s+(\d+)\s+(\d+)\s+Piece\b")
TOTAL_SKU_PAT = re.compile(r"^TOTAL\s+(\d{8})\b")

STYLE_TOKEN_PAT = re.compile(r"\b([A-Z][A-Z\-]*\d{3,})\b")
EXCLUDE_STYLE_TOKENS = {"YG750", "WG750", "RG750", "YG", "WG", "RG"}


def is_item_header_v2(line: str) -> bool:
    return bool(HEADER_PAT_A.search(line) or HEADER_PAT_B.search(line) or HEADER_PAT_C.search(line))


def find_style_in_block(block_lines: list[str]) -> str:
    for ln in block_lines[1:]:
        if ln.upper().startswith("TOTAL "):
            break
        m = STYLE_TOKEN_PAT.search(ln)
        if m:
            token = m.group(1)
            if token not in EXCLUDE_STYLE_TOKENS and not token.isdigit() and len(token) >= 5:
                return token
    return ""


def parse_items_v2(text: str, default_priority: str = "REG", default_stamp_var: str = "") -> list[dict]:
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    items: list[dict] = []
    i = 0
    while i < len(lines):
        line = lines[i]
        if not is_item_header_v2(line):
            i += 1
            continue

        style_code = ""
        sku_no = ""
        size_token = ""
        order_qty = ""

        mA = HEADER_PAT_A.search(line)
        mB = HEADER_PAT_B.search(line) if not mA else None
        mC = HEADER_PAT_C.search(line) if not (mA or mB) else None

        if mA:
            style_code = mA.group(1)
            sku_no = mA.group(4)
            size_token = mA.group(5)
            order_qty = mA.group(6)
        elif mB:
            style_code = mB.group(1)
            size_token = mB.group(2)
            order_qty = mB.group(3)
        elif mC:
            sku_no = mC.group(1)
            size_token = mC.group(2)
            order_qty = mC.group(3)
        else:
            i += 1
            continue

        block_lines = [line]
        j = i + 1
        sku_from_total = ""
        while j < len(lines):
            nxt = lines[j]
            if is_item_header_v2(nxt):
                break
            block_lines.append(nxt)
            t = TOTAL_SKU_PAT.search(nxt)
            if t:
                sku_from_total = t.group(1)
            j += 1
        i = j

        if not sku_no:
            sku_no = sku_from_total

        if not style_code:
            style_code = find_style_in_block(block_lines)

        tone = ''
        joined = " ".join(block_lines)
        if re.search(r"\bWG\s*750|WG750", joined):
            tone = 'W'
        elif re.search(r"\bRG\s*750|RG750", joined):
            tone = 'R'
        elif re.search(r"\bYG\s*750|YG750", joined):
            tone = 'Y'
        else:
            if re.search(r"WHITE", joined, re.IGNORECASE):
                tone = 'W'
            elif re.search(r"ROSE", joined, re.IGNORECASE):
                tone = 'R'
            elif re.search(r"YELLOW", joined, re.IGNORECASE):
                tone = 'Y'

        fineness = '750' if (re.search(r"\b18\s*CARA?\b", joined, re.IGNORECASE) or re.search(r"\b750\b", joined)) else ''

        diamond_quality = ''
        for bl in block_lines:
            m = re.search(r"\b([A-Z]{1,2}-?SI\d|[A-Z]{1,2}-?VS\d?|[A-Z]{1,2}-?VVS\d?|[A-Z]{1,2}-?I\d)\b", bl)
            if m:
                diamond_quality = m.group(1)
                break

        carat_line = ''
        for bl in block_lines:
            m = re.search(r"\b18\s*CARA?\s*-\s*750\b", bl, re.IGNORECASE)
            if m:
                carat_line = m.group(0)
                break
        if not carat_line and fineness == '750':
            carat_line = '18 CARA - 750'

        if style_code.startswith('R'):
            item_size = "" if size_token == 'STA' else (f"EU{size_token}" if size_token else "")
        else:
            item_size = ""

        metal = f"G{fineness}{tone}" if fineness and tone else (f"G{fineness}" if fineness else "")

        stamp_variable_text = 'lgd' if (default_stamp_var or '').lower() == 'lgd' else ''

        tone_to_desc = {'Y': 'YELLOW GOLD', 'W': 'WHITE GOLD', 'R': 'ROSE GOLD'}
        tone_desc = tone_to_desc.get(tone, '')
        parts = []
        if sku_no:
            parts.append(sku_no)
        if fineness or tone_desc:
            txt = " ".join([p for p in [fineness, tone_desc] if p]).strip()
            if txt:
                parts.append(txt)
        if diamond_quality:
            parts.append(f"DIA QLTY: {diamond_quality}")
        special_remarks = ",".join(parts)

        common_sentence = "Polishing and setting must be very well done."
        customer_prod_instruction = f"{carat_line}, {common_sentence}" if carat_line else common_sentence
        design_prod_instruction = "white rodium" if tone == 'W' else "no rodoium"

        items.append({
            'Sr.No': len(items) + 1,
            'Stylecode': style_code,
            'ItemSize': item_size,
            'OrderQty': order_qty,
            'OrderItemPcs': 1,
            'Metal': metal,
            'Tone': tone,
            'ItemPoNo': '',
            'ItemRefNo': '',
            'StockType': '',
            'Priority': default_priority,
            'MakeType': '',
            'CustomerProductionInstruction': customer_prod_instruction,
            'SpecialRemarks': special_remarks,
            'DesignProductionInstruction': design_prod_instruction,
            'StampInstruction': f"750+customer logo+{stamp_variable_text}".rstrip('+'),
            'OrderGroup': '',
            'Certificate': '',
            'SKUNo': sku_no,
            'Basestoneminwt': '',
            'Basestonemaxwt': '',
            'Basemetalminwt': '',
            'Basemetalmaxwt': '',
            'Productiondeliverydate': '',
            'Expecteddeliverydate': '',
            'SetPrice': '',
            'StoneQuality': '',
        })

    return items


def process_fsa_file(input_path: str, output_dir: str, default_priority: str = "REG", default_stamp_var: str = ""):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()

        if ext == '.pdf':
            text = extract_full_text_from_pdf(input_path)
            po_no = find_po_number(text)
            items = parse_items_v2(text, default_priority=(default_priority or "REG").upper(), default_stamp_var=(default_stamp_var or ""))
            for it in items:
                it['ItemPoNo'] = po_no
            requested_columns = [
                'Sr.No','Stylecode','ItemSize','OrderQty','OrderItemPcs','Metal','Tone','ItemPoNo','ItemRefNo','StockType','Priority','MakeType','CustomerProductionInstruction','SpecialRemarks','DesignProductionInstruction','StampInstruction','OrderGroup','Certificate','SKUNo','Basestoneminwt','Basestonemaxwt','Basemetalminwt','Basemetalmaxwt','Productiondeliverydate','Expecteddeliverydate','SetPrice','StoneQuality'
            ]
            df = pd.DataFrame(items)
            for col in requested_columns:
                if col not in df.columns:
                    df[col] = ''
            df = df[requested_columns]
            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
            df['Stylecode'] = df.apply(
                lambda row: _build_style_code(row['Stylecode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
            )
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            output_path = os.path.join(output_dir, f"{base_name}_FSA_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            # If Excel/CSV is provided, just echo structured content; mapping rules for Excel can be added later
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_FSA_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


