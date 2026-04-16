import os
import re
import pandas as pd
import pdfplumber
from typing import List, Dict, Optional


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'UP6.5', 'W') -> 'VR1943EEA-6.5WG'
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


def _read_pdf_text(pdf_path: str) -> str:
    full_text = ""
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            full_text += page_text + "\n"
    return full_text


def _extract_po_number(full_text: str) -> Optional[str]:
    m = re.search(r"PO\s*#\s*:\s*(\d+)", full_text, flags=re.IGNORECASE)
    return m.group(1) if m else None


def _split_items(full_text: str) -> List[str]:
    lines = [ln.strip() for ln in full_text.splitlines()]
    items: List[str] = []
    current: List[str] = []
    inside_items = False
    for ln in lines:
        if re.match(r"^\d+\.", ln):
            inside_items = True
            if current:
                items.append("\n".join(current).strip())
            current = [ln]
            continue
        if inside_items:
            current.append(ln)
            if "**FOR SHIMAYRA**" in ln or "FOR SHIMAYRA" in ln:
                items.append("\n".join(current).strip())
                current = []
                inside_items = False
    if current:
        items.append("\n".join(current).strip())
    return [it for it in items if it]


def _parse_first_line_tokens(line: str) -> List[str]:
    return re.findall(r"[A-Za-z0-9/*]+(?:-[A-Za-z0-9]+)*", line)


def _find_item_size_and_qty(line: str):
    tokens = re.findall(r"\d+\.\d+|\d+", line)
    if not tokens:
        return None, None
    size_idx = None
    for i, tok in enumerate(tokens):
        if re.match(r"^\d+\.\d+$", tok):
            size_idx = i
            break
    if size_idx is None:
        size_idx = 0
    size_val = tokens[size_idx]
    qty_val = None
    if size_idx + 1 < len(tokens):
        qty_val = tokens[size_idx + 1]
    return size_val, qty_val


def _find_item_ref_no(line: str) -> Optional[str]:
    m = re.search(r"(\d{5,}/\d+)", line)
    return m.group(1) if m else None


def _find_style_code(line: str) -> Optional[str]:
    tokens = _parse_first_line_tokens(line)
    for i, tok in enumerate(tokens):
        if re.match(r"^[A-Za-z]{3}/\d{1,2}/\d{4}$", tok):
            if i + 1 < len(tokens):
                next_tok = tokens[i + 1]
                if any(c.isalpha() for c in next_tok) and any(c.isdigit() for c in next_tok):
                    return next_tok
    candidate_codes: List[str] = []
    for t in tokens:
        if any(c.isalpha() for c in t) and any(c.isdigit() for c in t) and len(t) >= 6:
            candidate_codes.append(t)
    if len(candidate_codes) >= 2:
        return candidate_codes[-2]
    return candidate_codes[-1] if candidate_codes else None


def _find_sku(line: str) -> Optional[str]:
    m = re.search(r"\b([A-Z]{2,}[A-Z0-9]*-\d{1,})\b", line)
    if m:
        return m.group(1)
    tokens = _parse_first_line_tokens(line)
    for t in tokens:
        if '-' in t and any(ch.isdigit() for ch in t) and any(ch.isalpha() for ch in t):
            return t
    return None


def _find_metal_and_tone(block_text: str):
    metal_code = None
    tone = None
    if re.search(r"\bSILV\b", block_text, flags=re.IGNORECASE):
        metal_code = "AG925"
        m = re.search(r"\bSILV\b\s+([A-Z]{1,3})\b", block_text)
        if m:
            tone = m.group(1)
    if metal_code is None:
        if re.search(r"\bG14\w*\b|\b14K\b|\bGOLD\b", block_text, flags=re.IGNORECASE):
            metal_code = "G14"
            m2 = re.search(r"\bG14\w*\b\s+([A-Z]{1,3})\b", block_text)
            if m2:
                tone = m2.group(1)
    return metal_code, tone


def _find_customer_instruction_from_line(line: str) -> Optional[str]:
    toks = line.split()
    if not toks:
        return None
    sku_idx = None
    for i, t in enumerate(toks):
        if re.match(r"[A-Z]{2,}[A-Z0-9]*-\d{1,}$", t):
            sku_idx = i
            break
    if sku_idx is None:
        return None
    desc_terms: List[str] = []
    for j in range(sku_idx + 1, len(toks)):
        t = toks[j]
        if t.upper() in {"SILV", "G14", "G14W", "14K", "GOLD"}:
            break
        desc_terms.append(t)
    return " ".join(desc_terms) if desc_terms else None


def _extract_design_instructions(block_text: str) -> Optional[str]:
    phrases = re.findall(r"\*\*([^*]+)\*\*", block_text)
    return " ".join(p.strip() for p in phrases) if phrases else None


def _extract_stamp_instruction(block_text: str) -> Optional[str]:
    m = re.search(r"Special\s+Inst\.[^\n]*?STAMP\s+([^,\n]+)", block_text, flags=re.IGNORECASE)
    return m.group(1).strip() if m else None


def _extract_special_remarks(block_text: str) -> Optional[str]:
    for ln in block_text.splitlines():
        if ln.upper().startswith("SPECIAL INST."):
            return ln.split("Special Inst.", 1)[-1].strip()
    return None


def process_ambition_file(input_path: str, output_dir: str):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.pdf':
            full_text = _read_pdf_text(input_path)
            item_po_no = _extract_po_number(full_text) or ""
            blocks = _split_items(full_text)
            rows: List[Dict[str, str]] = []
            for blk in blocks:
                first_line = next((ln for ln in blk.splitlines() if re.match(r"^\d+\.\s", ln)),
                                  blk.splitlines()[0] if blk.splitlines() else "")
                sr_m = re.match(r"^(\d+)\.", first_line.strip())
                sr_no = sr_m.group(1) if sr_m else ""
                item_size, order_qty = _find_item_size_and_qty(first_line)
                item_ref_no = _find_item_ref_no(first_line) or ""
                style_code = _find_style_code(first_line) or ""
                sku_no = _find_sku(first_line) or ""
                metal, tone = _find_metal_and_tone(blk)
                cust_instr = _find_customer_instruction_from_line(first_line) or ""
                design_instr = _extract_design_instructions(blk) or ""
                stamp_instr = _extract_stamp_instruction(blk) or ""
                special_remarks = _extract_special_remarks(blk) or ""

                if style_code:
                    mt = re.search(r'-([A-Z]+)$', style_code)
                    if mt:
                        tone_full = mt.group(1)
                        tone = (tone_full[0] if tone_full else '').replace('V', 'W')
                        style_code = style_code[:mt.start()]

                # Convert size to INCH display format for ItemSize column
                if item_size:
                    try:
                        size_float = float(item_size)
                        item_size = f"{int(size_float)} INCH" if size_float.is_integer() else f"{size_float} INCH"
                    except ValueError:
                        item_size = f"{item_size} INCH"

                # Build StyleCode with INCH-formatted size so 'IN' is inserted correctly
                # e.g. '7 INCH' + 'W' -> 'BR0000094K-7INWG'
                # When Metal is AG925 (silver), use 'AG' suffix and keep Tone blank
                item_size = _map_item_size(item_size or '')
                _sc_tone = 'AG' if (metal or '').upper() == 'AG925' else (tone or '')
                style_code = _build_style_code(style_code, item_size or "", _sc_tone)

                rows.append({
                    "SrNo": sr_no,
                    "StyleCode": style_code,
                    "ItemSize": item_size or "",
                    "OrderQty": order_qty or "",
                    "OrderItemPcs": 1,
                    "Metal": metal or "",
                    "Tone": "" if (metal or '').upper() == 'AG925' else (tone or ""),
                    "ItemPoNo": item_po_no,
                    "ItemRefNo": item_ref_no,
                    "StockType": "",
                    "MakeType": "",
                    "CustomerProductionInstruction": cust_instr,
                    "SpecialRemarks": special_remarks,
                    "DesignProductionInstruction": design_instr,
                    "StampInstruction": stamp_instr,
                    "OrderGroup": "",
                    "Certificate": "",
                    "SKUNo": sku_no,
                    "Basestoneminwt": "",
                    "Basestonemaxwt": "",
                    "Basemetalminwt": "",
                    "Basemetalmaxwt": "",
                    "Productiondeliverydate": "",
                    "Expecteddeliverydate": "",
                    "BlankColumn": "",
                    "SetPrice": "",
                    "StoneQuality": "",
                })

            columns = [
                "SrNo","StyleCode","ItemSize","OrderQty","OrderItemPcs","Metal","Tone","ItemPoNo","ItemRefNo","StockType","MakeType","CustomerProductionInstruction","SpecialRemarks","DesignProductionInstruction","StampInstruction","OrderGroup","Certificate","SKUNo","Basestoneminwt","Basestonemaxwt","Basemetalminwt","Basemetalmaxwt","Productiondeliverydate","Expecteddeliverydate","BlankColumn","SetPrice","StoneQuality"
            ]
            df = pd.DataFrame(rows, columns=columns)
            output_path = os.path.join(output_dir, f"{base_name}_AMBITION_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_AMBITION_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


