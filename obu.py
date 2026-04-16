import os
import re
import pandas as pd
import pdfplumber


def _build_style_code(base, item_size, tone):
    """
    Build StyleCode as '<base>-<size_numeric><tone>G' (or PT for platinum).
    e.g. ('VR1943EEA', 'EU52', 'W') -> 'VR1943EEA-52WG'
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
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join([page.extract_text() or '' for page in pdf.pages])


def process_obu_file(input_path: str, output_dir: str):
    try:
        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext = os.path.splitext(input_path)[1].lower()
        if ext == '.pdf':
            text = _read_pdf_text(input_path)
            code_token_re = re.compile(r'[A-Z0-9][A-Z0-9\-]*[A-Z0-9]')
            sku_first_code_re = re.compile(r'^(?P<sku>\d+-[A-Z]{2}\d{3})')
            po_re = re.compile(r'PO#\s*:\s*(\d+)')
            article_header_re = re.compile(r'^Article code', re.IGNORECASE)
            quantity_line_re = re.compile(r'^(\d+)\s+(\d+)$')
            style_tone_re = re.compile(r'-([A-Z]{1,3})$')

            po_match = po_re.search(text)
            item_po_no = po_match.group(1) if po_match else ''

            lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
            blocks, current = [], []
            for ln in lines:
                if article_header_re.match(ln):
                    if current:
                        blocks.append(current)
                        current = []
                    current.append(ln)
                elif current:
                    current.append(ln)
            if current:
                blocks.append(current)

            items = []
            for b_index, block in enumerate(blocks):
                is_last_block = (b_index == len(blocks) - 1)
                btxt = "\n".join(block)

                codes = []
                for idx, ln in enumerate(block[1:4], start=1):
                    tokens = code_token_re.findall(ln)
                    if tokens:
                        codes.extend(tokens)
                    if len(codes) >= 2:
                        break
                codes = [re.sub(r'[^A-Z0-9\-]', '', c) for c in codes]

                sku_full = codes[0] if len(codes) >= 1 else ''
                item_ref_no = codes[1] if len(codes) == 3 else ''
                style_code = codes[2] if len(codes) == 3 else (codes[1] if len(codes) == 2 else '')

                item_size = ''
                tone_match = re.search(r'([YWR]G)-(\d+(?:-\d+)*)', sku_full)
                if tone_match:
                    after_tone = tone_match.group(2)
                    nums = re.findall(r'\d+', after_tone)
                    if len(nums) >= 2:
                        item_size = nums[1]
                if item_size:
                    item_size = f"EU{item_size}"

                tone = ''
                if style_code:
                    mt = style_tone_re.search(style_code)
                    if mt:
                        tone_full = mt.group(1)
                        tone = tone_full[0] if tone_full else ''
                        style_code = style_code[:mt.start()]

                sr_no = ''
                order_qty = ''
                order_item_pcs = ''
                try:
                    desc_idx = next(i for i, ln in enumerate(block) if ln.lower().startswith('description'))
                except StopIteration:
                    desc_idx = None
                if desc_idx is not None:
                    for ln in block[desc_idx:desc_idx+5]:
                        qm = quantity_line_re.match(ln)
                        if qm:
                            sr_no = qm.group(1)
                            order_qty = qm.group(2)
                            order_item_pcs = order_qty
                            break

                desc_lines = []
                for ln in block:
                    if article_header_re.match(ln) or quantity_line_re.match(ln) or ln.lower().startswith('description'):
                        continue
                    if code_token_re.fullmatch(ln.replace(' ', '')):
                        continue
                    desc_lines.append(ln)
                full_desc = ' '.join(desc_lines)
                if is_last_block:
                    pot_idx = re.search(r'Purchase order Total', full_desc, flags=re.IGNORECASE)
                    if pot_idx:
                        full_desc = full_desc[:pot_idx.start()].strip()
                split_match = re.search(r'\b(stamp\b.*)', full_desc, flags=re.IGNORECASE)
                if split_match:
                    customer_instr = full_desc[:split_match.start()].strip()
                    stamp_instr = full_desc[split_match.start():].strip()
                else:
                    customer_instr = full_desc
                    stamp_instr = ''
                customer_instr = re.sub(r'\s*\band\s*$', '', customer_instr, flags=re.IGNORECASE).strip()

                certificate = ''
                if sku_full:
                    nums = re.findall(r'\d+', sku_full)
                    if nums and nums[-1] == '100':
                        certificate = 'IGI Certified'

                sku_no = ''
                if sku_full:
                    m = re.match(r'^(\d+-[A-Z]{2}\d{3})', sku_full)
                    if m:
                        sku_no = m.group(1)
                    else:
                        parts = sku_full.split('-')
                        if len(parts) >= 2:
                            sku_no = parts[0] + '-' + parts[1]

                items.append({
                    'SrNo': sr_no,
                    'StyleCode': style_code,
                    'ItemSize': item_size,
                    'OrderQty': order_qty,
                    'OrderItemPcs': 1,
                    'Metal': '',
                    'Tone': tone,
                    'ItemPoNo': item_po_no,
                    'ItemRefNo': item_ref_no,
                    'StockType': '',
                    'MakeType': '',
                    'CustomerProductionInstruction': customer_instr,
                    'SpecialRemarks': '',
                    'DesignProductionInstruction': '',
                    'StampInstruction': stamp_instr,
                    'OrderGroup': '',
                    'Certificate': certificate,
                    'SKUNo': sku_no,
                    'Basestoneminwt': '',
                    'Basestonemaxwt': '',
                    'Basemetalminwt': '',
                    'Basemetalmaxwt': '',
                    'Productiondeliverydate': '',
                    'Expecteddeliverydate': '',
                    'Blank': '',
                    'SetPrice': '',
                    'StoneQuality': 'VVS+' if re.search(r'\bVVS\+\b', btxt) else ''
                })

            columns_order = [
                'SrNo','StyleCode','ItemSize','OrderQty','OrderItemPcs','Metal','Tone','ItemPoNo','ItemRefNo',
                'StockType','MakeType','CustomerProductionInstruction','SpecialRemarks','DesignProductionInstruction',
                'StampInstruction','OrderGroup','Certificate','SKUNo','Basestoneminwt','Basestonemaxwt','Basemetalminwt',
                'Basemetalmaxwt','Productiondeliverydate','Expecteddeliverydate','Blank', 'SetPrice','StoneQuality'
            ]
            df = pd.DataFrame(items, columns=columns_order)
            df['ItemSize'] = df['ItemSize'].apply(_map_item_size)
            df['StyleCode'] = df.apply(
                lambda row: _build_style_code(row['StyleCode'], row['ItemSize'], 'AG' if str(row.get('Metal', '')).upper() == 'AG925' else str(row['Tone'])), axis=1
            )
            df.loc[df['Metal'].astype(str).str.upper() == 'AG925', 'Tone'] = ''
            output_path = os.path.join(output_dir, f"{base_name}_OBU_MAPPED.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        elif ext in ['.xlsx', '.xls', '.csv']:
            df = pd.read_excel(input_path) if ext in ['.xlsx', '.xls'] else pd.read_csv(input_path)
            output_path = os.path.join(output_dir, f"{base_name}_OBU_PASSTHROUGH.xlsx")
            df.to_excel(output_path, index=False)
            return True, output_path, None, df
        else:
            return False, None, f"Unsupported file type: {ext}", None
    except Exception as e:
        return False, None, str(e), None


