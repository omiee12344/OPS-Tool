from flask import Flask, render_template, request, send_file, redirect, url_for
import os
import json
from datetime import datetime
from werkzeug.utils import secure_filename
import tempfile
from pathlib import Path
# Original modules
from OMJ import process_omj_file
from SHEFI import process_shefi_file
# Previously added modules
from ambition import process_ambition_file
from craft import process_craft_file
from hk import process_hk_file
from fsa import process_fsa_file
from jjl import process_jjl_file
from obu import process_obu_file
from rbl import process_rbl_file
from anaya import process_anaya_file
from uneek import process_uneek_file
from JU import process_ju_excel_file
# New customer modules
from AAM import process_aam_file
from Bhakti_Dharam import process_bhakti_dharm_file
from DCT import process_dct_file
from MOR import process_mor_file
from NGL import process_ngl_file
from PC2 import process_pc2_file
from PCB import process_pcb_file
from SGI import process_sgi_file
from VIMCO import process_vimco_file
import importlib.util as _ilu
_shefi_new_spec = _ilu.spec_from_file_location(
    'shefi_dhaval',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SHEFI_PO_DHAVAL', 'shefi.py')
)
_shefi_new_mod = _ilu.module_from_spec(_shefi_new_spec)
_shefi_new_spec.loader.exec_module(_shefi_new_mod)
process_shefi_new_file = _shefi_new_mod.process_shefi_new_file
del _shefi_new_spec, _shefi_new_mod, _ilu

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf', 'csv'}

# --- Order Stats Tracking ---
STATS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'order_stats.json')

def load_stats():
    if os.path.exists(STATS_FILE):
        try:
            with open(STATS_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return {'customers': {}, 'total_files': 0, 'total_orders': 0}

def save_stats(stats):
    try:
        with open(STATS_FILE, 'w') as f:
            json.dump(stats, f, indent=2)
    except Exception:
        pass

def record_processing(customer_name, files_count, rows_count):
    stats = load_stats()
    cust = stats['customers'].setdefault(customer_name, {
        'files_processed': 0, 'orders_processed': 0, 'last_processed': None
    })
    cust['files_processed'] += files_count
    cust['orders_processed'] += int(rows_count or 0)
    cust['last_processed'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    stats['total_files'] = stats.get('total_files', 0) + files_count
    stats['total_orders'] = stats.get('total_orders', 0) + int(rows_count or 0)
    save_stats(stats)
# --- End Stats Tracking ---

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    stats = load_stats()
    return render_template('index.html', stats=stats)

@app.route('/omj')
def omj_tool():
    return render_template('index_omj.html')

@app.route('/shefi')
def shefi_tool():
    return render_template('index_shefi.html')

# New tool pages
@app.route('/ambition')
def ambition_tool():
    return render_template('index_ambition.html')

@app.route('/craft')
def craft_tool():
    return render_template('index_craft_hk.html')

@app.route('/hk')
def hk_tool():
    return render_template('index_hk.html')

@app.route('/fsa')
def fsa_tool():
    return render_template('index_fsa.html')

@app.route('/jjl')
def jjl_tool():
    return render_template('index_jjl.html')

@app.route('/obu')
def obu_tool():
    return render_template('index_obu.html')

@app.route('/rbl')
def rbl_tool():
    return render_template('index_rbl.html')

@app.route('/anaya')
def anaya_tool():
    return render_template('index_anaya.html')

@app.route('/uneek')
def uneek_tool():
    return render_template('index_uneek.html')

@app.route('/ju')
def ju_tool():
    return render_template('index_ju.html')

@app.route('/aam')
def aam_tool():
    return render_template('index_aam.html')

@app.route('/bhakti_dharam')
def bhakti_dharam_tool():
    return render_template('index_bhakti_dharam.html')

@app.route('/dct')
def dct_tool():
    return render_template('index_dct.html')

@app.route('/mor')
def mor_tool():
    return render_template('index_mor.html')

@app.route('/ngl')
def ngl_tool():
    return render_template('index_ngl.html')

@app.route('/pc2')
def pc2_tool():
    return render_template('index_pc2.html')

@app.route('/pcb')
def pcb_tool():
    return render_template('index_pcb.html')

@app.route('/sgi')
def sgi_tool():
    return render_template('index_sgi.html')

@app.route('/vimco')
def vimco_tool():
    return render_template('index_vimco.html')

@app.route('/process-omj', methods=['POST'])
def process_omj():
    try:
        # Support both single file ('file') and multiple files ('files')
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']  # Filter empty files
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template('index_omj.html', error='No file selected')
        
        if not files:
            return render_template('index_omj.html', error='No file selected')
        
        # Check if separate processing is requested
        process_separately = request.form.get('separate') == 'true'
        
        # Save uploaded files
        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)
        
        if not valid_files:
            return render_template('index_omj.html', error='No valid Excel files found')
        
        try:
            download_urls = []
            
            if process_separately or len(valid_files) == 1:
                # Process each file separately
                for filepath in valid_files:
                    success, output_path, error, df = process_omj_file(filepath, app.config['UPLOAD_FOLDER'])
                    
                    if success:
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                    else:
                        return render_template('index_omj.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                success_msg = f'Successfully processed {len(valid_files)} file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
                
                # Clean up input files
                for filepath in valid_files:
                    try:
                        os.remove(filepath)
                    except:
                        pass
                
                record_processing('OMJ', len(valid_files), sum(d.get('rows', 0) or 0 for d in download_urls))
                return render_template('index_omj.html', 
                                     success=success_msg,
                                     download_urls=download_urls if len(download_urls) > 1 else None,
                                     download_url=download_urls[0]['url'] if len(download_urls) == 1 else None)
            else:
                # Combine all files
                all_dataframes = []
                for filepath in valid_files:
                    success, output_path, error, df = process_omj_file(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        all_dataframes.append(df)
                    else:
                        # Clean up on error
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template('index_omj.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                # Combine all dataframes
                import pandas as pd
                combined_df = pd.concat(all_dataframes, ignore_index=True)
                
                # Generate combined output filename
                output_filename = 'OMJ_CASTING_PO_Cleaned_Combined.csv'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                combined_df.to_csv(output_path, index=False)
                
                # Clean up input files
                for filepath in valid_files:
                    try:
                        os.remove(filepath)
                    except:
                        pass
                
                record_processing('OMJ', len(valid_files), len(combined_df))
                return render_template('index_omj.html', 
                                     success=f'Successfully processed and combined {len(valid_files)} file(s)!',
                                     download_url=url_for('download_file', filename=output_filename))
        
        except Exception as proc_error:
            # Clean up input files on processing error
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            raise proc_error
    
    except Exception as e:
        return render_template('index_omj.html', error=f'Error processing file: {str(e)}')

@app.route('/process-shefi', methods=['POST'])
def process_shefi():
    try:
        # Support both single file ('file') and multiple files ('files')
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']  # Filter empty files
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template('index_shefi.html', error='No file selected')
        
        if not files:
            return render_template('index_shefi.html', error='No file selected')
        
        # Check if separate processing is requested
        process_separately = request.form.get('separate') == 'true'
        
        # Save uploaded files
        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)
        
        if not valid_files:
            return render_template('index_shefi.html', error='No valid Excel files found')
        
        try:
            download_urls = []
            
            if process_separately or len(valid_files) == 1:
                # Process each file separately
                for filepath in valid_files:
                    success, output_path, error, df = process_shefi_file(filepath, app.config['UPLOAD_FOLDER'])
                    
                    if success:
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                    else:
                        return render_template('index_shefi.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                success_msg = f'Successfully processed {len(valid_files)} file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
                
                # Clean up input files
                for filepath in valid_files:
                    try:
                        os.remove(filepath)
                    except:
                        pass
                
                record_processing('SHEFI', len(valid_files), sum(d.get('rows', 0) or 0 for d in download_urls))
                return render_template('index_shefi.html', 
                                     success=success_msg,
                                     download_urls=download_urls if len(download_urls) > 1 else None,
                                     download_url=download_urls[0]['url'] if len(download_urls) == 1 else None)
            else:
                # Combine all files
                all_dataframes = []
                for filepath in valid_files:
                    success, output_path, error, df = process_shefi_file(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        all_dataframes.append(df)
                    else:
                        # Clean up on error
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template('index_shefi.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                # Combine all dataframes
                import pandas as pd
                combined_df = pd.concat(all_dataframes, ignore_index=True)
                
                # Generate combined output filename
                output_filename = 'GATI_FORMAT_SHEFI_CLEAN_Combined.xlsx'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                combined_df.to_excel(output_path, index=False)
                
                # Clean up input files
                for filepath in valid_files:
                    try:
                        os.remove(filepath)
                    except:
                        pass
                
                record_processing('SHEFI', len(valid_files), len(combined_df))
                return render_template('index_shefi.html', 
                                     success=f'Successfully processed and combined {len(valid_files)} file(s)!',
                                     download_url=url_for('download_file', filename=output_filename))
        
        except Exception as proc_error:
            # Clean up input files on processing error
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            raise proc_error
    
    except Exception as e:
        return render_template('index_shefi.html', error=f'Error processing file: {str(e)}')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(
        os.path.join(app.config['UPLOAD_FOLDER'], filename),
        as_attachment=True,
        download_name=filename
    )

# Generic processor utility for newly added tools
def _handle_generic_processing(request, template_name, processor_func, output_ext_default, customer_name=None):
    try:
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template(template_name, error='No file selected')

        if not files:
            return render_template(template_name, error='No file selected')

        process_separately = request.form.get('separate') == 'true'

        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)

        if not valid_files:
            return render_template(template_name, error='No valid files found')

        try:
            download_urls = []
            if process_separately or len(valid_files) == 1:
                for filepath in valid_files:
                    success, output_path, error, df = processor_func(filepath, app.config['UPLOAD_FOLDER'])
                    if success:
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                    else:
                        return render_template(template_name, error=f'Error processing {os.path.basename(filepath)}: {error}')

                success_msg = f'Successfully processed {len(valid_files)} file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
                for fp in valid_files:
                    try:
                        os.remove(fp)
                    except:
                        pass
                if customer_name:
                    total_rows = sum(d.get('rows', 0) or 0 for d in download_urls)
                    record_processing(customer_name, len(valid_files), total_rows)
                return render_template(template_name,
                                       success=success_msg,
                                       download_urls=download_urls if len(download_urls) > 1 else None,
                                       download_url=download_urls[0]['url'] if len(download_urls) == 1 else None)
            else:
                import pandas as pd
                dataframes = []
                for filepath in valid_files:
                    success, output_path, error, df = processor_func(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        dataframes.append(df)
                    else:
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template(template_name, error=f'Error processing {os.path.basename(filepath)}: {error}')

                if not dataframes:
                    for fp in valid_files:
                        try:
                            os.remove(fp)
                        except:
                            pass
                    return render_template(template_name, error='No dataframes produced to combine')

                combined_df = pd.concat(dataframes, ignore_index=True)
                output_filename = f'combined_output.{output_ext_default}'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], output_filename)
                if output_ext_default == 'xlsx':
                    combined_df.to_excel(output_path, index=False)
                else:
                    combined_df.to_csv(output_path, index=False)

                for fp in valid_files:
                    try:
                        os.remove(fp)
                    except:
                        pass
                if customer_name:
                    record_processing(customer_name, len(valid_files), len(combined_df))
                return render_template(template_name,
                                       success=f'Successfully processed and combined {len(valid_files)} file(s)!',
                                       download_url=url_for('download_file', filename=output_filename))
        except Exception as proc_error:
            for fp in valid_files:
                try:
                    os.remove(fp)
                except:
                    pass
            raise proc_error
    except Exception as e:
        return render_template(template_name, error=f'Error processing file: {str(e)}')


@app.route('/process-ambition', methods=['POST'])
def process_ambition():
    return _handle_generic_processing(request, 'index_ambition.html', process_ambition_file, 'xlsx', 'Ambition')


@app.route('/process-craft', methods=['POST'])
def process_craft():
    # Read user inputs from form
    size_prefix = (request.form.get('size_prefix') or 'US').strip()
    default_priority = (request.form.get('default_priority') or 'REG').strip().upper()

    # Bind arguments via a wrapper so generic handler can call with (path, out_dir)
    def _proc(path, out_dir):
        return process_craft_file(path, out_dir, size_prefix=size_prefix, default_priority=default_priority)

    return _handle_generic_processing(request, 'index_craft_hk.html', _proc, 'xlsx', 'Craft')


@app.route('/process-hk', methods=['POST'])
def process_hk():
    size_prefix = (request.form.get('size_prefix') or 'US').strip()
    default_priority = (request.form.get('default_priority') or 'REG').strip().upper()

    def _proc(path, out_dir):
        return process_hk_file(path, out_dir, size_prefix=size_prefix, default_priority=default_priority)

    return _handle_generic_processing(request, 'index_hk.html', _proc, 'xlsx', 'HK')


@app.route('/process-fsa', methods=['POST'])
def process_fsa():
    default_priority = (request.form.get('default_priority') or 'REG').strip().upper()
    stamp_var = (request.form.get('stamp_var') or '').strip().lower()  # '' or 'lgd'

    def _proc(path, out_dir):
        return process_fsa_file(path, out_dir, default_priority=default_priority, default_stamp_var=stamp_var)

    return _handle_generic_processing(request, 'index_fsa.html', _proc, 'xlsx', 'FSA')


@app.route('/process-jjl', methods=['POST'])
def process_jjl():
    default_priority = (request.form.get('default_priority') or 'REG').strip()
    diamond_quality = (request.form.get('diamond_quality') or 'REG').strip()
    def _proc(path, out_dir):
        return process_jjl_file(path, out_dir, default_priority=default_priority, default_diamond_quality=diamond_quality)
    return _handle_generic_processing(request, 'index_jjl.html', _proc, 'xlsx', 'JJL')


@app.route('/process-obu', methods=['POST'])
def process_obu():
    return _handle_generic_processing(request, 'index_obu.html', process_obu_file, 'xlsx', 'OBU')


@app.route('/process-rbl', methods=['POST'])
def process_rbl():
    end_customer_name = (request.form.get('end_customer_name') or '').strip()
    priority = (request.form.get('priority') or '').strip()
    def _proc(path, out_dir):
        return process_rbl_file(path, out_dir, end_customer_name=end_customer_name, priority_value=priority)
    return _handle_generic_processing(request, 'index_rbl.html', _proc, 'xlsx', 'RBL')


@app.route('/process-anaya', methods=['POST'])
def process_anaya():
    tone = (request.form.get('tone') or 'Y').strip().upper()
    def _proc(path, out_dir):
        return process_anaya_file(path, out_dir, tone=tone)
    return _handle_generic_processing(request, 'index_anaya.html', _proc, 'csv', 'Anaya')


@app.route('/process-uneek', methods=['POST'])
def process_uneek():
    po_value = (request.form.get('po_value') or '').strip()
    item_no = (request.form.get('item_no') or '').strip()
    base_serial_start_raw = (request.form.get('base_serial_start') or '').strip()
    style_code = (request.form.get('style_code') or '').strip()
    item_size = (request.form.get('item_size') or '').strip()

    # Safely convert base_serial_start to int if provided
    base_serial_start = None
    if base_serial_start_raw:
        try:
            base_serial_start = int(base_serial_start_raw)
        except ValueError:
            base_serial_start = None

    def _proc(path, out_dir):
        return process_uneek_file(
            path,
            out_dir,
            po_value=po_value,
            item_no=item_no,
            base_serial_start=base_serial_start,
            style_code_input=style_code,
            item_size_input=item_size,
        )

    return _handle_generic_processing(request, 'index_uneek.html', _proc, 'xlsx', 'UNEEK')


@app.route('/process-ju', methods=['POST'])
def process_ju():
    item_po_no = (request.form.get('item_po_no') or '').strip()
    priority   = (request.form.get('priority') or 'REG').strip().upper()

    def _proc(path, out_dir):
        return process_ju_excel_file(path, out_dir, item_po_no=item_po_no, priority=priority)

    return _handle_generic_processing(request, 'index_ju.html', _proc, 'xlsx', 'JU')


@app.route('/process-aam', methods=['POST'])
def process_aam():
    priority = (request.form.get('priority') or '').strip()
    
    def _proc(path, out_dir):
        return process_aam_file(path, out_dir, priority_value=priority)
    
    return _handle_generic_processing(request, 'index_aam.html', _proc, 'xlsx', 'AAM')


@app.route('/process-bhakti_dharam', methods=['POST'])
def process_bhakti_dharam():
    item_po_no = (request.form.get('item_po_no') or '').strip()
    stamp_instruction = (request.form.get('stamp_instruction') or '').strip()
    order_group = (request.form.get('order_group') or '').strip()
    priority = (request.form.get('priority') or '5 day').strip()
    po_no = (request.form.get('po_no') or '').strip()
    size_prefix = (request.form.get('size_prefix') or 'US').strip()
    
    def _proc(path, out_dir):
        return process_bhakti_dharm_file(path, out_dir, item_po_no=item_po_no, 
                                         stamp_instruction=stamp_instruction,
                                         order_group=order_group, priority_value=priority,
                                         po_no_value=po_no, size_prefix=size_prefix)
    
    return _handle_generic_processing(request, 'index_bhakti_dharam.html', _proc, 'xlsx', 'Bhakti & Dharam')


@app.route('/process-dct', methods=['POST'])
def process_dct():
    priority = (request.form.get('priority') or '').strip()
    
    def _proc(path, out_dir):
        return process_dct_file(path, out_dir, priority=priority)
    
    return _handle_generic_processing(request, 'index_dct.html', _proc, 'csv', 'DCT')


@app.route('/process-mor', methods=['POST'])
def process_mor():
    item_po_no = (request.form.get('item_po_no') or '').strip()
    priority = (request.form.get('priority') or '').strip()
    
    def _proc(path, out_dir):
        return process_mor_file(path, out_dir, item_po_no=item_po_no, priority_value=priority)
    
    return _handle_generic_processing(request, 'index_mor.html', _proc, 'xlsx', 'MOR')


@app.route('/process-ngl', methods=['POST'])
def process_ngl():
    order_qty = (request.form.get('order_qty') or '').strip()
    item_po_no = (request.form.get('item_po_no') or '').strip()
    priority = (request.form.get('priority') or '').strip()
    additional_after_dia = (request.form.get('additional_after_dia') or '').strip()
    
    def _proc(path, out_dir):
        return process_ngl_file(path, out_dir, order_qty=order_qty, item_po_no=item_po_no,
                               priority=priority, additional_after_dia=additional_after_dia)
    
    return _handle_generic_processing(request, 'index_ngl.html', _proc, 'csv', 'NGL')


@app.route('/process-pc2', methods=['POST'])
def process_pc2():
    return _handle_generic_processing(request, 'index_pc2.html', process_pc2_file, 'xlsx', 'PC2')


@app.route('/process-pcb', methods=['POST'])
def process_pcb():
    priority = (request.form.get('priority') or '').strip()
    skuno = (request.form.get('skuno') or '').strip()
    
    def _proc(path, out_dir):
        return process_pcb_file(path, out_dir, priority_value=priority, skuno_value=skuno)
    
    return _handle_generic_processing(request, 'index_pcb.html', _proc, 'csv', 'PCB')


@app.route('/process-sgi', methods=['POST'])
def process_sgi():
    cust_order_no = (request.form.get('cust_order_no') or '').strip()
    
    def _proc(path, out_dir):
        return process_sgi_file(path, out_dir, cust_order_no=cust_order_no)
    
    return _handle_generic_processing(request, 'index_sgi.html', _proc, 'csv', 'SGI')


@app.route('/process-vimco', methods=['POST'])
def process_vimco():
    item_po_no = (request.form.get('item_po_no') or '').strip()
    order_group = (request.form.get('order_group') or '').strip()
    priority = (request.form.get('priority') or '5 day').strip()
    
    def _proc(path, out_dir):
        return process_vimco_file(path, out_dir, item_po_no=item_po_no,
                                  order_group=order_group, priority_value=priority)
    
    return _handle_generic_processing(request, 'index_vimco.html', _proc, 'csv', 'VIMCO')


@app.route('/shefi-new')
def shefi_new_tool():
    return render_template('index_shefi_new.html')


@app.route('/process-shefi-new', methods=['POST'])
def process_shefi_new():
    return _handle_generic_processing(request, 'index_shefi_new.html',
                                      process_shefi_new_file, 'xlsx', 'SHEFI New PO')


if __name__ == '__main__':
    #app.run(debug=True)
    app.run(debug=True, host='0.0.0.0', port=5000)