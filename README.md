# OPS Tool — Order Processing Dashboard

A Flask-based internal operations tool for processing, cleaning, and normalizing purchase-order data across multiple customers. Upload Excel, CSV, or PDF files through a web dashboard and receive standardized, downloadable outputs.

## Supported Customers

Ambition, Anaya, AAM, Bhakti Dharam, Craft, DCT, FSA, HK, JJL, JU, MOR, NGL, OBU, OMJ, PC2, PCB, RBL, SGI, SHEFI, SHEFI New PO, Uneek, VIMCO, and more.

## Tech Stack

| Layer | Technology |
|---|---|
| Backend | Python 3, Flask |
| Data Processing | pandas, NumPy, openpyxl, xlrd |
| PDF Parsing | pdfplumber |
| Frontend | Jinja2 templates, Bootstrap 5, Chart.js |

## Project Structure

```
OPS_Tool_09042026/
├── app.py                  # Main Flask application (routes, uploads, stats)
├── app5_pdf2excel.py       # Standalone PDF-to-Excel utility app
├── requirements.txt        # Python dependencies
├── order_stats.json        # Dashboard statistics (auto-updated)
├── templates/              # HTML templates (one per customer + dashboard)
│   └── index.html          # Order Processing Dashboard
├── OMJ.py, SHEFI.py, ...   # Per-customer processing modules
├── SHEFI_PO_DHAVAL/        # Alternate SHEFI pipeline (dynamically loaded)
├── Jupyter_Notebooks/      # Development & exploration notebooks
└── README.md
```

## Getting Started

### Prerequisites

- Python 3.10 or higher

### Installation

```bash
cd OPS_Tool_09042026
python -m venv venv
venv\Scripts\activate        # On Windows
pip install -r requirements.txt
```

### Running the Application

```bash
python app.py
```

The dashboard will be available at **http://localhost:5000**.

### PDF-to-Excel Utility (Optional)

```bash
python app5_pdf2excel.py
```

## How It Works

1. Open the dashboard in your browser.
2. Select the customer-specific route from the navigation.
3. Upload a purchase-order file (`.xlsx`, `.xls`, `.csv`, or `.pdf`).
4. The corresponding processor cleans and normalizes the data (builds style codes, maps sizes, standardizes columns, etc.).
5. Download the processed output file.
6. Per-customer file and order counts are tracked on the main dashboard via `order_stats.json`.

## Configuration

| Setting | Default | Location |
|---|---|---|
| Port | `5000` | `app.py` |
| Max Upload Size | 16 MB | `app.py` |
| Upload Folder | System temp directory | `app.py` |

## License

Internal use only.
