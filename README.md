# GMP Automation System

A local web application that converts pharmaceutical-facility environmental measurement PDFs into structured Microsoft Excel reports with averages, limit indicators, and charts.

## Features

- Processes multiple PDFs for different AHUs in one request.
- Supports Online OCR through an API and Offline OCR through a self-hosted OCR endpoint.
- Produces an Excel workbook for each test type with data sheets, summary tables, and charts.
- Highlights values outside configured limits in red.
- Groups results by AHU and measurement semester.

## Supported Tests

| Code | Test | Excel File |
| --- | --- | --- |
| A | Airborne Particle Test | `Airborne_Particle_Test_Result_and_Graph.xlsx` |
| B | Air Velocity Test | `Air_Velocity_Test_Result_and_Graph.xlsx` |
| C | Air Change Rate Test | `Air_Change_Rate_Test_Result_and_Graph.xlsx` |
| D | HEPA Filter Test | `HEPA_Filter_Test_Result_and_Graph.xlsx` |
| E | Airflow Pattern Test | `Airflow_Pattern_Test_Result_and_Graph.xlsx` |

## Requirements

- Python 3.10 or later.
- Poppler, used to convert PDF pages into images.
- An internet connection for OCR.
- One of the following OCR modes:
  - Online OCR with an Anthropic API key.
  - Offline OCR with a self-hosted OCR server and its endpoint URL.

Install Poppler:

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

On Windows, download [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases), extract it to a location such as `C:\poppler`, and add `C:\poppler\Library\bin` to `PATH`.

## Installation

```bash
git clone <REPOSITORY_URL>
cd gmp_automation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

On Windows, activate the virtual environment with:

```bat
.venv\Scripts\activate
```

## Running the Application

```bash
python app.py
```

Or use the startup scripts:

```bash
# macOS/Linux
bash START_LINUX.sh

# Windows
START_WINDOWS.bat
```

Open `http://localhost:5001` in a browser. The application redirects to the Online OCR page at `http://localhost:5001/online`.

## Usage

1. Open `/online` for Online OCR or `/offline` for Offline OCR. Use the navigation arrows on the page to switch modes.
2. Enter an Anthropic API key for Online OCR, or an endpoint URL for Offline OCR.
3. Select one test type.
4. Upload one or more PDFs of the same test type.
5. Generate and download the Excel report.

Each PDF should contain one AHU for one semester. Result accuracy depends on the readability of the scanned PDF.

## Offline OCR with Kaggle

Offline OCR is a no-API-cost alternative to Online OCR. Run the OCR server notebook in Kaggle with a GPU and internet enabled: [Offline OCR Backend Server](https://www.kaggle.com/code/irfanqs/deepseek-ocr-backend-server).

1. Open the [Offline OCR Backend Server](https://www.kaggle.com/code/irfanqs/deepseek-ocr-backend-server) notebook.
2. Set the accelerator to GPU and enable Internet.
3. Add a Kaggle Secret named `NGROK_AUTHTOKEN`.
4. Run all notebook cells.
5. Copy the displayed `https://*.ngrok-free.app` URL into the Offline OCR page.

Kaggle sessions and ngrok URLs are temporary. The Offline OCR parser depends on the supported GMP document format; use `debug_ocr.py` to inspect raw OCR text when extraction needs to be reviewed.

```bash
python debug_ocr.py <path_to_pdf> <ngrok_url>
```

## Project Structure

```text
gmp_automation/
├── app.py                 # Flask application and upload/download endpoints
├── config.py              # Limits and test-type configuration
├── ocr_engine.py          # PDF extraction through Online OCR
├── deepseek_ocr/          # Offline OCR client, parser, and backend notebook
├── excel_generator.py     # Excel report and chart generation
├── templates/index.html   # Web interface
├── boilerplate/           # Example Excel templates
├── uploads/               # Temporary PDF storage during processing
├── outputs/               # Generated Excel reports
├── debug_ocr.py           # Offline OCR inspection tool
└── requirements.txt       # Python dependencies
```

## Notes

- Online OCR incurs charges according to Anthropic account usage.
- Limits and Excel settings are configured in `config.py`.
- The application limits each upload request to 100 MB.
- Uploaded files are removed after a successful Excel report is generated.
