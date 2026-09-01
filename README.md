# GMP Automation System

[![Korean README](https://img.shields.io/badge/README-한국어-2563eb)](README.kr.md)

A local web application that analyzes pharmaceutical and biotechnology facility environmental measurement PDF records and converts them into Microsoft Excel reports with averages, alert/action limits, and charts.

## Features

- **Batch processing for multiple PDFs:** Upload and process PDF records for multiple AHUs in one request.
- **Flexible OCR options:** Supports API-based Online OCR and a self-hosted Offline OCR server.
- **Automatic Excel report generation:** Creates data sheets, summary tables, and pivot/chart sheets for each test type.
- **Limit violation highlighting:** Highlights measurements outside the configured alert/action limits in red.
- **Automatic AHU and semester sorting:** Groups data by air handling unit (AHU) number and measurement semester.

## Supported Tests

| Code | Test | Generated Excel File |
| --- | --- | --- |
| A | Airborne Particle Test | `Airborne_Particle_Test_Result_and_Graph.xlsx` |
| B | Air Velocity Test | `Air_Velocity_Test_Result_and_Graph.xlsx` |
| C | Air Change Rate Test | `Air_Change_Rate_Test_Result_and_Graph.xlsx` |
| D | HEPA Filter Test | `HEPA_Filter_Test_Result_and_Graph.xlsx` |
| E | Airflow Pattern Test | `Airflow_Pattern_Test_Result_and_Graph.xlsx` |

## Requirements

- Python 3.10 or later
- Poppler for converting PDF pages into images
- A network connection for OCR processing
- One of the following OCR modes:
  - **Online OCR** with an Anthropic API key
  - **Offline OCR** with a self-hosted OCR endpoint URL

### Installing Poppler

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

On Windows, download [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases), extract it to `C:\poppler`, and add `C:\poppler\Library\bin` to the system `PATH` environment variable.

## Installation

```bash
git clone https://github.com/irfanqs/gmp-automation.git
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

Alternatively, use the provided startup scripts:

```bash
# macOS/Linux
bash START_LINUX.sh

# Windows
START_WINDOWS.bat
```

Opening `http://localhost:5001` redirects to the Online OCR page at `http://localhost:5001/online`.

## Docker Deployment

Docker Engine and the Docker Compose plugin are required. Start the application with:

```bash
bash deploy.sh
```

By default, the web application is available at `http://localhost:5000/`. To use a different host port:

```bash
HOST_PORT=8080 bash deploy.sh
```

Generated Excel files and temporary uploads are stored in Docker volumes. Use the following commands to inspect logs or stop the application:

```bash
docker compose -p gmp-main logs -f
docker compose -p gmp-main down
```

## Usage

1. Open `/online` for Online OCR or `/offline` for Offline OCR. Use the navigation buttons at the top of the page to switch modes.
2. Enter an Anthropic API key for Online OCR or an Offline OCR endpoint URL.
3. Select one of the supported test types (A-E).
4. Upload one or more PDFs of the same test type.
5. Click `Start Excel Generation` to generate and download the report.

Each PDF must contain the measurement record for one AHU and one semester. OCR accuracy depends on the quality of the scanned PDF.

## Offline OCR with Kaggle

Offline OCR is an alternative that does not incur API costs. Run the server in a Kaggle environment with GPU and internet access enabled: [Offline OCR Backend Server](https://www.kaggle.com/code/irfanqs/deepseek-ocr-backend-server).

1. Open the [Offline OCR Backend Server](https://www.kaggle.com/code/irfanqs/deepseek-ocr-backend-server) notebook.
2. Set the accelerator to GPU and enable Internet.
3. Add your ngrok token to Kaggle Secrets as `NGROK_AUTHTOKEN`.
4. Run all notebook cells.
5. Copy the displayed `https://*.ngrok-free.app` URL into the Offline OCR endpoint field in the web UI.

Kaggle sessions and ngrok URLs are temporary, so the URL changes when the notebook restarts. Use `debug_ocr.py` to inspect and validate extracted OCR data:

```bash
python debug_ocr.py <PATH_TO_PDF> <NGROK_URL>
```

## Project Structure

```text
gmp_automation/
├── app.py                 # Flask application and endpoints
├── config.py              # Test limits and environment configuration
├── ocr_engine.py          # Online OCR PDF extraction engine
├── deepseek_ocr/          # Offline OCR client, parser, and server code
├── excel_generator.py     # Excel report and chart generator
├── templates/index.html   # Web interface template
├── boilerplate/           # Example Excel templates
├── uploads/               # Temporary PDF uploads
├── outputs/               # Generated Excel reports
├── debug_ocr.py           # Offline OCR debugging tool
└── requirements.txt       # Python dependencies
```

## Notes

- Online OCR may incur API charges depending on your Anthropic account usage.
- Test limits and settings can be changed in `config.py`.
- The upload limit is 100 MB per request.
- Temporary uploaded PDFs are deleted automatically after an Excel report is generated successfully.
