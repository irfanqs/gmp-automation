# GMP Automation System - Online

[![Korean README](https://img.shields.io/badge/README-한국어-2563eb)](README.kr.md)

A local web application that converts pharmaceutical-facility environmental measurement PDFs into structured Microsoft Excel reports with averages, limit indicators, and charts.

## Features

- Processes multiple PDFs for different AHUs in one request.
- Extracts PDF data through the Anthropic Claude API.
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

- Python 3.10 or later
- Poppler, used to convert PDF pages into images
- An internet connection and an Anthropic API key for OCR

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
git clone --branch gmp-online --single-branch https://github.com/irfanqs/gmp-automation.git
cd gmp-automation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

On Windows, activate the virtual environment with:

```bat
.venv\Scripts\activate
```

## Configure the Anthropic API Key

The Online OCR service requires an Anthropic API key. Create one in the [Anthropic Console](https://console.anthropic.com/settings/keys), then create the local environment file:

```bash
cp .env.example .env
```

Open `.env` and add the key after the equals sign:

```env
ANTHROPIC_API_KEY=your_anthropic_api_key_here
```

Keep `.env` private. It is ignored by Git and must not be committed or shared. The application reads this key on the server, so users do not enter it in the web interface.

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

On the first Windows startup, the script asks for an Anthropic API key and stores it privately in `.env`. The key is not displayed in the terminal or stored in Git.

Open `http://localhost:5001/online` in a browser.

## Docker Deployment

Docker Engine and the Docker Compose plugin are required. Complete the API key setup above, then deploy this branch with:

```bash
bash deploy.sh
```

The API key remains on the server and is not sent to the browser.

The Online OCR application is available at `http://localhost:5001/online`. To use a different host port:

```bash
HOST_PORT=8081 bash deploy.sh
```

Generated files and temporary uploads are stored in Docker volumes. Use the following commands to inspect logs or stop the application:

```bash
docker compose -p gmp-online logs -f
docker compose -p gmp-online down
```

## Usage

1. Select one of the supported test types (A-E).
2. Upload one or more PDFs of the same test type.
3. Click `Start Excel Generation` to generate and download the report.

Each PDF must contain the measurement record for one AHU and one semester. OCR accuracy depends on the quality of the scanned PDF.

## Project Structure

```text
gmp-automation/
├── app.py                 # Flask application and endpoints
├── config.py              # Test limits and environment configuration
├── ocr_engine.py          # PDF extraction through Anthropic Claude
├── excel_generator.py     # Excel report and chart generator
├── templates/index.html   # Web interface template
├── boilerplate/           # Example Excel templates
├── uploads/               # Temporary PDF uploads
├── outputs/               # Generated Excel reports
└── requirements.txt       # Python dependencies
```

## Notes

- OCR requests may incur charges depending on your Anthropic account usage.
- Test limits and settings can be changed in `config.py`.
- The upload limit is 100 MB per request.
- Temporary uploaded PDFs are deleted automatically after an Excel report is generated successfully.
