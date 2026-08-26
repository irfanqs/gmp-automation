# GMP Automation System - Online

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

- Python 3.10 or later.
- Poppler, used to convert PDF pages into images.
- An internet connection and an Anthropic API key for OCR.

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

## Configure Anthropic API Key

The online OCR service requires an Anthropic API key. Create one in the [Anthropic Console](https://console.anthropic.com/settings/keys), then create the local environment file:

```bash
cp .env.example .env
```

Open `.env` in an editor and add the key after the equals sign:

```env
ANTHROPIC_API_KEY=your_anthropic_api_key_here
```

Keep `.env` private. It is ignored by Git and must not be committed or shared. The application reads this key on the server, so users do not enter it in the web interface.

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

Open `http://localhost:5001/online` in a browser.

## Docker Deployment

Docker Engine with the Docker Compose plugin is required. Complete the API-key setup above, then deploy this branch with:

```bash
bash deploy.sh
```

The API key stays on the server and is not sent to the browser.

The online application is available at `http://localhost:5001/online`; override the host port when needed:

```bash
HOST_PORT=8081 bash deploy.sh
```

Generated files and temporary uploads are stored in Docker volumes. To inspect the deployment or stop it:

```bash
docker compose -p gmp-online logs -f
docker compose -p gmp-online down
```

## Usage

1. Select one test type.
2. Upload one or more PDFs of the same test type.
3. Generate and download the Excel report.

Each PDF should contain one AHU for one semester. Result accuracy depends on the readability of the scanned PDF.

## Project Structure

```text
gmp_automation/
├── app.py                 # Flask application and upload/download endpoints
├── config.py              # Limits and test-type configuration
├── ocr_engine.py          # PDF extraction through Anthropic Claude
├── excel_generator.py     # Excel report and chart generation
├── templates/index.html   # Web interface
├── boilerplate/           # Example Excel templates
├── uploads/               # Temporary PDF storage during processing
├── outputs/               # Generated Excel reports
└── requirements.txt       # Python dependencies
```

## Notes

- OCR requests incur charges according to Anthropic account usage.
- Limits and Excel settings are configured in `config.py`.
- The application limits each upload request to 100 MB.
- Uploaded files are removed after a successful Excel report is generated.
