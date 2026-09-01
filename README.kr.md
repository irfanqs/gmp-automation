# GMP Automation System - Offline (오프라인 OCR)

[![English README](https://img.shields.io/badge/README-English-2563eb)](README.md)

제약 및 바이오 시설의 환경측정 PDF 기록서를 분석하여 평균값, 기준선(경고/조치), 차트가 포함된 Microsoft Excel 보고서로 자동 변환하는 로컬 웹 애플리케이션입니다.

## 주요 기능

- **다중 PDF 일괄 처리:** 여러 AHU의 PDF 기록서를 한 번에 업로드하여 처리합니다.
- **로컬 DeepSeek OCR:** 셀프 호스팅 DeepSeek OCR 엔드포인트를 통해 PDF 데이터를 추출합니다.
- **자동 Excel 보고서 생성:** 시험 항목별로 데이터 시트, 요약 테이블, 피벗/차트 시트를 자동으로 생성합니다.
- **기준 초과 시각화:** 설정된 경고/조치 기준을 벗어난 측정값을 빨간색 하이라이트로 표시합니다.
- **AHU 및 반기별 자동 정렬:** 공조기(AHU) 번호와 측정 반기(상/하반기) 기준 데이터 자동 그룹화.

## 지원되는 측정 항목

| 코드 | 측정 항목 | 생성 Excel 파일명 |
| --- | --- | --- |
| A | 부유입자 측정 (Airborne Particle Test) | `Airborne_Particle_Test_Result_and_Graph.xlsx` |
| B | 풍속 측정 (Air Velocity Test) | `Air_Velocity_Test_Result_and_Graph.xlsx` |
| C | 환기횟수 측정 (Air Change Rate Test) | `Air_Change_Rate_Test_Result_and_Graph.xlsx` |
| D | HEPA 필터 검사 (HEPA Filter Test) | `HEPA_Filter_Test_Result_and_Graph.xlsx` |
| E | 기류패턴 시험 (Airflow Pattern Test) | `Airflow_Pattern_Test_Result_and_Graph.xlsx` |

## 사전 요구 사항

- Python 3.10 이상
- Poppler (PDF 페이지를 이미지로 변환하기 위해 필요)
- Docker Desktop은 선택 사항이며 컨테이너 배포에만 필요합니다.
- 최소 16GB 시스템 RAM. GTX 1050 호환성을 위해 로컬 모델은 CPU에서 실행되므로 처리 속도가 느릴 수 있습니다.

### Poppler 설치 방법:

```bash
# macOS
brew install poppler

# Ubuntu/Debian
sudo apt install poppler-utils
```

Windows의 경우 [poppler-windows](https://github.com/oschwartz10612/poppler-windows/releases)를 다운로드하여 `C:\poppler`에 압축 해제 후, `C:\poppler\Library\bin` 경로를 System `PATH` 환경 변수에 추가하세요.

## 설치 방법

```bash
git clone --branch gmp-offline --single-branch https://github.com/irfanqs/gmp-automation.git
cd gmp-automation
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows 환경에서는 아래 명령어로 가상환경을 활성화합니다:

```bat
.venv\Scripts\activate
```

## Windows GPU 호환성 확인

로컬 DeepSeek OCR을 설정하기 전에 `CHECK_HARDWARE.bat`를 실행하세요. 하드웨어 정보를 외부로 전송하지 않고 디스플레이 어댑터, NVIDIA 드라이버 설치 여부, VRAM을 확인합니다.

현재 배포는 CPU 추론을 사용하므로 지원되는 GPU 없이도 실행할 수 있습니다. 이 스크립트는 향후 CUDA 가속을 위해 NVIDIA 드라이버와 VRAM 정보를 표시하며, GTX 1050 및 내장 GPU는 계속 CPU 모드를 사용합니다.

## Docker 없이 Windows에서 실행

`START_WINDOWS.bat`를 실행하세요. 이 스크립트는 `.venv`를 생성하고 CPU OCR 의존성을 설치한 뒤 로컬 DeepSeek OCR 모델을 시작합니다. 모델 로딩이 완료되면 `http://localhost:5002/offline`에서 GMP 웹 애플리케이션을 실행합니다.

모델은 별도의 `GMP Offline OCR Model` 창에서 시작됩니다. 애플리케이션을 사용하는 동안 이 창을 닫지 마세요. 처음 실행할 때는 모델을 다운로드하므로 CPU 환경에서 시간이 오래 걸릴 수 있으며, 이후 실행부터는 다운로드된 모델 캐시를 재사용합니다.

## Docker 배포

Windows에서는 Docker Desktop, 그 외 환경에서는 Docker Engine 및 Docker Compose 플러그인이 필요합니다. 아래 명령어로 이 브랜치를 배포하세요:

```bash
bash deploy.sh
```

처음 배포할 때 Hugging Face에서 DeepSeek OCR 모델을 다운로드하여 CPU 메모리에 로드합니다. 이 과정은 시간이 오래 걸릴 수 있으며 최초 다운로드 시에만 인터넷 연결이 필요합니다. Docker는 모델을 영구 볼륨에 저장하므로 이후 실행에서는 다시 다운로드하지 않습니다.

`ocr` 서비스가 정상 상태가 되면 `http://localhost:5002/offline`에서 Offline OCR 애플리케이션에 접속할 수 있습니다. 웹 인터페이스에 URL이나 API Key를 입력할 필요가 없습니다. 최초 모델 다운로드 및 로딩 진행 상황은 다음 명령어로 확인하세요:

```bash
docker compose -p gmp-offline logs -f ocr
```

호스트 포트를 변경해야 하는 경우:

```bash
HOST_PORT=8082 bash deploy.sh
```

생성된 Excel 파일과 임시 업로드 파일은 Docker 볼륨에 저장됩니다. 로그 확인 및 정지는 아래 명령어를 사용하세요:

```bash
docker compose -p gmp-offline logs -f
docker compose -p gmp-offline down
```

## 사용 방법

1. 로컬 `ocr` 서비스의 모델 로딩이 완료될 때까지 기다립니다.
2. 측정할 시험 항목(A~E) 중 하나를 선택합니다.
3. 동일한 시험 항목의 PDF 파일을 하나 이상 업로드합니다.
4. `Excel 자동 생성 시작` 버튼을 클릭하여 보고서를 생성하고 다운로드합니다.

*각 PDF 파일은 1개 AHU의 1개 반기 측정 기록서여야 합니다. OCR 인식률은 스캔된 PDF의 화질에 영향을 받습니다.*

## 로컬 CPU OCR

현재 설정은 호환되는 CUDA GPU 없이도 실행할 수 있도록 의도적으로 CPU 추론을 사용합니다. GTX 1050, Intel/AMD 내장 그래픽 및 기타 지원되지 않는 GPU는 추론에 사용되지 않습니다. 특히 여러 페이지로 구성된 PDF를 처리할 때 CPU OCR 속도가 느릴 수 있습니다.

최초 모델 다운로드 시에는 인터넷 연결이 필요합니다. Docker에 모델이 캐시된 후에는 Kaggle, ngrok 또는 외부 OCR API를 사용하지 않습니다. 추출 결과를 검토해야 할 때는 `debug_ocr.py`로 원본 OCR 텍스트를 확인할 수 있습니다:

```bash
python debug_ocr.py <PDF_파일_경로>
```

## 프로젝트 구조

```text
gmp-automation/
├── app.py                 # Flask 웹 애플리케이션 및 엔드포인트
├── config.py              # 시험별 기준값 및 환경 설정
├── deepseek_ocr/          # 로컬 OCR 클라이언트, 파서 및 모델 서버
├── excel_generator.py     # Excel 보고서 및 차트 자동 생성 모듈
├── templates/index.html   # 웹 인터페이스 템플릿
├── boilerplate/           # Excel 양식 템플릿
├── uploads/               # PDF 업로드 임시 저장소
├── outputs/               # 생성된 Excel 보고서 저장소
├── debug_ocr.py           # Offline OCR 디버깅 도구
└── requirements.txt       # Python 의존성 패키지 목록
```

## 참고 사항

- 최초 모델 다운로드는 시간이 오래 걸리고 많은 디스크 공간을 사용할 수 있습니다.
- 각 시험 항목별 기준값 및 설정은 `config.py`에서 변경할 수 있습니다.
- 1회 업로드 용량 제한은 100MB입니다.
- Excel 보고서가 성공적으로 생성되면 업로드된 임시 PDF 파일은 자동으로 삭제됩니다.
