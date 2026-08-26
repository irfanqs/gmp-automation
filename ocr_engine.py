"""
GMP Automation System - OCR Engine
Uses Anthropic Claude API to extract structured data from scanned PDF images.
"""

import base64
import json
import requests
import os
from pdf2image import convert_from_path
from io import BytesIO
from config import ANTHROPIC_API_KEY, ANTHROPIC_MODEL, ANTHROPIC_API_URL


def pdf_to_images(pdf_path, dpi=150):
    """Convert PDF pages to PIL Images."""
    images = convert_from_path(pdf_path, dpi=dpi)
    return images


def image_to_base64(pil_image):
    """Convert PIL Image to base64 string."""
    buffer = BytesIO()
    pil_image.save(buffer, format='PNG')
    return base64.standard_b64encode(buffer.getvalue()).decode('utf-8')


def call_claude_api(images_b64, prompt, api_key=None):
    """Call Claude API with images and a prompt. Returns parsed JSON."""
    key = api_key or ANTHROPIC_API_KEY
    if not key:
        raise ValueError("Anthropic API key is required. Set ANTHROPIC_API_KEY in the server environment.")

    content = []
    for img_b64 in images_b64:
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": "image/png",
                "data": img_b64
            }
        })
    content.append({"type": "text", "text": prompt})

    payload = {
        "model": ANTHROPIC_MODEL,
        "max_tokens": 8192,
        "messages": [{"role": "user", "content": content}]
    }

    headers = {
        "Content-Type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01"
    }

    response = requests.post(ANTHROPIC_API_URL, json=payload, headers=headers, timeout=300)

    if response.status_code != 200:
        raise Exception(f"Claude API Error {response.status_code}: {response.text}")

    result = response.json()
    text = ""
    for block in result.get("content", []):
        if block.get("type") == "text":
            text += block["text"]

    # Extract JSON from response
    text = text.strip()
    if text.startswith("```json"):
        text = text[7:]
    if text.startswith("```"):
        text = text[3:]
    if text.endswith("```"):
        text = text[:-3]
    text = text.strip()

    try:
        return json.loads(text)
    except json.JSONDecodeError as e:
        raise Exception(f"Failed to parse Claude response as JSON: {e}\nResponse: {text[:500]}")


# =============================================================================
# OCR PROMPTS FOR EACH TEST TYPE
# =============================================================================

PROMPT_AIRBORNE_PARTICLE = """You are analyzing a scanned Korean GMP document: 부유입자 측정 기록서 (Airborne Particle Test Record).

Extract ALL data and return ONLY valid JSON (no other text) with this exact structure:

{
  "ahu": "the AHU number from 해당 공조기 field (e.g., '33' if it says 공조기-33)",
  "date": "the measurement date from 측정일자 field (e.g., '2025.08.14')",
  "result": "측정결과 value (e.g., '적합')",
  "rooms": [
    {
      "no_start": 1,
      "no_end": 6,
      "grade": "B",
      "room_number": "2142",
      "room_name": "무균 실험실",
      "measurements": [
        {"point": 1, "value_05": 121, "value_50": 7},
        {"point": 2, "value_05": 194, "value_50": 0}
      ]
    }
  ]
}

IMPORTANT RULES:
- Extract EVERY row from the bottom table (Tabel Bawah / 측정값 table)
- "grade" is from 청정등급 column (A, B, C, or D)
- "room_number" is from 실번호 column
- "room_name" is from 실명 column
- "point" is from 측정번호 column
- "value_05" is the 0.5 µm measurement value (integer)
- "value_50" is the 5.0 µm measurement value (integer)
- Group measurements by room (same room_number + room_name = same room object)
- Include ALL pages of data
- For the AHU number, extract only the number (e.g., if it says "공조기-33", return "33")
- Return ONLY the JSON, no markdown, no explanation"""


PROMPT_AIR_VELOCITY = """You are analyzing a scanned Korean GMP document: 풍속 측정 기록서 (Air Velocity Test Record).

Extract ALL data and return ONLY valid JSON (no other text) with this exact structure:

{
  "ahu": "the AHU number (e.g., '33' if it says 공조기-33)",
  "date": "the measurement date from 측정일자 (e.g., '2025.08.02')",
  "result": "측정결과 value",
  "machines": [
    {
      "no_start": 1,
      "no_end": 4,
      "grade": "A",
      "room_number": "2142",
      "machine_name": "무균시험실 BSC\\nBio SafetyCabinet-1(16830)",
      "measurements": [
        {"point": 1, "value": 0.47},
        {"point": 2, "value": 0.44},
        {"point": 3, "value": 0.46},
        {"point": 4, "value": 0.44}
      ]
    }
  ]
}

IMPORTANT RULES:
- Extract EVERY row from the measurement table
- "grade" is from 청정등급 column (A, B, C, or D)
- "room_number" is from 실번호 column
- "machine_name" is from 실명 column. Include the full name with both Korean name and model/code on separate lines using \\n
- "value" is the measurement value in m/s (decimal number)
- Group measurements by machine (same machine_name = same machine object)
- Include ALL pages of data
- Return ONLY the JSON, no markdown, no explanation"""


PROMPT_AIR_CHANGE_RATE = """You are analyzing a scanned Korean GMP document: 환기횟수 측정 기록서 (Air Change Rate Test Record).

Extract ALL data and return ONLY valid JSON (no other text) with this exact structure:

{
  "ahu": "the AHU number (e.g., '33' if it says 공조기-33)",
  "date": "the measurement date from 측정일자 (e.g., '2025.08.02')",
  "result": "측정결과 value",
  "rooms": [
    {
      "no": 1,
      "grade": "B",
      "room_number": "2142",
      "room_name": "무균시험실",
      "volume": 22.4,
      "air_flow_measurements": [
        {"point": 1, "air_flow": 657.8},
        {"point": 2, "air_flow": 760.1}
      ],
      "total_air_flow": 1417.9,
      "ach": 63
    }
  ]
}

IMPORTANT RULES:
- Extract EVERY row from the measurement table
- "grade" is from 청정등급 column (B, C, or D - no A for ACH)
- "room_number" is from 실번호 column
- "room_name" is from 실명 column
- "volume" is from 체적 column (decimal number)
- "air_flow" values are from 풍량(m³/hr) column
- "total_air_flow" is the 합계 value (sum of air_flow values). If only 1 measurement point, total = that single value.
- "ach" is from 환기횟수(회/hr) column (integer)
- Include ALL rows
- Return ONLY the JSON, no markdown, no explanation"""


PROMPT_HEPA_FILTER = """You are analyzing a scanned Korean GMP document: HEPA FILTER 성능 검사 집계표 (HEPA Filter Test Record).

Extract ALL data and return ONLY valid JSON (no other text) with this exact structure:

{
  "ahu": "the AHU number (e.g., '33' if it says 공조기-33)",
  "date": "the measurement date from 측정일자 (e.g., '2025.08.03')",
  "result": "측정결과 value",
  "standard": "측정기준 value (e.g., '0.01%')",
  "items": [
    {
      "no_start": 1,
      "no_end": 1,
      "room_number": "2142",
      "item_name": "무균시험실 BSC",
      "measurements": [
        {"point": 1, "value": 0.003}
      ]
    }
  ]
}

IMPORTANT RULES:
- Extract EVERY row from the measurement table
- "room_number" is from 실번호 column
- "item_name" is from 실명 column
- "value" is the 측정값 percentage value as a NUMBER (e.g., if it shows "0.003%", enter 0.003)
- Do NOT include the % sign in the value - just the number
- Group measurements by item (same room_number + item_name = same item object)
- There is NO 청정등급 column in this test type
- Include ALL pages of data
- Return ONLY the JSON, no markdown, no explanation"""


PROMPT_AIRFLOW_PATTERN = """You are analyzing scanned Korean GMP documents: 기류패턴시험 기록서 (Airflow Pattern Test Records).

Each page is a separate test for one equipment/room. Extract ALL data from ALL pages and return ONLY valid JSON (no other text) with this exact structure:

{
  "ahu": "the AHU number - look for it in the document context or filename",
  "items": [
    {
      "name": "무균시험실 BSC",
      "date": "2025.08.02",
      "criteria": "1. 육안상 단일방향류가 형성되어야 함\\n2. 측정대상 크린장비 내부에 난류가 형성되는 구역이 없어야 함",
      "video_attached": "첨부",
      "judgment": "적합"
    }
  ]
}

IMPORTANT RULES:
- Each page represents one equipment/room test
- "name" is from 측정대상 field (the equipment/room name)
- "date" is from 측정일자 field
- "criteria" is from 측정기준 section
- "video_attached" is from 동영상 첨부 section
- "judgment" is from 판정결과 section (적합 or 부적합)
- Extract data from ALL pages
- Return ONLY the JSON, no markdown, no explanation"""


def extract_airborne_particle(pdf_path, api_key=None):
    """Extract data from Airborne Particle Test PDF."""
    images = pdf_to_images(pdf_path)
    images_b64 = [image_to_base64(img) for img in images]
    return call_claude_api(images_b64, PROMPT_AIRBORNE_PARTICLE, api_key)


def extract_air_velocity(pdf_path, api_key=None):
    """Extract data from Air Velocity Test PDF."""
    images = pdf_to_images(pdf_path)
    images_b64 = [image_to_base64(img) for img in images]
    return call_claude_api(images_b64, PROMPT_AIR_VELOCITY, api_key)


def extract_air_change_rate(pdf_path, api_key=None):
    """Extract data from Air Change Rate Test PDF."""
    images = pdf_to_images(pdf_path)
    images_b64 = [image_to_base64(img) for img in images]
    return call_claude_api(images_b64, PROMPT_AIR_CHANGE_RATE, api_key)


def extract_hepa_filter(pdf_path, api_key=None):
    """Extract data from HEPA Filter Test PDF."""
    images = pdf_to_images(pdf_path)
    images_b64 = [image_to_base64(img) for img in images]
    return call_claude_api(images_b64, PROMPT_HEPA_FILTER, api_key)


def extract_airflow_pattern(pdf_path, api_key=None):
    """Extract data from Airflow Pattern Test PDF."""
    images = pdf_to_images(pdf_path)
    images_b64 = [image_to_base64(img) for img in images]
    return call_claude_api(images_b64, PROMPT_AIRFLOW_PATTERN, api_key)


# Map test types to extraction functions
EXTRACTORS = {
    'airborne_particle': extract_airborne_particle,
    'air_velocity': extract_air_velocity,
    'air_change_rate': extract_air_change_rate,
    'hepa_filter': extract_hepa_filter,
    'airflow_pattern': extract_airflow_pattern,
}
