#!/usr/bin/env bash
set -euo pipefail

PROJECT_NAME="gmp-online"

if ! command -v docker >/dev/null 2>&1; then
    echo "Docker is required. Install Docker Engine and the Docker Compose plugin first." >&2
    exit 1
fi

if ! docker compose version >/dev/null 2>&1; then
    echo "Docker Compose plugin is required." >&2
    exit 1
fi

cd "$(dirname "$0")"
docker compose --project-name "$PROJECT_NAME" up --detach --build --remove-orphans
docker compose --project-name "$PROJECT_NAME" ps

echo "GMP online is available at http://localhost:${HOST_PORT:-5001}/online"
