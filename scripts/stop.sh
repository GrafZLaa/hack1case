#!/usr/bin/env sh
set -eu

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
PROJECT_ROOT="$(CDPATH= cd -- "$SCRIPT_DIR/.." && pwd)"
COMPOSE_BASE="infra/docker-compose.yml"

cd "$PROJECT_ROOT"

if ! command -v docker >/dev/null 2>&1; then
  echo "[ERROR] Docker is not installed or not in PATH."
  exit 1
fi

docker compose -f "$COMPOSE_BASE" down
echo "Services are stopped."
