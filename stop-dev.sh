#!/usr/bin/env sh
set -eu

if ! command -v docker >/dev/null 2>&1; then
  echo "[ERROR] Docker is not installed or not in PATH."
  exit 1
fi

docker compose -f docker-compose.yml -f docker-compose.dev.yml down
echo "Dev services are stopped."
