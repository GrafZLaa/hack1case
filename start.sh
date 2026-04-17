#!/usr/bin/env sh
set -eu

if ! command -v docker >/dev/null 2>&1; then
  echo "[ERROR] Docker is not installed or not in PATH."
  exit 1
fi

if ! docker info >/dev/null 2>&1; then
  echo "[ERROR] Docker daemon is not running."
  exit 1
fi

echo "[1/4] Building and starting containers..."
docker compose up -d --build

ENABLE_LLM="1"
OLLAMA_MODEL="llama3.2-vision"

if [ -f ".env" ]; then
  ENV_ENABLE_LLM="$(grep -E '^ENABLE_LLM=' .env | tail -n1 | cut -d= -f2- | tr -d '\r' | xargs || true)"
  ENV_OLLAMA_MODEL="$(grep -E '^OLLAMA_MODEL=' .env | tail -n1 | cut -d= -f2- | tr -d '\r' | xargs || true)"
  if [ -n "$ENV_ENABLE_LLM" ]; then ENABLE_LLM="$ENV_ENABLE_LLM"; fi
  if [ -n "$ENV_OLLAMA_MODEL" ]; then OLLAMA_MODEL="$ENV_OLLAMA_MODEL"; fi
fi

case "$(printf '%s' "$ENABLE_LLM" | tr '[:upper:]' '[:lower:]')" in
  0|false|no)
    echo "[4/4] Ready. LLM is disabled by .env (ENABLE_LLM=$ENABLE_LLM)."
    echo "Open http://localhost:5000"
    exit 0
    ;;
esac

echo "[2/4] Waiting for Ollama..."
ATTEMPTS=0
while [ "$ATTEMPTS" -lt 30 ]; do
  if docker compose exec -T ollama ollama list >/dev/null 2>&1; then
    break
  fi
  ATTEMPTS=$((ATTEMPTS + 1))
  sleep 2
done

if [ "$ATTEMPTS" -ge 30 ]; then
  echo "[WARN] Ollama did not become ready in time. Skip model pull."
  echo "Run later: docker compose exec ollama ollama pull $OLLAMA_MODEL"
  echo "[4/4] Ready. Open http://localhost:5000"
  exit 0
fi

if [ -z "$OLLAMA_MODEL" ]; then
  OLLAMA_MODEL="llama3.2-vision"
fi

echo "[3/4] Pulling model $OLLAMA_MODEL (if missing)..."
if ! docker compose exec -T ollama ollama pull "$OLLAMA_MODEL"; then
  echo "[WARN] Could not pull model now."
  echo "Run later: docker compose exec ollama ollama pull $OLLAMA_MODEL"
fi

echo "[4/4] Ready. Open http://localhost:5000"
