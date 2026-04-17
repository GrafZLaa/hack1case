@echo off
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "PROJECT_ROOT=%%~fI"
set "COMPOSE_BASE=infra\docker-compose.yml"
set "COMPOSE_DEV=infra\docker-compose.dev.yml"

pushd "%PROJECT_ROOT%" >nul

where docker >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Docker is not installed or not in PATH.
  popd >nul
  exit /b 1
)

docker info >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Docker daemon is not running. Start Docker Desktop and try again.
  popd >nul
  exit /b 1
)

echo [1/4] Starting DEV mode with hot-reload...
docker compose -f "%COMPOSE_BASE%" -f "%COMPOSE_DEV%" up -d --build
if errorlevel 1 (
  echo [ERROR] Failed to start docker compose services in dev mode.
  popd >nul
  exit /b 1
)

set "ENABLE_LLM=1"
set "OLLAMA_MODEL=llama3.2-vision"

if exist ".env" (
  for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    set "K=%%~A"
    set "V=%%~B"
    if /i "!K!"=="ENABLE_LLM" set "ENABLE_LLM=!V!"
    if /i "!K!"=="OLLAMA_MODEL" set "OLLAMA_MODEL=!V!"
  )
)

if /i "!ENABLE_LLM!"=="0" goto done
if /i "!ENABLE_LLM!"=="false" goto done
if /i "!ENABLE_LLM!"=="no" goto done

echo [2/4] Waiting for Ollama...
set /a ATTEMPTS=0
:wait_ollama
docker compose -f "%COMPOSE_BASE%" -f "%COMPOSE_DEV%" exec -T ollama ollama list >nul 2>&1
if not errorlevel 1 goto pull_model
set /a ATTEMPTS+=1
if !ATTEMPTS! GEQ 30 goto ollama_timeout
timeout /t 2 >nul
goto wait_ollama

:pull_model
if "!OLLAMA_MODEL!"=="" set "OLLAMA_MODEL=llama3.2-vision"
echo [3/4] Pulling model !OLLAMA_MODEL! (if missing)...
docker compose -f "%COMPOSE_BASE%" -f "%COMPOSE_DEV%" exec -T ollama ollama pull !OLLAMA_MODEL!
if errorlevel 1 (
  echo [WARN] Could not pull model now. You can run it later manually:
  echo        docker compose -f "%COMPOSE_BASE%" -f "%COMPOSE_DEV%" exec ollama ollama pull !OLLAMA_MODEL!
)
goto done

:ollama_timeout
echo [WARN] Ollama did not become ready in time. Skip model pull.
echo        You can run later:
echo        docker compose -f "%COMPOSE_BASE%" -f "%COMPOSE_DEV%" exec ollama ollama pull !OLLAMA_MODEL!

:done
echo [4/4] DEV mode is ready. Code changes reload automatically.
echo        Open http://localhost:5000
popd >nul
exit /b 0
