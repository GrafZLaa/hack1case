@echo off
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
for %%I in ("%SCRIPT_DIR%..") do set "PROJECT_ROOT=%%~fI"
set "COMPOSE_BASE=infra\docker-compose.yml"

pushd "%PROJECT_ROOT%" >nul

where docker >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Docker is not installed or not in PATH.
  popd >nul
  exit /b 1
)

docker compose -f "%COMPOSE_BASE%" down
if errorlevel 1 (
  echo [ERROR] Failed to stop services.
  popd >nul
  exit /b 1
)

echo Services are stopped.
popd >nul
exit /b 0
