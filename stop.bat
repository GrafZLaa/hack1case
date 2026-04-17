@echo off
setlocal EnableExtensions

where docker >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Docker is not installed or not in PATH.
  exit /b 1
)

docker compose down
if errorlevel 1 (
  echo [ERROR] Failed to stop services.
  exit /b 1
)

echo Services are stopped.
exit /b 0
