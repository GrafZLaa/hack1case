@echo off
setlocal EnableExtensions

where docker >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Docker is not installed or not in PATH.
  exit /b 1
)

docker compose -f docker-compose.yml -f docker-compose.dev.yml down
if errorlevel 1 (
  echo [ERROR] Failed to stop dev services.
  exit /b 1
)

echo Dev services are stopped.
exit /b 0
