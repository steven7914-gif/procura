@echo off
cd /d "%~dp0"
where node >nul 2>nul
if errorlevel 1 (
  echo Node.js no esta instalado. Instale Node.js para iniciar el reporte.
  pause
  exit /b 1
)
set ABRIR_REPORTE=1
node servidor.js
pause