@echo off
REM ===================================================================
REM  Forecast Lab - update and publish the dashboard (one click)
REM  1) picks the most recent workbook in output_forecast_lab
REM  2) rebuilds index.html
REM  3) commits and pushes to GitHub (GitHub Pages republishes)
REM ===================================================================
cd /d "%~dp0"
python gerar_dashboard.py
if errorlevel 1 (
  echo.
  echo *** Dashboard generation failed. Nothing was published. ***
  pause
  exit /b 1
)
git add index.html
git commit -m "Dashboard update %date% %time%"
git push
echo.
echo Done. GitHub Pages will refresh in 1-2 minutes.
pause
