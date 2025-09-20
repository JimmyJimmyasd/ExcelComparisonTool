#!/usr/bin/env pwsh
# Excel Comparison Tool Launcher

Write-Host "Starting Excel Comparison Tool..." -ForegroundColor Green
Write-Host ""
Write-Host "Dependencies already installed!" -ForegroundColor Yellow
Write-Host ""
Write-Host "Launching application..." -ForegroundColor Cyan
Write-Host "The app will open in your default browser at http://localhost:8501" -ForegroundColor White
Write-Host ""
Write-Host "Press Ctrl+C to stop the application" -ForegroundColor Red
Write-Host ""

& "C:/Users/gamal/AppData/Local/Programs/Python/Python313/python.exe" -m streamlit run app.py

Write-Host ""
Write-Host "Application stopped. Press any key to exit..." -ForegroundColor Yellow
Read-Host