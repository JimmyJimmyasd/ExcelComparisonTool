@echo off
echo Starting Excel Comparison Tool...
echo.
echo Dependencies already installed!
echo.
echo Launching application...
echo The app will open in your default browser at http://localhost:8501
echo.
echo Press Ctrl+C to stop the application
echo.
powershell -Command "& 'C:/Users/gamal/AppData/Local/Programs/Python/Python313/python.exe' -m streamlit run app.py"
pause