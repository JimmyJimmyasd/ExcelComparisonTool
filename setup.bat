@echo off
echo ===============================================
echo    Excel Comparison Tool - Setup Script
echo ===============================================
echo.

echo Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python not found!
    echo.
    echo 📥 Please install Python 3.10+ from:
    echo    https://www.python.org/downloads/
    echo.
    echo ⚠️  IMPORTANT: Check "Add Python to PATH" during installation
    echo.
    pause
    exit /b 1
)

python --version
echo ✅ Python found!
echo.

echo 📦 Installing dependencies...
pip install -r requirements_simple.txt
if %errorlevel% neq 0 (
    echo ⚠️ Standard installation failed, trying alternative...
    pip install streamlit pandas openpyxl rapidfuzz numpy xlsxwriter
)
echo.

echo 🧪 Creating sample test files...
python create_sample_data.py
echo.

echo ✅ Setup complete!
echo.
echo 🚀 To run the app:
echo    1. Double-click 'run_app.bat'
echo    2. Or run: python launch.py
echo    3. Or run: streamlit run app.py
echo.
echo 🌐 App will be available at: http://localhost:8501
echo.
pause