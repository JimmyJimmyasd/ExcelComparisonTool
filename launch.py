#!/usr/bin/env python3
"""
Universal launcher for Excel Comparison Tool
Works on Windows, Mac, and Linux
"""
import subprocess
import sys
import os
import platform

def check_python():
    """Check if Python is available and get version"""
    try:
        version = sys.version_info
        if version.major >= 3 and version.minor >= 10:
            print(f"✅ Python {version.major}.{version.minor}.{version.micro} found")
            return True
        else:
            print(f"❌ Python {version.major}.{version.minor} found, but 3.10+ required")
            return False
    except:
        print("❌ Python not found")
        return False

def install_dependencies():
    """Install required packages"""
    print("📦 Installing dependencies...")
    requirements_files = ['requirements.txt', 'requirements_simple.txt']
    
    for req_file in requirements_files:
        if os.path.exists(req_file):
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', req_file])
                print(f"✅ Dependencies installed from {req_file}")
                return True
            except subprocess.CalledProcessError:
                print(f"⚠️ Failed to install from {req_file}, trying next...")
                continue
    
    # Fallback: install packages individually
    packages = ['streamlit', 'pandas', 'openpyxl', 'rapidfuzz', 'numpy', 'xlsxwriter']
    for package in packages:
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            print(f"✅ Installed {package}")
        except subprocess.CalledProcessError:
            print(f"⚠️ Failed to install {package}")
    
    return True

def run_app():
    """Launch the Streamlit app"""
    print("🚀 Launching Excel Comparison Tool...")
    print("📍 App will be available at: http://localhost:8501")
    print("🛑 Press Ctrl+C to stop the application")
    print("-" * 50)
    
    try:
        subprocess.run([sys.executable, '-m', 'streamlit', 'run', 'app.py'])
    except KeyboardInterrupt:
        print("\n👋 Application stopped by user")
    except FileNotFoundError:
        print("❌ Streamlit not found. Please install dependencies first.")
    except Exception as e:
        print(f"❌ Error launching app: {e}")

def main():
    """Main launcher function"""
    print("🎯 Excel Comparison Tool - Universal Launcher")
    print("=" * 50)
    print(f"💻 Operating System: {platform.system()} {platform.release()}")
    print(f"📂 Working Directory: {os.getcwd()}")
    print("")
    
    # Check Python
    if not check_python():
        print("\n❌ Please install Python 3.10+ and try again")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Check if app.py exists
    if not os.path.exists('app.py'):
        print("❌ app.py not found in current directory")
        print("📂 Please navigate to the Excel Comparison Tool folder")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Install dependencies
    print("\n" + "=" * 50)
    install_dependencies()
    
    # Run the app
    print("\n" + "=" * 50)
    run_app()

if __name__ == "__main__":
    main()