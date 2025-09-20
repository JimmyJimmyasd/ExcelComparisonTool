# 🚀 Deployment Guide - Running Excel Comparison Tool on Another PC

## 📋 **Prerequisites for Target PC**

### **Required Software:**
1. **Python 3.10 or higher** (3.13 recommended)
   - Download from: https://www.python.org/downloads/
   - ⚠️ **IMPORTANT**: During installation, check "Add Python to PATH"
2. **pip** (usually comes with Python)
3. **Web browser** (Chrome, Firefox, Edge, Safari)

### **System Requirements:**
- **OS**: Windows 10/11, macOS, or Linux
- **RAM**: 4GB minimum (8GB recommended)
- **Storage**: 500MB free space
- **Internet**: Required for initial package installation

---

## 📦 **Method 1: Complete File Transfer (Recommended)**

### **Step 1: Copy Project Files**
Transfer the entire `APP` folder to the target PC, including:
```
APP/
├── app.py
├── requirements.txt
├── utils.py
├── create_sample_data.py
├── README.md
├── START_HERE.md
├── DEPLOYMENT_GUIDE.md
├── run_app.bat
├── run_app.ps1
├── sample_customers.xlsx
├── sample_orders.xlsx
└── .streamlit/
    └── config.toml
```

### **Step 2: Install Dependencies**
Open **Command Prompt** or **PowerShell** in the APP folder and run:
```bash
pip install -r requirements.txt
```

### **Step 3: Run the Application**
Choose any of these methods:

**Option A: Batch File (Windows)**
```bash
# Double-click run_app.bat or run in cmd:
run_app.bat
```

**Option B: PowerShell Script (Windows)**
```powershell
# Right-click run_app.ps1 → "Run with PowerShell" or:
powershell -ExecutionPolicy Bypass -File run_app.ps1
```

**Option C: Direct Python Command**
```bash
python -m streamlit run app.py
```

**Option D: If Python isn't in PATH**
```bash
# Find Python installation path and use full path:
"C:\Path\To\Python\python.exe" -m streamlit run app.py
```

---

## 🌐 **Method 2: Online Deployment (No Installation)**

### **Option A: Streamlit Cloud (Free)**
1. **Upload to GitHub:**
   - Create a GitHub repository
   - Upload all project files
   
2. **Deploy on Streamlit Cloud:**
   - Go to https://share.streamlit.io
   - Connect your GitHub account
   - Select your repository
   - Click "Deploy"
   - Share the public URL with users

### **Option B: Heroku (Free Tier Available)**
Create these additional files:

**`Procfile`:**
```
web: streamlit run app.py --server.port=$PORT --server.address=0.0.0.0
```

**`runtime.txt`:**
```
python-3.11.9
```

Then deploy via Heroku CLI or GitHub integration.

---

## 🐳 **Method 3: Docker Deployment (Advanced)**

### **Create Dockerfile:**
```dockerfile
FROM python:3.11-slim

WORKDIR /app

# Copy requirements first for better caching
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application files
COPY . .

# Expose port
EXPOSE 8501

# Health check
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

# Run the application
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

### **Build and Run:**
```bash
# Build image
docker build -t excel-comparison-tool .

# Run container
docker run -p 8501:8501 excel-comparison-tool
```

---

## 🔧 **Troubleshooting Common Issues**

### **Issue 1: "Python not found" or "pip not found"**
**Solution:**
- Reinstall Python with "Add to PATH" checked
- Or find Python path: `where python` (Windows) / `which python` (Mac/Linux)
- Use full path: `"C:\Python\python.exe" -m pip install -r requirements.txt`

### **Issue 2: Permission Errors**
**Solution:**
```bash
# Windows (Run as Administrator):
pip install --user -r requirements.txt

# Mac/Linux:
pip install --user -r requirements.txt
# or
sudo pip install -r requirements.txt
```

### **Issue 3: Package Installation Fails**
**Solution:**
```bash
# Update pip first:
python -m pip install --upgrade pip

# Install packages one by one:
pip install streamlit
pip install pandas
pip install openpyxl
pip install rapidfuzz
pip install numpy
```

### **Issue 4: Port Already in Use**
**Solution:**
```bash
# Use different port:
streamlit run app.py --server.port=8502

# Or kill existing process (Windows):
netstat -ano | findstr :8501
taskkill /PID <process_id> /F
```

### **Issue 5: Firewall/Network Issues**
**Solution:**
- Allow Python/Streamlit through Windows Firewall
- For corporate networks, use: `streamlit run app.py --server.address=0.0.0.0`

---

## 📱 **Method 4: Portable Version (No Installation Required)**

### **Create Portable Package:**
1. **On source PC, create virtual environment:**
```bash
python -m venv excel_app_env
excel_app_env\Scripts\activate  # Windows
# or
source excel_app_env/bin/activate  # Mac/Linux

pip install -r requirements.txt
```

2. **Copy entire environment:**
   - Copy `excel_app_env` folder with your app files
   - Transfer to target PC

3. **Run on target PC:**
```bash
excel_app_env\Scripts\activate
python -m streamlit run app.py
```

---

## 🌍 **Network Deployment (Multiple Users)**

### **Run as Network Service:**
```bash
# Allow external connections:
streamlit run app.py --server.address=0.0.0.0 --server.port=8501

# Users access via:
http://YOUR_PC_IP:8501
```

### **Find Your IP Address:**
```bash
# Windows:
ipconfig

# Mac/Linux:
ifconfig
```

---

## 📋 **Quick Setup Checklist for Target PC**

- [ ] Python 3.10+ installed with PATH configured
- [ ] All project files copied to target PC
- [ ] Dependencies installed: `pip install -r requirements.txt`
- [ ] Firewall allows Python/Streamlit (if needed)
- [ ] Test run: `python -m streamlit run app.py`
- [ ] Browser opens to `http://localhost:8501`
- [ ] Sample files work correctly

---

## 🆘 **Emergency Backup Plan**

If nothing works, use **Anaconda/Miniconda**:

1. **Install Anaconda:** https://www.anaconda.com/download
2. **Create environment:**
```bash
conda create -n excel_app python=3.11
conda activate excel_app
conda install -c conda-forge streamlit pandas openpyxl numpy
pip install rapidfuzz
```
3. **Run app:**
```bash
streamlit run app.py
```

---

## 🎯 **Best Practices for Deployment**

1. **Test on similar systems first**
2. **Include all dependencies in requirements.txt**
3. **Document Python version used**
4. **Provide multiple installation methods**
5. **Include troubleshooting steps**
6. **Test with sample data**

---

**📞 Support**: If deployment fails, check the troubleshooting section or contact the original developer with system details (OS, Python version, error messages).