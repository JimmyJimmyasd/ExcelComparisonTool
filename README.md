# 📊 Excel Comparison Tool

A powerful web-based application built with Python and Streamlit that allows users to upload two Excel files, compare data between them using exact and fuzzy matching, and export the comparison results.

## ✨ Features

- **📁 File Upload**: Support for `.xlsx` and `.xls` files with multi-sheet selection
- **🔍 Smart Matching**: 
  - Exact matching for identical values
  - Fuzzy matching with configurable similarity threshold (50-100%)
  - Case-insensitive and whitespace-normalized comparison
- **📊 Interactive Results**: 
  - Categorized results: Matched, Suggested Matches, Unmatched
  - Real-time similarity scores
  - Comprehensive data preview
- **📥 Export Functionality**: Download results as Excel with separate sheets for each category
- **⚙️ Customizable Settings**: Adjustable match threshold and text normalization options
- **🎯 User-Friendly Interface**: Clean, intuitive Streamlit-based UI

## 🚀 Quick Start

### Prerequisites

- Python 3.10 or higher
- pip (Python package installer)

### Installation

1. **Clone or download this repository**
   ```bash
   git clone <repository-url>
   cd excel-comparison-tool
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the application**
   ```bash
   streamlit run app.py
   ```

4. **Open your browser**
   - The app will automatically open at `http://localhost:8501`
   - If it doesn't open automatically, navigate to the URL shown in the terminal

## 📦 **Deployment to Another PC**

### **Easy Setup Method:**
1. **Copy entire project folder** to target PC
2. **Install Python 3.10+** with PATH configuration
3. **Run setup:** Double-click `setup.bat` (Windows) or `python launch.py`
4. **Launch app:** `python launch.py` or `streamlit run app.py`

### **Detailed Instructions:**
- See **`DEPLOYMENT_GUIDE.md`** for comprehensive deployment options
- See **`EASY_DEPLOYMENT.md`** for simplified deployment steps
- Includes troubleshooting, network deployment, and Docker options

## 📋 How to Use

### Step 1: Upload Files
1. Use the sidebar to upload your two Excel files (Sheet A and Sheet B)
2. Select the appropriate sheet/tab from each file
3. Preview the data to ensure it loaded correctly

### Step 2: Configure Matching
1. **Select Key Columns**: Choose the columns to use for matching (e.g., Customer Name, ID)
2. **Choose Columns to Extract**: Select which columns from Sheet B to merge into the results
3. **Set Match Threshold**: Adjust the similarity threshold for fuzzy matching (default: 80%)
4. **Configure Options**: Enable case-insensitive matching if needed

### Step 3: Run Comparison
1. Click "🔍 Start Comparison" to begin the matching process
2. View the summary metrics showing match statistics
3. Explore results in the three tabs:
   - **✅ Matched**: Exact and high-confidence fuzzy matches
   - **⚠️ Suggested Matches**: Lower-confidence fuzzy matches for review
   - **❌ Unmatched**: Records with no suitable matches found

### Step 4: Export Results
1. Click "📥 Download Excel" to generate the results file
2. The exported file contains separate sheets for each result category
3. Each row includes the original data plus similarity scores and match information

## 🛠️ Technical Details

### Architecture
- **Frontend**: Streamlit web framework
- **Data Processing**: Pandas for Excel manipulation
- **Fuzzy Matching**: RapidFuzz library for high-performance string matching
- **Export**: XlsxWriter for Excel file generation

### Matching Algorithm
1. **Exact Match**: Direct string comparison (case-insensitive option)
2. **Fuzzy Match**: Uses Levenshtein distance-based similarity scoring
3. **Threshold-based Categorization**: 
   - ≥90% similarity → Matched
   - Threshold-89% → Suggested
   - <Threshold → Unmatched

### Performance Considerations
- Optimized for files up to 10,000 rows
- Memory-efficient processing with pandas
- Progress indicators for large datasets
- Chunked processing for very large files (future enhancement)

## 📁 Project Structure

```
excel-comparison-tool/
├── app.py                     # Main Streamlit application
├── utils.py                   # Utility functions and validation
├── create_sample_data.py      # Generate test Excel files
├── requirements.txt           # Python dependencies
├── README.md                  # This file
└── sample_files/              # Test files (generated)
    ├── sample_customers.xlsx
    └── sample_orders.xlsx
```

## 🧪 Testing

### Generate Sample Data
Create test Excel files to try the application:

```bash
python create_sample_data.py
```

This creates:
- `sample_customers.xlsx`: Customer data with multiple sheets
- `sample_orders.xlsx`: Order data with fuzzy matching scenarios

### Test Scenarios
The sample data includes:
- **Exact Matches**: Identical customer names
- **Fuzzy Matches**: Similar names with typos (Robert vs Bob, Charles vs Charlie)
- **Unmatched Records**: Names that don't exist in the other file
- **Multiple Sheets**: Test sheet selection functionality

## 🚀 Deployment Options

### Option 1: Local Development
```bash
streamlit run app.py
```

### Option 2: Streamlit Cloud (Free)
1. Push your code to GitHub
2. Visit [share.streamlit.io](https://share.streamlit.io)
3. Connect your GitHub repository
4. Deploy with one click

### Option 3: Docker Deployment
Create a `Dockerfile`:
```dockerfile
FROM python:3.10-slim

WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
```

Build and run:
```bash
docker build -t excel-comparison-tool .
docker run -p 8501:8501 excel-comparison-tool
```

### Option 4: Azure Web App
1. Create an Azure Web App for Python
2. Upload your code
3. Configure startup command: `python -m streamlit run app.py --server.port=8000 --server.address=0.0.0.0`

## 🔧 Configuration

### Environment Variables
- `STREAMLIT_SERVER_PORT`: Port number (default: 8501)
- `STREAMLIT_SERVER_ADDRESS`: Host address (default: localhost)
- `STREAMLIT_THEME_BASE`: UI theme (light/dark)

### Advanced Settings
Edit `config.toml` in `.streamlit/` folder:
```toml
[server]
maxUploadSize = 200

[theme]
primaryColor = "#FF6B6B"
backgroundColor = "#FFFFFF"
```

## 🐛 Troubleshooting

### Common Issues

**Issue**: "Import could not be resolved" errors
- **Solution**: Install dependencies with `pip install -r requirements.txt`

**Issue**: File upload fails
- **Solution**: Check file size (<200MB) and format (.xlsx/.xls only)

**Issue**: Fuzzy matching is slow
- **Solution**: Reduce dataset size or increase threshold for faster processing

**Issue**: Memory errors with large files
- **Solution**: Process files in smaller chunks or increase system memory

### Performance Tips
- Use exact matching when possible (faster than fuzzy)
- Set appropriate thresholds (higher = faster)
- Clean data before uploading (remove empty rows/columns)
- Use meaningful column names for better UX

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

If you encounter any issues or have questions:
1. Check the [Troubleshooting](#🐛-troubleshooting) section
2. Create an issue on GitHub
3. Review the sample data and test scenarios

## 🔮 Future Enhancements

- [ ] Multi-column matching (First Name + Last Name)
- [ ] CSV file support
- [ ] Database integration for historical comparisons
- [ ] Advanced highlighting of differences
- [ ] Batch processing for multiple file pairs
- [ ] API endpoints for programmatic access
- [ ] Custom matching algorithms
- [ ] Data validation and cleaning tools

---

**Made with ❤️ using Python and Streamlit**