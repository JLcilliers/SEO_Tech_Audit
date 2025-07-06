# SEO Tech Audit Processor

An automated tool that analyzes Screaming Frog exports and creates comprehensive technical SEO audit reports with just a few clicks.

![Version](https://img.shields.io/badge/version-1.0-green)
![Platform](https://img.shields.io/badge/platform-Windows-blue)
![License](https://img.shields.io/badge/license-MIT-brightgreen)

## 🚀 Features

- **Automated Analysis**: Processes 50+ SEO metrics from Screaming Frog exports
- **Pass/Fail Grading**: Automatic assessment based on SEO best practices
- **Excel Integration**: Imports existing Excel files as additional tabs
- **Client-Ready Reports**: Professional reports with client names and timestamps
- **No Installation Required**: Just download and run!

## 📥 Download

**[Download Tech Audit Processor v1.0](https://github.com/JLc111ers/SEO_Tech_Audit/releases/latest)**

Just download the .exe file from the releases page - no installation needed!

## 🎯 Quick Start

1. **Export from Screaming Frog**
   - Run your website crawl
   - Go to `File → Export → Bulk Export → All Export Options`
   - Save all files to a single folder

2. **Run Tech Audit Processor**
   - Double-click `Tech Audit Processor.exe`
   - Enter your client's name
   - Select the folder with your exports
   - Click "Run Tech Audit"

3. **Get Your Report**
   - Find it on your Desktop
   - File format: `ClientName_Technical_Audit_YYYYMMDD_HHMMSS.xlsx`

## 📊 What Gets Analyzed

### Indexation & Crawlability
- Sitemap issues
- Canonical problems
- Robots.txt blocks
- Noindex/nofollow pages

### On-Page SEO
- Missing/duplicate page titles
- Title length issues
- Missing/duplicate meta descriptions
- Meta description length issues
- Missing/duplicate H1 tags

### Technical Issues
- 404 errors
- Server errors (5xx)
- Redirect chains
- Temporary redirects
- Mixed content issues

### Images
- Missing alt text
- Broken images
- Large image files (>100KB)

### And much more!

## 💻 System Requirements

- Windows 7/8/10/11
- Screaming Frog SEO Spider (free or paid version)
- Excel or compatible spreadsheet software to view reports

## 📁 Expected Folder Structure

Your Screaming Frog export folder should contain:
```
📁 Client Export Folder/
  ├── internal_all.csv
  ├── external_all.csv
  ├── response_codes_all.csv
  ├── page_titles_all.csv
  ├── meta_descriptions_all.csv
  ├── h1_all.csv
  ├── images_all.csv
  ├── canonical_all.csv
  └── Any Excel files (will be imported as tabs)

## 🔧 For Developers

If you want to modify or build from source:

### Prerequisites
- Python 3.8+
- Required packages: `pandas`, `openpyxl`, `tkinter`

### Building from Source
```bash
# Clone the repository
git clone https://github.com/JLc111ers/SEO_Tech_Audit.git
cd SEO_Tech_Audit

# Install dependencies
pip install -r requirements.txt

# Run the script
python tech_audit.py

# Build executable
pyinstaller --onefile --noconsole --add-data "Template __ Tech Audit.xlsx;." --name "Tech Audit Processor" tech_audit.py
```

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## 🐛 Issues

Found a bug or have a feature request? Please open an [issue](https://github.com/JLc111ers/SEO_Tech_Audit/issues).
