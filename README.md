# PESU Academy Resource Downloader

Download course materials from PESU Academy with automatic conversion and merging.

## Quick Start

### 1. Clone the Repository
```bash
git clone https://github.com/ilb225112/pesu_course_downloader.git
cd pesu_course_downloader
```

### 2. (Optional) Add Credentials
Create a `.env` file so the downloader doesn't prompt every run:
```
PESU_USERNAME=your_srn
PESU_PASSWORD=your_password
```
If you skip this, the downloader will ask for credentials when it starts.


### 3. Run the Setup Script
The setup script handles everything — Python version check, virtual environment, dependencies, and launch.
 
**Windows:**
```bash
py setup_env.py
```
**Linux / macOS:**
```bash
python3 setup_env.py
```
> **Windows note:** Python 3.11 or 3.12 required (`windows-curses` does not support 3.13+).
> The script will automatically pick the right version if both are installed.
 
That's it. On subsequent runs, `setup_env.py` skips already-installed packages and launches immediately.
 
---


## Manual Setup (alternative)
If you prefer to manage the environment yourself:

<details>
<summary>Expand manual steps</summary>
  
**Windows** (must use 3.11 or 3.12):
```bash
py -3.12 -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python interactive_download.py
```
 
**Linux (Ubuntu/Debian):**
```bash
sudo apt install python3-venv   # one-time
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 interactive_download.py
```
 
**macOS:**
```bash
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 interactive_download.py
```
</details>

---

## 📄 What Each Script Does

### **`interactive_download.py`**  (Main Script) :<br>

  Complete interactive workflow: <br>
  -  Login → Select Course → Select Units → Download → Convert PPTX/DOCX to PDF → Detect & Remove Duplicates → Merge PDFs → Cleanup.
  -  Includes automatic corruption repair for damaged files. **This is the only file you need to run.**
---

### **`setup_env.py`** — Setup & Launcher
- Detects the right Python version and creates a virtual environment
- Pre-checks installed packages — only installs what's missing
- Launches the downloader automatically
- Safe to re-run; skips steps that are already done

### **`pdf_dedup.py`** (Auto-loaded by main script) :
Detects and removes duplicate PDFs after conversion.
- Skips files with different sizes instantly (zero cost)
- Uses perceptual hashing (pHash) only on same-size candidates
- Prompts before deleting — or pass `auto_delete=True` for batch use

**Requires:** `pip install pymupdf`

##  Notes

- **Windows users:** PowerPoint COM provides best conversion quality (requires MS Office installed)
- **Cross-platform:** Use Aspose.Slides or LibreOffice as fallback
- Files are numbered sequentially within each unit for easy merging
- Empty files and temporary data are automatically cleaned up

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to fork and submit a PR.

