# URL to DOCX Web Scraper

Scrapes structured Tamil devotional content from a website into a formatted Word (.docx) document with a Table of Contents and per-chapter sections.

---

## Requirements

- **Python 3.10 or higher** (tested on 3.13.3)
- **Google Chrome** installed on the local PC
- Windows 10/11

---

## Installation

### 1. Clone or download this project

Place the project folder anywhere on your PC, e.g.:
```
D:\koyil.org\kaink1_tiruvirutham_ebook\
```

### 2. Create and activate a virtual environment (recommended)

```powershell
cd "D:\koyil.org\kaink1_tiruvirutham_ebook"
python -m venv venv
.\venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
pip install -r requirements.txt
```

### 4. Install Playwright's Chrome browser

```powershell
playwright install chromium
```

---

## Usage

### Generate DOCX from website

```powershell
python src\scraper.py --url "https://example.com/page" --output "D:\output\myfile.docx" --title "My Document Title"
```

**Parameters:**
- `--url` (required): Full URL of the page to scrape (http/https only)
- `--output` (required): Full file path for output document (must end in .docx)
- `--title` (required): Document title (used in header and table of contents)

**Example:**
```powershell
python src\scraper.py `
  --url "https://divyaprabandham.koyil.org/index.php/2020/11/thiruviruththam/" `
  --output "D:\output\priya_tirumozi.docx" `
  --title "SrI Kural - PriAya TiruviruththAm"
```

### Convert DOCX to PDF

After generating the .docx file, convert it to PDF using:

```powershell
.\src\convert-docx-to-pdf.ps1 -InputPath "D:\output\myfile.docx"
```

**Parameters:**
- `-InputPath` (required): Full path to the .docx file
- `-OutputPath` (optional): Full path for output .pdf file (defaults to same directory/name with .pdf extension)
- `-OpenPDF` (optional): Boolean to open PDF after conversion (default: $false)

**Examples:**
```powershell
# Convert to PDF in same directory
.\src\convert-docx-to-pdf.ps1 -InputPath "D:\output\priya_tirumozi.docx"

# Specify custom output location
.\src\convert-docx-to-pdf.ps1 -InputPath "D:\output\priya_tirumozi.docx" -OutputPath "C:\pdfs\priya_tirumozi.pdf"

# Convert and open PDF
.\src\convert-docx-to-pdf.ps1 -InputPath "D:\output\priya_tirumozi.docx" -OpenPDF $true
```

**Requirements for PDF conversion:**
- Microsoft Word must be installed
- Script requires Windows 10/11

### Arguments

| Argument | Description |
|----------|-------------|
| `--url`  | Full URL of the page to scrape (required) |
| `--output` | Full file path with `.docx` extension for the output file (required) |

---

## How It Works

1. Opens the URL in a Chromium browser (headed)
2. Extracts clean readable content (Reader Mode equivalent)
3. Starts scraping from one line **before** `"varavaramunayE nama:"`
4. Stops scraping one line **before** `"ramanuja dasan"` or `"adiyEn"`
5. Detects `<li><a href="...">` links → treats each as a **Chapter**
6. Builds a **Table of Contents** on Page 1 of the document
7. Appends main page content as Page 2
8. Iterates each chapter link, scrapes each page, appends under Heading 2

---

## Output Document Structure

```
Page 1  — Table of Contents
Page 2  — Main page content
Page 3+ — Chapter 1 (Heading 2) + content
           Chapter 2 (Heading 2) + content
           ...
```

---

## Project Structure

```
kaink1_tiruvirutham_ebook/
├── src/
│   └── scraper.py          # Main scraper script
├── requirements.txt        # Python dependencies
├── tasks.md                # Development task tracker
└── README.md               # This file
```

---

## Security Notes

- Input URL is validated before use (SSRF protection)
- Output path is validated to prevent directory traversal
- Scraped content is never executed or evaluated
- Maximum chapter link count is capped to prevent infinite loops

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| `playwright install` fails | Run PowerShell as Administrator |
| Chrome does not open | Ensure Chromium was installed via `playwright install chromium` |
| Output .docx not created | Check the output directory exists and is writable |
| Start/stop marker not found | Verify the page contains `varavaramunayE nama:` in its content |

