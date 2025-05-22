# GeoChemical Analysis Extractor

This desktop application helps geotechnical and environmental professionals extract, verify, and convert complex geochemical analysis tables from **PDF and Excel** files into a structured `.xlsx` format, compatible with GIS or reporting tools.

It supports both **automated extraction from structured tables** (via Camelot) and **OCR-based image table extraction** (via Tesseract + OpenCV). It also includes a **mapping tool** to merge extracted data into a standard attribute table.

---

## 🔧 Features

### 📄 PDF Table Extraction
- Detect and extract tables from PDF pages based on predefined keywords (e.g. *"Essai laboratoire"*, *"in situ"*).
- Use **Camelot** (lattice/stream) for structured tables.
- Detect and log pages where no tables are found.

### 🖼 OCR on PDF Images
- Convert image-based tables to text using **Tesseract OCR** and **OpenCV**.
- Detect grid structures, segment cells, and extract text.
- Option to extract page-by-page or batch-process multiple pages.

### 📊 Validation Interface (Sondage)
- Interactive UI to **manually verify and clean extracted values** (Pf*, Pl*, EM/Module).
- Detect inconsistencies (spacing anomalies).
- Visualize and validate borehole-specific data.
- Export validated data to clean Excel format.

### 🔁 Attribute Mapping Tool
- Automatically map extracted values from multiple sheets to a predefined **attribute table template**.
- Match "in situ" and "laboratoire" tests to standard columns.
- Export a fully cleaned and standardized `.xlsx` file.

---

## 🧪 How to Use

### Step 1 – Launch
Double-click on `main.exe` #to be done

### Step 2 – Extract Tables from PDF
1. Select one or more PDF files.
2. Choose an output folder.
3. Click **"Lancer l'extraction"**.
4. If some pages are detected as image-only, you will be prompted to run OCR on them.

### Step 3 – OCR Table from Images (optional)
1. Use **"Tester détection tableau image (OpenCV)"** to manually extract from image-based pages.
2. Optionally batch-process segmented images using OCR.

### Step 4 – Mapping (Optional)
1. Click **"Mapping express vers table attributaire"**.
2. Select which sheets correspond to *"in situ"* and *"laboratoire"* tests.
3. The script will generate a cleaned attribute table based on your template.

---

## 📁 Files & Configuration

| File                          | Description                                              |
|-------------------------------|----------------------------------------------------------|
| `main.py`                     | Main launcher with GUI                                   |
| `image_table_extract.py`      | Extracts tables from PDF images using OCR                |
| `pdf_table_extract.py`        | Extracts structured tables from PDFs (Camelot)           |
| `pdf_sondage_extract.py`      | Validates borehole data (Pf*, Pl*, EM) with UI           |
| `mapping_tool.py`             | Maps extracted data to attribute table format            |
| `template_attributaire.xlsx`  | Template used as a reference for final export            |

---

## 📦 Requirements (if running from source)

- Python 3.10+
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) (installed & path set)
- Camelot
- fitz (PyMuPDF)
- pandas
- openpyxl
- opencv-python
- pdfplumber

Install dependencies:
```bash
pip install -r requirements.txt