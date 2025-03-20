
# PDF Toolkit

A desktop application built with Python and Tkinter for managing PDF files. The toolkit includes various PDF manipulation features like merging, splitting, converting images to PDFs, removing pages, and adding signatures.

---

## Features

### ✅ Image to PDF
- Convert multiple image files into a single PDF.
- Supports formats: JPG, JPEG, PNG, BMP, TIFF.

### ✅ PDF Merger
- Merge multiple PDF files into one.
- Reorder PDF files before merging.

### ✅ PDF Splitter
- Split PDF files by range or into individual pages.

### ✅ Page Remover
- Remove specific pages from a PDF.

### ✅ Fill & Sign
- Add text and signature to a PDF.
- Preview the PDF and adjust signature position.

---
## Download
You can download from here: https://sourceforge.net/projects/pdf-tools/
## Installation

### 1. Clone the repository:
```bash
git clone https://github.com/PolyMokashi/Pdf-Toolkit.git
cd Pdf-Toolkit
```

### 2. Create a virtual environment and activate it:
```bash
python -m venv venv
source venv/bin/activate    # On Windows use: .\venv\Scripts\activate
```

### 3. Install dependencies:
```bash
pip install -r requirements.txt
```

---

## Usage
1. Run the application:
```bash
python Pdf_Tools.py
```

2. Use the tab-based interface to access different features:
   - **Image to PDF** – Convert images to PDF.
   - **PDF Merger** – Merge multiple PDFs.
   - **PDF Splitter** – Split a PDF file.
   - **Page Remover** – Remove specific pages from a PDF.
   - **Fill & Sign** – Add text and signatures to PDFs.

---

## Dependencies
- `tkinter`
- `Pillow`
- `PyPDF2`
- `fitz` (PyMuPDF)

Install them using:
```bash
pip install tkinter pillow pypdf2 pymupdf
```

## Contributing
1. Fork the repository.
2. Create a new branch (`git checkout -b feature-branch`).
3. Commit changes (`git commit -m 'Add feature'`).
4. Push to the branch (`git push origin feature-branch`).
5. Open a pull request.

---

## License
This project is licensed under the **MIT License**.

---

## Author
[Pradeep Dilip Mokashi](https://github.com/PolyMokashi)

---
