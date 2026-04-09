# RDV Certificate Tools

A set of Python tools built for **Rani Durgavati Vishwavidyalaya, Jabalpur** to automate the generation of degree certificates and the conversion of Hindi text between Unicode (Devanagari) and KrutiDev 10 encoding.

---

## Tools in this repo

| Folder | Tool | What it does |
|---|---|---|
| `cert_generator/` | `cert.py` | Fills student data into a `.docx` certificate template |
| `krutidev_converter/` | `krutidev_converter.py` | Converts Hindi Unicode text in Excel to KrutiDev 10 encoding |

---

## Files NOT included in this repo

These files are kept off GitHub for privacy and licensing reasons.
You must supply them yourself and place them in the correct folder before running.

| File | Where to put it | Why not included |
|---|---|---|
| `certificate_template.docx` | `cert_generator/` | University document — not public |
| `Book1_krutidev.xlsx` | `cert_generator/` | Contains student personal data |
| Any other `.docx` templates | `cert_generator/` | University documents |

---

## Requirements

- Python 3.10 or newer — [python.org](https://www.python.org/downloads/)
- See each tool's folder for its specific dependencies

---

## Quick start

```bash
# Clone the repo
git clone https://github.com/YOUR_USERNAME/rdv-certificate-tools.git
cd rdv-certificate-tools

# Install dependencies for the certificate generator
pip install python-docx openpyxl

# Install dependencies for the KrutiDev converter
pip install openpyxl pandas

# Run the certificate generator
python cert_generator/cert.py

# Run the KrutiDev converter
python krutidev_converter/krutidev_converter.py
```

---

## Platform

Developed and tested on **Windows 10/11**.
Both tools use `tkinter` for their GUI, which is included with standard Python on Windows.

---

## Developer

Built for internal use at Rani Durgavati Vishwavidyalaya, Jabalpur.
