# Certificate Generator

Generates degree certificates for **Rani Durgavati Vishwavidyalaya** by filling student data into a Word (`.docx`) template. Supports both the Hindi (KrutiDev) and English sides of the certificate simultaneously.

---

## Files in this folder

| File | Description |
|---|---|
| `cert.py` | Main program — run this |

### Files you must supply yourself

These are not in the repo. Place them in this same folder before running.

| File | Description |
|---|---|
| `certificate_template.docx` | Your university certificate template (`.doc` also works) |
| `Book1_krutidev.xlsx` | Student data Excel file |

---

## Requirements

```
python-docx
openpyxl
```

Install with:
```bash
pip install python-docx openpyxl
```

For `.doc` file support (older Word format) you also need **LibreOffice**:
- Download from [libreoffice.org](https://www.libreoffice.org/download/download/)
- `.docx` files work without it

---

## How to run

```bash
python cert.py
```

Or double-click `cert.py` if Python is associated with `.py` files on your system.

---

## How it works

1. **Load Template** — click the button and select your `.docx` or `.doc` certificate template
2. **Load Excel** — browse to your student data file (`Book1_krutidev.xlsx`)
3. **Look Up Student** — type a Roll Number, Application Number, or Enrolment Number and press Enter or click Look Up
4. **Edit if needed** — all fields are editable before generating
5. **Set the date** — pick from the preset dates dropdown, or choose "Type custom date…" to enter your own
6. **Generate** — click 🎓 Generate DOCX to produce the certificate

The generated certificate opens automatically in Word.

---

## Excel columns used

The program reads the following columns from your Excel file:

| Column | Used for |
|---|---|
| `roll_number` | Primary student lookup |
| `applicationno` | Secondary lookup |
| `enroll_number` | Tertiary lookup |
| `student_name` | Name in English |
| `student_name_hindi_KrutiDev` | Name in KrutiDev Hindi encoding |
| `college_dec` | College name |
| `year_term_code` | Year of examination |
| `division` | Result division (FIRST / SECOND / THIRD / PASS / DISTINCTION) |
| `gender` | M or F — controls all Hindi gender words |

---

## Gender-sensitive Hindi fields

The program automatically switches the following KrutiDev words based on the student's gender column:

| Field | Male | Female | Meaning |
|---|---|---|---|
| Pronoun | `bUgsa` | `mUgsa` | उन्हें |
| Genitive | `ds` | `dh` | के / की |
| Student | `Nk=` | `Nk=k` | छात्र / छात्रा |
| Passed suffix | `mÙkh.kZ dhA` | `mÙkh.kZ dh` | उत्तीर्ण किया / की |
| Certify verb | `tkrk` | `tkrh` | जाता / जाती |

All of these are also visible and manually editable in the Certificate Particulars panel before generating.

---

## Certificate date

The date section has two modes:

- **Preset dropdown** — select from a list of predefined convocation dates
- **Custom** — choose "Type custom date…" at the bottom of the dropdown and type any date in the format `5th March, 2024.`

The KrutiDev (Hindi) date is always auto-converted from the English date you select or type.

To update the preset date list, open `cert.py` in a text editor and find the `PRESET_DATES` list near the top of the `App` class.

---

## Search → Replace Mappings

The right-hand panel shows all the find-and-replace rules that are applied to the template. Each rule has:

- **Search text** — the exact placeholder text in your template
- **Value key** — the data field to substitute in (e.g. `NAME_EN`, `ROLL_NO`)

You can add, remove, reorder, and toggle rules. Save your custom rules as a `.certmaps` file to reuse them with different templates.

---

## How fonts are preserved

The program works directly on the Word XML. It only changes the text content of each `<w:t>` node — it never touches `<w:rPr>` (the run properties element that carries font, size, bold, colour). This means every font set by the template designer — including KrutiDev display fonts — is preserved exactly.

---

## Troubleshooting

**"No record found" when looking up a student**
- Make sure the roll number matches exactly (no leading zeros dropped, etc.)
- Try the application number or enrolment number instead

**Hindi text looks garbled after generating**
- The KrutiDev fonts must be installed on the computer opening the `.docx`
- Install the Kruti Dev font family from your system fonts

**`.doc` file fails to convert**
- Install LibreOffice and make sure `soffice` is on your system PATH
- Or open the `.doc` in Word and Save As `.docx` manually, then load the `.docx`

**A field is not being replaced**
- Open the Mappings panel and check the Search text matches exactly what is in your template (including spaces)
- Copy the text directly from the template and paste it into the Search text box
