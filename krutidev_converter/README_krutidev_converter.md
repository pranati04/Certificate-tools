# KrutiDev Converter

Converts Hindi text from **Unicode (Devanagari)** to **KrutiDev 10** encoding. Reads an Excel file, detects all Hindi columns automatically, and writes a new Excel file with `_KD` suffix columns added beside each source column, formatted with Kruti Dev 010 font.

Fully self-contained — no internet connection, no external files, no git clone required.

---

## Files in this folder

| File | Description |
|---|---|
| `krutidev_converter.py` | Main program — run this or convert to `.exe` |

---

## Requirements

```
openpyxl
pandas
```

Install with:
```bash
pip install openpyxl pandas
```

---

## How to run

```bash
python krutidev_converter.py
```

Or build to a standalone `.exe` (no Python needed on target machine):
```bash
pip install pyinstaller
pyinstaller --onefile --windowed krutidev_converter.py
```
The `.exe` will appear in the `dist\` folder.

---

## How to use

1. Click **📂 Browse…** next to "Input Excel file" and select your Excel file
2. The output path is filled in automatically (same folder, `_KrutiDev` added to filename)
3. Change the output path with the second Browse button or the **💾 Save As…** button if needed
4. Click **🔄 Convert**
5. The progress bar fills as rows are processed
6. When done, a prompt asks if you want to open the output file

---

## What it does to your Excel file

- Reads every column that contains Hindi (Devanagari Unicode) text
- For each Hindi column, adds a new `_KD` column immediately to its right
- The `_KD` columns are formatted with **Kruti Dev 010** font and a green header
- Original columns are unchanged
- All other (non-Hindi) columns are copied across as-is

### Example

| student_name | student_name_KD | city | city_KD |
|---|---|---|---|
| राम | jke | दिल्ली | fnYyh |
| सीता | lhrk | मुम्बई | eqEcbZ |

---

## Supported conversions

- All standard Devanagari consonants and vowels
- Vowel matras (including the ि matra position fix — KrutiDev convention)
- Conjunct consonants: क्ष, त्र, ज्ञ, श्र, प्र, स्त, and more
- Nukta variants: ज़, ड़, ढ़, फ़
- Devanagari digits (० १ २ … → 0 1 2 …)
- Punctuation: । ॥ ॐ

---

## Buttons

| Button | What it does |
|---|---|
| 🔄 Convert | Convert the input file and save to the output path shown |
| 💾 Save As… | Pick a new save location then convert immediately |
| 📂 Open output folder | Opens the folder where the output file was saved |

---

## Troubleshooting

**"No Devanagari columns found"**
- Make sure the Excel file actually contains Hindi Unicode text (not already KrutiDev-encoded text)
- KrutiDev-encoded text looks like Latin characters (`jke`, `lhrk`) — it is already converted and does not need processing

**Output file looks correct in Excel but Hindi columns show boxes or wrong characters**
- The Kruti Dev 010 font must be installed on your computer
- Install the Kruti Dev font family from your system fonts

**File is locked / cannot save**
- Make sure the output Excel file is not already open in Excel
- Close it and try again
