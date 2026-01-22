# EDI Notepad++ Search & Screenshot Tool

A Python automation tool that opens EDI files in Notepad++, searches for specific segments/fields, highlights them, and captures screenshots with a red rectangle around the found line.

## Features

- 🔍 **Search EDI segments** - Search for any EDI segment (ISA, CLM, NM1, DTP, etc.)
- 📸 **Automated screenshots** - Captures the Notepad++ window with highlighted line
- 🔲 **Visual highlighting** - Draws a red rectangle around the searched line
- 📁 **Multiple input formats** - Command line, text file, or Excel file
- 📊 **Sub-element support** - Parse CLM05-1, CLM05-2, SV101-1, etc.
- 📝 **Not found logging** - Logs fields that couldn't be found

## Installation

### Requirements

- Python 3.8+
- Notepad++ installed
- Windows OS

### Install Dependencies

```bash
pip install pyautogui pyperclip pywin32 pillow pandas openpyxl
```

## Usage

### Basic Usage

```bash
# Single field
python main.py --file EDI.txt --word "BHT03"

# Multiple fields (comma-separated)
python main.py --file EDI.txt --word "BHT03,CLM05,DTP01"
```

### List Input

```bash
python main.py --file EDI.txt --list BHT03 CLM05 NM101 DTP01
```

### Text File Input

```bash
python main.py --file EDI.txt --txt fields.txt
```

**fields.txt format:**
```
# Comments start with #
BHT03
CLM05
CLM05-1
CLM05-2
NM101
DTP01
SV101-1
```

### Excel Input

```bash
python main.py --file EDI.txt --excel input.xlsx
```

**Excel format:**

| File name | Field |
|-----------|-------|
| ISA Segment | ISA01 |
| BHT Segment | BHT03 |
| CLM Segment | CLM05 |
| CLM Sub-element 1 | CLM05-1 |
| NM1 Segment | NM101 |

- **File name** column: Used as screenshot filename
- **Field** column: Field code to search for

## Field Code Format

### Supported Formats

| Format | Description | Search Term |
|--------|-------------|-------------|
| `BHT03` | BHT segment, element 03 | `BHT*` |
| `CLM05` | CLM segment, element 05 | `CLM*` |
| `CLM05-1` | CLM element 05, sub-element 1 | `CLM*` |
| `CLM05-2` | CLM element 05, sub-element 2 | `CLM*` |
| `NM101` | NM1 segment, element 01 | `NM1*` |
| `NM109` | NM1 segment, element 09 | `NM1*` |
| `SV101` | SV1 segment, element 01 | `SV1*` |
| `SV101-1` | SV1 element 01, sub-element 1 | `SV1*` |
| `ISA01` | ISA segment, element 01 | `ISA*` |

### Sub-elements Explained

In EDI 837 files, composite elements contain sub-elements separated by `:`. For example:

```
CLM*36463774*100***11:B:1*Y*A*Y*Y**...
                   ↑  ↑  ↑
               CLM05-1 CLM05-2 CLM05-3
```

- `CLM05-1` = Place of Service Code (11 = Office)
- `CLM05-2` = Facility Code Value (B = Hospital)
- `CLM05-3` = Claim Frequency Code (1 = Original)

## Output

### Screenshots

Screenshots are saved to the `screenshots/` folder with:
- Full screen capture
- Red rectangle highlighting the found line
- Filename based on "File name" column (Excel) or field code

### Console Output

```
============================================================
  COMPLETE
============================================================
Total screenshots: 19
Found: 19
Not found: 0

Screenshots:
  - C:\...\screenshots\ISA Segment.png
  - C:\...\screenshots\CLM Segment.png
  ...
```

### Not Found Log

If any fields are not found, they are logged to `screenshots/not_found.log`:

```
Not Found Fields - 2026-01-23 00:30:00
==================================================

XYZ01 (searched: XYZ*)
ABC02 (searched: ABC*)
```

## Command Line Options

| Option | Short | Description |
|--------|-------|-------------|
| `--file FILE` | `-f` | **Required.** EDI file to open in Notepad++ |
| `--word WORD` | `-w` | Field codes (comma-separated) |
| `--list LIST...` | `-l` | Field codes (space-separated) |
| `--txt TXT` | `-t` | Text file with field codes (one per line) |
| `--excel EXCEL` | `-e` | Excel file with "File name" and "Field" columns |
| `--help` | `-h` | Show help message |

## Notes

⚠️ **Don't move the mouse** during the automation process  
⚠️ The tool has a 2-second countdown before starting  
⚠️ Notepad++ must be installed on your system  

## License

MIT License
