# JEBON Cover Generator

macOS GUI tool for batch-generating Korean PDF cover pages from Excel data.

## Features
- Read `RAW` sheet (`권/날짜/지급번호`) from `.xlsx/.xlsm`
- Bulk input modes:
  - Excel import
  - Clipboard paste (3 columns)
  - Volume range generator (`1-100`, `1,3,5-8`)
- Real-time progress dashboard and per-volume status
- A4 portrait cover PDF output with Korean font support

## Run
```bash
python3 cover_generator.py
```

## Install dependencies
```bash
pip3 install pandas openpyxl fpdf2
```

## Build macOS app
```bash
pip3 install pyinstaller
python3 -m PyInstaller --noconfirm --clean --windowed --name JEBONCoverGenerator cover_generator.py
```
