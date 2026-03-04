# AGENTS.md (Project - jebon-cover-generator)

## Trigger
- If user says `제본`, they mean this project.

## Read First
1. `JEBON_COVER_CLI_HANDOFF.md`
2. `README.md`
3. `cover_generator.py`

## Scope
- GUI batch cover generator for Korean PDF covers.
- Inputs: Excel (`RAW` A:C), Clipboard (`권/날짜/지급번호`), Volume range.
- Outputs: A4 cover PDFs with volume/date/payment number.

## Must Keep
- Real-time progress UI + per-volume status table.
- Cancel button behavior.
- Cross-platform branches:
  - macOS/Windows/Linux folder open
  - Korean font fallback candidates.

## Common Commands
```bash
python3 cover_generator.py
python3 -m py_compile cover_generator.py
python3 -m PyInstaller --noconfirm --clean --windowed --name JEBONCoverGenerator cover_generator.py
```
