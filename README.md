# vbaExcel Skill

A GitHub Copilot skill to export and reimport VBA code from Excel .xlsm files on Windows.

## What is included

- `skills/vbaExcel/SKILL.md`
- `skills/vbaExcel/scripts/export_vba.py`
- `skills/vbaExcel/scripts/import_vba.py`
- `skills/vbaExcel/scripts/enable_vba_access.reg`

## Requirements

- Windows
- Excel installed
- Python 3.10+
- `pywin32` installed

## Quick start

1. Export VBA modules:

```powershell
python skills/vbaExcel/scripts/export_vba.py path\file.xlsm VBA_EXPORT
```

2. Refactor the `.bas` files in `VBA_EXPORT`.

3. Import VBA modules:

```powershell
python skills/vbaExcel/scripts/import_vba.py path\file.xlsm VBA_EXPORT
```

## Notes

- Close Excel before exporting or importing.
- Create a backup of your XLSM before importing.
- If access to VBProject is blocked, run:

```powershell
reg import skills\vbaExcel\scripts\enable_vba_access.reg
```

## Version

v1.0
