# vbaExcel Skill

GitHub Copilot skill to export and reimport VBA code from Excel .xlsm files on Windows.

## Structure

- `skill-bundle/` (clean install bundle)
- `install-skill.ps1` (installer)

## Install

```powershell
./install-skill.ps1
```

This copies `skill-bundle` to `%USERPROFILE%\.copilot\skills\vbaExcel`.

## Use

```powershell
python skill-bundle/scripts/export_vba.py path\file.xlsm VBA_EXPORT
python skill-bundle/scripts/import_vba.py path\file.xlsm VBA_EXPORT
```

## Requirements

- Windows
- Excel installed
- Python 3.10+
- pywin32

## Version

v1.0
