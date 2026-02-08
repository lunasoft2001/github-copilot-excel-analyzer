# vbaExcel Skill

Skill de GitHub Copilot para exportar y reimportar codigo VBA desde Excel .xlsm en Windows.

## Overview

Este skill permite:

- Exportar modulos VBA a archivos .bas
- Refactorizar el codigo en VS Code
- Reimportar los cambios al XLSM

## Repository Structure

```
github-copilot-excel-analyzer/
├── skill-bundle/              # Bundle limpio listo para instalar
│   ├── SKILL.md               # Definicion del skill
│   ├── scripts/               # Scripts de export/import
│   └── INSTALL.txt            # Instrucciones minimas
├── install-skill.ps1          # Instalador automatico
├── README.md                  # Este archivo
└── LICENSE
```

## Instalacion rapida

```powershell
./install-skill.ps1
```

El instalador copia `skill-bundle` a `%USERPROFILE%\.copilot\skills\vbaExcel` y valida:

- Excel instalado
- Acceso a VBProject (AccessVBOM)

## Uso basico

Exportar:

```powershell
python skill-bundle/scripts/export_vba.py path\file.xlsm VBA_EXPORT
```

Refactoriza los `.bas` en `VBA_EXPORT`.

Reimportar:

```powershell
python skill-bundle/scripts/import_vba.py path\file.xlsm VBA_EXPORT
```

## Requisitos

- Windows
- Excel instalado
- Python 3.10+
- pywin32

## Seguridad

- Cierra Excel antes de exportar o importar.
- Haz backup del XLSM antes de reimportar.

## Troubleshooting

Si el acceso a VBProject esta bloqueado:

```powershell
reg import skill-bundle\scripts\enable_vba_access.reg
```

## Version

v1.0
