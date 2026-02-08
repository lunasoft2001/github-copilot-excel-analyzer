#!/usr/bin/env python
# Importa codigo VBA desde archivos .bas a un XLSM.
# Uso: python import_vba.py ruta\archivo.xlsm dir_bas

import os
import sys
from pathlib import Path

import win32com.client


def import_vba(xlsm_path, bas_dir):
    if not os.path.exists(xlsm_path):
        print("XLSM not found")
        return False
    if not os.path.exists(bas_dir):
        print("BAS dir not found")
        return False

    bas_files = list(Path(bas_dir).glob("*.bas"))
    if not bas_files:
        print("No .bas files found")
        return False

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(os.path.abspath(xlsm_path))

    try:
        proj = wb.VBProject
        for bas_file in bas_files:
            name = bas_file.stem
            found = False
            for i in range(1, proj.VBComponents.Count + 1):
                comp = proj.VBComponents(i)
                if comp.Name == name:
                    found = True
                    with open(bas_file, "r", encoding="utf-8") as f:
                        code = f.read()
                    cm = comp.CodeModule
                    if cm.CountOfLines > 0:
                        cm.DeleteLines(1, cm.CountOfLines)
                    if code:
                        cm.InsertLines(1, code)
                    break
            if not found:
                print(f"Module not found in VBA project: {name}")

        wb.Save()
        return True
    finally:
        wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python import_vba.py <xlsm_path> <bas_dir>")
        sys.exit(1)

    ok = import_vba(sys.argv[1], sys.argv[2])
    sys.exit(0 if ok else 1)
