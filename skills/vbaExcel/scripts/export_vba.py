#!/usr/bin/env python
# Exporta codigo VBA a archivos .bas desde un XLSM.
# Uso: python export_vba.py ruta\archivo.xlsm salida_dir

import os
import sys
import subprocess
import time
from pathlib import Path

try:
    import win32com.client  # pywin32
except Exception:
    win32com = None


def run_vbscript(xlsm_path, output_dir):
    vbs = r'''
Set fso = CreateObject("Scripting.FileSystemObject")
Set excel = CreateObject("Excel.Application")

xlsm = WScript.Arguments.Item(0)
output_folder = WScript.Arguments.Item(1)

excel.Visible = False
Set wb = excel.Workbooks.Open(xlsm)
Set proj = wb.VBProject

If Not fso.FolderExists(output_folder) Then
    fso.CreateFolder(output_folder)
End If

count = 0
For Each comp In proj.VBComponents
    name = comp.Name
    Set cm = comp.CodeModule
    lines = cm.CountOfLines
    If lines > 0 Then
        code = cm.Lines(1, lines)
        out_file = output_folder & "\\" & name & ".bas"
        Set f = fso.CreateTextFile(out_file, True)
        f.Write code
        f.Close
        count = count + 1
    End If
Next

wb.Close False
excel.Quit

WScript.Echo "modules=" & count
'''

    vbs_path = os.path.join(output_dir, "_export_vba.vbs")
    with open(vbs_path, "w", encoding="utf-8") as f:
        f.write(vbs)

    result = subprocess.run(
        ["cscript.exe", vbs_path, xlsm_path, output_dir],
        capture_output=True,
        text=True,
        timeout=60,
    )

    return result.returncode == 0, result.stdout, result.stderr


def run_com(xlsm_path, output_dir):
    if win32com is None:
        return False, "pywin32 not available"

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(xlsm_path)

    try:
        proj = wb.VBProject
        count = 0
        for i in range(1, proj.VBComponents.Count + 1):
            comp = proj.VBComponents(i)
            name = comp.Name
            cm = comp.CodeModule
            lines = cm.CountOfLines
            if lines > 0:
                code = cm.Lines(1, lines)
                out_file = os.path.join(output_dir, f"{name}.bas")
                with open(out_file, "w", encoding="utf-8") as f:
                    f.write(code)
                count += 1
        return True, f"modules={count}"
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python export_vba.py <xlsm_path> [output_dir]")
        sys.exit(1)

    xlsm_path = os.path.abspath(sys.argv[1])
    output_dir = os.path.abspath(sys.argv[2]) if len(sys.argv) > 2 else os.path.abspath("VBA_EXPORT")

    os.makedirs(output_dir, exist_ok=True)

    ok, out, err = run_vbscript(xlsm_path, output_dir)
    if ok:
        print(out.strip())
        sys.exit(0)

    ok, out = run_com(xlsm_path, output_dir)
    if ok:
        print(out)
        sys.exit(0)

    print("Export failed.")
    if err:
        print(err.strip())
    sys.exit(1)
