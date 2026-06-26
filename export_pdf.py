# -*- coding: utf-8 -*-
"""
export_sheets_to_pdf.py - Excel多Sheet导出PDF（每个Sheet单独一个PDF）
用法：python export_sheets_to_pdf.py <Excel路径> <输出目录>
依赖：Windows + Excel（通过 COM 接口）
"""
import sys, io, os, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

import pythoncom
import win32com.client

if len(sys.argv) < 3:
    print('用法: python export_sheets_to_pdf.py <Excel路径> <输出目录>')
    sys.exit(1)

SRC_XLSX = os.path.abspath(sys.argv[1])
OUT_DIR  = sys.argv[2]
os.makedirs(OUT_DIR, exist_ok=True)

def safe_name(s):
    """去除文件名中的非法字符"""
    return re.sub(r'[\\/:*?"<>|]', '-', s)

pythoncom.CoInitialize()
try:
    excel = win32com.client.Dispatch('Excel.Application')
    excel.Visible       = False
    excel.DisplayAlerts = False

    wb = excel.Workbooks.Open(SRC_XLSX)
    count = wb.Worksheets.Count
    print(f'打开: {os.path.basename(SRC_XLSX)}  ({count} Sheets)')

    for i in range(1, count + 1):
        ws        = wb.Worksheets(i)
        sheet_name = ws.Name
        pdf_path  = os.path.join(OUT_DIR, f'{safe_name(sheet_name)}.pdf')

        ws.PageSetup.PrintArea       = ws.UsedRange.Address
        ws.PageSetup.CenterVertically = True
        ws.ExportAsFixedFormat(0, pdf_path)
        print(f'  [{i:2d}/{count}] {sheet_name}')

    wb.Close(False)
    excel.Quit()
    print(f'\n完成！{count} 个PDF → {OUT_DIR}')

except Exception as e:
    print(f'错误: {e}')
    try: wb.Close(False); excel.Quit()
    except: pass
finally:
    pythoncom.CoUninitialize()
