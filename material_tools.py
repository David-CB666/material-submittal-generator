# -*- coding: utf-8 -*-
"""
Material Submittal Toolkit — 材料審批工具箱
Usage:
  material-tools gen <config.json>      生成材料報批表
  material-tools inspect <模板.xlsx>    診斷模板結構
  material-tools merge-bq <output.xlsx> <bq-folder>  合併BQ頁
  material-tools export-pdf <input.xlsx> [output.pdf] 匯出PDF
"""
import sys
import os

# Fix encoding for Windows console
if sys.platform == 'win32':
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')


def show_help():
    print("""
╔══════════════════════════════════════════════════╗
║      Material Submittal Toolkit v1.0             ║
║      材料審批工具箱 — MEP Construction            ║
╚══════════════════════════════════════════════════╝

用法:
  material-tools gen <config.json>              生成材料報批表
  material-tools inspect <模板.xlsx>            診斷模板結構
  material-tools merge-bq <output.xlsx> <bq-folder>  合併BQ標書頁
  material-tools export-pdf <input.xlsx> [output.pdf] 匯出PDF

快速入門:
  1. 先診斷模板:   material-tools inspect 模板.xlsx
  2. 準備config:   參考 docs/config-example.json
  3. 一鍵生成:     material-tools gen config.json
  4. 合併BQ頁:     material-tools merge-bq output.xlsx ./BQ_Pages/
  5. 匯出PDF:      material-tools export-pdf output.xlsx

更多資訊: https://github.com/David-CB666/material-submittal-generator
""")


def main():
    if len(sys.argv) < 2:
        show_help()
        sys.exit(0)

    cmd = sys.argv[1].lower()

    if cmd in ('-h', '--help', 'help'):
        show_help()
        sys.exit(0)

    if cmd == 'gen':
        if len(sys.argv) < 3:
            print('❌ 用法: material-tools gen <config.json>')
            print('   請提供 config JSON 文件路徑')
            sys.exit(1)
        from gen import generate
        generate(sys.argv[2])

    elif cmd == 'inspect':
        if len(sys.argv) < 3:
            print('❌ 用法: material-tools inspect <模板.xlsx>')
            print('   請提供 Excel 模板文件路徑')
            sys.exit(1)
        from template_inspect import inspect
        inspect(sys.argv[2])

    elif cmd == 'merge-bq':
        if len(sys.argv) < 4:
            print('❌ 用法: material-tools merge-bq <output.xlsx> <bq-folder>')
            print('   output.xlsx: 生成的材料報批表')
            print('   bq-folder:   放BQ標書PDF的資料夾')
            sys.exit(1)
        from merge_bq import merge_bq_pages
        merge_bq_pages(sys.argv[2], sys.argv[3])

    elif cmd == 'export-pdf':
        if len(sys.argv) < 3:
            print('❌ 用法: material-tools export-pdf <input.xlsx> [output.pdf]')
            print('   需要安裝 Microsoft Excel')
            sys.exit(1)
        out_pdf = sys.argv[3] if len(sys.argv) > 3 else sys.argv[2].replace('.xlsx', '.pdf')
        from export_pdf import export_sheets_to_pdf
        export_sheets_to_pdf(sys.argv[2], out_pdf)

    else:
        print(f'❌ 未知命令: {cmd}')
        show_help()
        sys.exit(1)


if __name__ == '__main__':
    try:
        main()
    except FileNotFoundError as e:
        print(f'❌ 找不到文件: {e}')
        sys.exit(1)
    except Exception as e:
        print(f'❌ 錯誤: {e}')
        import traceback
        traceback.print_exc()
        sys.exit(1)
