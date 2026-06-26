# -*- coding: utf-8 -*-
"""
Template Inspector — diagnose .xlsx structure
Usage: python inspect.py <template.xlsx>
Output: placeholder list, cell style map, template overview
"""
import sys, io, os, json, zipfile, re

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')


def validate_xml(data, path=''):
    try:
        import xml.etree.ElementTree as ET
        ET.fromstring(data)
        return True, ''
    except Exception as e:
        return False, str(e)


def inspect(path):
    print(f'文件: {path}')
    print('=' * 60)

    with zipfile.ZipFile(path, 'r') as z:
        names = sorted(z.namelist())
        print(f'ZIP内文件总数: {len(names)}')
        for n in names:
            print(f'  {n}: {len(z.read(n)):,} bytes')

        # sharedStrings
        if 'xl/sharedStrings.xml' in names:
            ss = z.read('xl/sharedStrings.xml').decode('utf-8')
            entries = re.findall(r'<si>(.*?)</si>', ss, re.DOTALL)
            print(f'\n=== SharedStrings ({len(entries)}条) ===')
            for idx, content in enumerate(entries):
                texts = re.findall(r'<t[^>]*>([^<]*)</t>', content)
                val = ''.join(texts).strip()
                if val:
                    tag = '(占位符)' if (val.startswith('{') and val.endswith('}')) else ''
                    print(f'  [{idx}] {val} {tag}')
        else:
            print('\n[无 sharedStrings]')

        # worksheet
        if 'xl/worksheets/sheet1.xml' in names:
            sheet = z.read('xl/worksheets/sheet1.xml').decode('utf-8')
            print(f'\n=== Sheet1 ({len(sheet):,} bytes) ===')

            # 占位符 → 单元格位置
            if 'xl/sharedStrings.xml' in names:
                ss = z.read('xl/sharedStrings.xml').decode('utf-8')
                entries = re.findall(r'<si>(.*?)</si>', ss, re.DOTALL)
                ph_to_idx = {}
                for idx, content in enumerate(entries):
                    texts = re.findall(r'<t[^>]*>([^<]*)</t>', content)
                    val = ''.join(texts).strip()
                    if val.startswith('{') and val.endswith('}'):
                        ph_to_idx[val] = idx

                print('\n占位符 → 单元格位置:')
                for ph, idx in ph_to_idx.items():
                    # 通用匹配（支持命名空间属性顺序）
                    m = re.search(r'<c\b([^>]*)>(?:[^<]|<(?!c[ />]))*<v>' + str(idx) + r'</v>', sheet)
                    if m:
                        attrs = m.group(1)
                        ref_m = re.search(r'r="([^"]+)"', attrs)
                        s_m = re.search(r's="(\d+)"', attrs)
                        ref = ref_m.group(1) if ref_m else '?'
                        s_val = s_m.group(1) if s_m else '无'
                        print(f'  {ph} → {ref} (s={s_val})')

            # 合并单元格
            merges = re.findall(r'<mergeCell ref="([^"]+)"', sheet)
            print(f'\n合并单元格 ({len(merges)}个):')
            for m in merges:
                print(f'  {m}')

            # 列宽
            cols = re.findall(r'<col\s+min="(\d+)"\s+max="(\d+)"\s+width="([^"]+)"', sheet)
            if cols:
                print(f'\n列宽定义 ({len(cols)}段):')
                for mn, mx, w in cols:
                    print(f'  列{min(mn)}~{mx}: width={w}')

            # 图片
            if 'xl/worksheets/_rels/sheet1.xml.rels' in names:
                rels = z.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
                drawings = re.findall(r'Target="([^"]*drawing[^"]*)"', rels)
                print(f'\n图片关联: {drawings}')

        # workbook
        if 'xl/workbook.xml' in names:
            wb = z.read('xl/workbook.xml').decode('utf-8')
            sheets = re.findall(r'<sheet\s+name="([^"]+)"[^>]*r:id="([^"]+)"', wb)
            print(f'\n=== Workbook ({len(sheets)}个Sheet) ===')
            for name, rid in sheets:
                print(f'  {rid}: {name}')

        # printerSettings
        if 'xl/printerSettings/printerSettings1.bin' in names:
            print('\n打印机设置: ✅ 存在')
        else:
            print('\n打印机设置: ❌ 缺失')

        # media
        media = [n for n in names if n.startswith('xl/media/')]
        if media:
            print(f'媒体文件: {media}')
        else:
            print('媒体文件: ❌ 缺失')

    # 文件大小
    size = os.path.getsize(path)
    print(f'\n文件大小: {size:,} bytes ({size / 1024:.0f} KB)')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python inspect.py <模板.xlsx>')
        sys.exit(1)
    inspect(sys.argv[1])
