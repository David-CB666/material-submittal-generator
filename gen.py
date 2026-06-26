# -*- coding: utf-8 -*-
"""
Material Submittal Generator — ZIP engine
Usage: python gen.py <config.json>
Dependencies: zipfile, re (standard library)
"""
import sys, io, os, json, zipfile, re, shutil

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')


# ──────────────────────────────────────────────────────────────
# 工具函数
# ──────────────────────────────────────────────────────────────

def esc(s):
    """XML特殊字符转义"""
    if s is None:
        return ''
    return str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


def validate_xml(data, path=''):
    """XML格式验证"""
    try:
        import xml.etree.ElementTree as ET
        ET.fromstring(data)
        return True, ''
    except Exception as e:
        return False, str(e)


def append_before_close(xml_str, tag_name, new_lines):
    """在 </tag_name> 之前插入新行（用 rfind 避免嵌套干扰）"""
    close_tag = f'</{tag_name}>'
    idx = xml_str.rfind(close_tag)
    if idx == -1:
        raise ValueError(f'找不到 </{tag_name}>，XML可能损坏')
    return xml_str[:idx] + '\n'.join(new_lines) + '\n' + xml_str[idx:]


# ──────────────────────────────────────────────────────────────
# 核心：共享字符串 → inlineStr 转换（保留样式索引）
# ──────────────────────────────────────────────────────────────

def build_inline_sheet(sheet1_xml, ss_entries, data, field_map):
    """
    1. 从 sharedStrings 提取所有占位符的实际值
    2. 将 t="s" 单元格转为 inlineStr，同时保留 s="N" 样式索引
    """
    # 占位符文本 → 实际值
    ph_to_val = {ph: data.get(fk, '') for ph, fk in field_map.items()}

    # 构建替换后的 sharedStrings 纯文本列表
    new_ss = []
    for content in ss_entries:
        result = content
        for ph, val in ph_to_val.items():
            result = result.replace(ph, esc(val))
        texts = re.findall(r'<t[^>]*>([^<]*)</t>', result)
        new_ss.append(''.join(texts).strip())

    # 转换 t="s" → t="inlineStr"，关键：保留原单元格 s 属性
    def cell_replacer(m):
        cell_xml = m.group(0)
        ref_m = re.search(r'r="([^"]+)"', cell_xml)
        ref = ref_m.group(1) if ref_m else ''
        # 提取并保留样式索引
        s_m = re.search(r'\bs="(\d+)"', cell_xml)
        s_attr = f' s="{s_m.group(1)}"' if s_m else ''
        v_m = re.search(r'<v>(\d+)</v>', cell_xml)
        if v_m:
            idx = int(v_m.group(1))
            if idx < len(new_ss):
                val = new_ss[idx]
                if val:
                    return f'<c r="{ref}"{s_attr} t="inlineStr"><is><t>{val}</t></is></c>'
        return cell_xml

    return re.sub(r'<c [^>]*t="s"[^>]*>.*?</c>', cell_replacer, sheet1_xml, flags=re.DOTALL)


# ──────────────────────────────────────────────────────────────
# 主函数
# ──────────────────────────────────────────────────────────────

def generate(config_path):
    with open(config_path, 'r', encoding='utf-8') as f:
        cfg = json.load(f)

    tpl       = cfg['template']
    out       = cfg['output']
    data_list = cfg['data']
    field_map = cfg['fields']

    print(f'模板: {tpl}')
    print(f'输出: {out}')
    print(f'数量: {len(data_list)} 条')

    # 1. 读取模板（完整ZIP内容）
    with zipfile.ZipFile(tpl, 'r') as z:
        tpl_files = {n: z.read(n) for n in z.namelist()}

    # 2. 解析 sharedStrings → 占位符列表
    ss_entries = []
    if 'xl/sharedStrings.xml' in tpl_files:
        ss_raw = tpl_files['xl/sharedStrings.xml'].decode('utf-8')
        ss_entries = re.findall(r'<si>(.*?)</si>', ss_raw, re.DOTALL)
        print(f'SharedStrings: {len(ss_entries)} 条，检测到占位符:')
        for idx, content in enumerate(ss_entries):
            texts = re.findall(r'<t[^>]*>([^<]*)</t>', content)
            val = ''.join(texts).strip()
            if val.startswith('{') and val.endswith('}'):
                print(f'  索引{idx}: {val} → 字段"{field_map.get(val, "???")}"')

    # 3. 读取 sheet1 原稿
    sheet1_raw = tpl_files['xl/worksheets/sheet1.xml'].decode('utf-8')

    # 4. 复制全部模板文件
    new_files = {k: v for k, v in tpl_files.items()}

    # 5. 生成每个 Sheet
    el_ids = []
    for i, data in enumerate(data_list):
        el_ids.append(data.get(list(field_map.values())[0], f'Row{i + 1}'))
        xml_content = build_inline_sheet(sheet1_raw, ss_entries, data, field_map)
        if i == 0:
            # 第一项替换 sheet1
            new_files['xl/worksheets/sheet1.xml'] = xml_content.encode('utf-8')
        else:
            # 新建后续 Sheet
            sn = i + 1
            new_files[f'xl/worksheets/sheet{sn}.xml'] = xml_content.encode('utf-8')
            # 复制 sheet1 的关系文件（含 printerSettings + drawing）
            if 'xl/worksheets/_rels/sheet1.xml.rels' in tpl_files:
                new_files[f'xl/worksheets/_rels/sheet{sn}.xml.rels'] = \
                    tpl_files['xl/worksheets/_rels/sheet1.xml.rels']

    print(f'生成: {len(el_ids)} 个 Sheet')

    # 6. 更新 sheet1.xml.rels（printerSettings + drawing 关联）
    new_files['xl/worksheets/_rels/sheet1.xml.rels'] = (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
        b'  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Target="../printerSettings/printerSettings1.bin"/>\n'
        b'  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>\n'
        b'</Relationships>'
    )

    # 7. 更新 workbook.xml（Sheet列表）
    wb_xml = tpl_files['xl/workbook.xml'].decode('utf-8')

    def sheet_rid(i):
        # i=0 → rId1（模板原有 sheet1），i≥1 → rId{i+4}（避免覆盖 theme/styles）
        return 'rId1' if i == 0 else f'rId{i + 4}'

    sheets_parts = [
        f'<sheet name="{esc(name)}" sheetId="{i + 1}" state="visible" r:id="{sheet_rid(i)}"/>'
        for i, name in enumerate(el_ids)
    ]
    wb_xml = re.sub(r'<sheets>.*?</sheets>', '<sheets>' + ''.join(sheets_parts) + '</sheets>',
                    wb_xml, flags=re.DOTALL)
    new_files['xl/workbook.xml'] = wb_xml.encode('utf-8')

    # 8. 更新 workbook.xml.rels（新增 Sheet 关系）
    # 模板原有：rId1=sheet1, rId2=theme, rId3=styles, rId4=sharedStrings
    # 新增 Sheet 从 rId5 开始
    wb_rels = tpl_files['xl/_rels/workbook.xml.rels'].decode('utf-8')
    # 移除 sharedStrings 引用（已转为 inlineStr，不再需要）
    wb_rels = re.sub(r'<Relationship[^>]*sharedStrings[^>]*/>', '', wb_rels)
    # 追加新 Sheet（i=1..N → rId{i+4}）
    new_rel_lines = []
    for i in range(1, len(data_list)):
        rid = i + 4
        new_rel_lines.append(
            f'  <Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{rid}.xml"/>'
        )
    wb_rels_fixed = append_before_close(wb_rels, 'Relationships', new_rel_lines)
    ok, err = validate_xml(wb_rels_fixed, 'workbook.xml.rels')
    if not ok:
        raise RuntimeError(f'workbook.xml.rels XML损坏: {err}')
    new_files['xl/_rels/workbook.xml.rels'] = wb_rels_fixed.encode('utf-8')

    # 9. 更新 Content_Types
    ct = tpl_files['[Content_Types].xml'].decode('utf-8')
    ct = re.sub(r'\s*<Override[^>]*sharedStrings[^>]*/>', '', ct)
    new_override_lines = [
        f'  <Override PartName="/xl/worksheets/sheet{i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for i in range(1, len(data_list))
    ]
    ct_fixed = append_before_close(ct, 'Types', new_override_lines)
    ok, err = validate_xml(ct_fixed, 'Content_Types')
    if not ok:
        raise RuntimeError(f'Content_Types XML损坏: {err}')
    new_files['[Content_Types].xml'] = ct_fixed.encode('utf-8')

    # 10. 删除 sharedStrings.xml（已转为 inlineStr）
    new_files.pop('xl/sharedStrings.xml', None)

    # 11. 写入输出文件
    temp = os.path.join(os.environ.get('TEMP', '.'), '_dem_gen_temp.xlsx')
    with zipfile.ZipFile(temp, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in new_files.items():
            z.writestr(name, data)
    shutil.copy(temp, out)
    if os.path.exists(temp):
        os.remove(temp)

    # 12. 全面验证
    print(f'\n=== 验证 ===')
    with zipfile.ZipFile(out, 'r') as z:
        names = sorted(z.namelist())
        print(f'文件总数: {len(names)}')

        xml_ok = xml_err = 0
        for n in names:
            if n.endswith('.xml'):
                ok, err = validate_xml(z.read(n), n)
                if ok:
                    xml_ok += 1
                else:
                    xml_err += 1
                    print(f'  ❌ {n}: {err}')
        print(f'XML: {xml_ok} OK, {xml_err} errors')

        for f in ['xl/worksheets/sheet1.xml', 'xl/drawings/drawing1.xml',
                  'xl/media/image1.jpeg', 'xl/printerSettings/printerSettings1.bin']:
            if f in names:
                print(f'  ✅ {f}: {len(z.read(f)):,} bytes')
            else:
                print(f'  ❌ {f}: MISSING')

        print('\n数据验证（前5条）:')
        for i in range(min(5, len(data_list))):
            row = data_list[i]
            c = z.read(f'xl/worksheets/sheet{i + 1}.xml').decode('utf-8')
            key = list(field_map.values())[0]
            val = str(row.get(key, ''))[:12]
            checks = {fk: (str(row.get(fk, ''))[:10] in c) for fk in list(field_map.values())[:4]}
            print(f'  [{i + 1}] {val}: {checks}')

        print('\nworkbook.xml.rels 验证:')
        rels = z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
        rids = re.findall(r'Id="(rId\d+)"', rels)
        dup = [r for r in set(rids) if rids.count(r) > 1]
        print(f'  rIds: {rids}')
        if dup:
            print(f'  ❌ 重复ID: {dup}')
        else:
            print('  ✅ 无重复ID')

    size = os.path.getsize(out)
    print(f'\n大小: {size:,} bytes ({size / 1024:.0f} KB)')
    print(f'输出: {out}')
    print('\n✅ 完成！请在 Excel 中打开验证。')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('用法: python gen.py <config.json>')
        sys.exit(1)
    try:
        generate(sys.argv[1])
    except Exception as e:
        print(f'\n错误: {e}', file=sys.stderr)
        import traceback
        traceback.print_exc()
        sys.exit(1)
