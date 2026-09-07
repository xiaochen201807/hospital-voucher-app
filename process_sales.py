#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
西药销售明细表生成脚本
功能：
1. 读取Excel表格的第一个Sheet（西药销售汇总数据）。
2. 根据【药品名称】、【规格】、【剂型】、【制药厂】4列作为唯一键进行分组聚合。
3. 对【数量】、【进价金额】、【零价金额】进行求和汇总，保留单位与最新库存。
4. 按照【数量】从大到小降序排列（同数量时按药品名称拼音排序）。
5. 自动生成第二个Sheet（销售明细表），并附带标题、表头、数据明细及合计行，设置标准格式与列宽。
6. 支持 .xls（基于 xlrd / xlwt / xlutils）与 .xlsx（基于 openpyxl）格式。
"""

import os
import sys
import argparse
from collections import OrderedDict

# 默认列定义
GROUP_KEYS = ['药品名称', '规格', '剂型', '制药厂']
TARGET_HEADERS = ['药品名称', '规格', '剂型', '制药厂', '单位', '数量', '进价金额', '零价金额', '库存']

# 列宽设置 (xlwt 字符单位)
COL_WIDTHS_XLS = {
    0: 4600,  # 药品名称
    1: 2400,  # 规格
    2: 1600,  # 剂型
    3: 3000,  # 制药厂
    4: 1000,  # 单位
    5: 2200,  # 数量
    6: 3000,  # 进价金额
    7: 2800,  # 零价金额
    8: 2200,  # 库存
}

# openpyxl 列宽 (字符数)
COL_WIDTHS_XLSX = {
    'A': 26,
    'B': 16,
    'C': 10,
    'D': 28,
    'E': 8,
    'F': 12,
    'G': 15,
    'H': 15,
    'I': 12,
}


def pinyin_sort_key(name):
    """用于拼音排序的辅助函数"""
    try:
        return name.encode('gbk', errors='ignore')
    except Exception:
        return name


def process_sales_data(rows):
    """
    对从第一个Sheet读取的原始行数据进行清洗与聚合：
    rows: 包含表头及所有数据行的二维列表
    返回: (header_map, sorted_records, original_count, totals)
    """
    if not rows:
        raise ValueError("数据表为空！")

    headers = [str(cell).strip() for cell in rows[0]]
    header_map = {}
    for idx, col in enumerate(headers):
        header_map[col] = idx

    # 校验必要字段
    for col in GROUP_KEYS + ['单位', '数量', '进价金额', '零价金额', '库存']:
        if col not in header_map:
            raise KeyError(f"第一个Sheet缺少必要列: '{col}'")

    idx_name = header_map['药品名称']
    idx_spec = header_map['规格']
    idx_form = header_map['剂型']
    idx_mfr = header_map['制药厂']
    idx_unit = header_map['单位']
    idx_qty = header_map['数量']
    idx_in_amt = header_map['进价金额']
    idx_retail_amt = header_map['零价金额']
    idx_stock = header_map['库存']

    groups = OrderedDict()
    original_count = 0
    source_totals = None

    for r in range(1, len(rows)):
        row = rows[r]
        if not row or not any(row):
            continue

        raw_name = str(row[idx_name]).strip() if idx_name < len(row) else ''
        if not raw_name:
            continue
        if '合计' in raw_name:
            # 如果第一个Sheet自带合计行，保存源表的合计数值以保持完全一致
            def try_get(col_idx):
                if col_idx < len(row):
                    val = row[col_idx]
                    try:
                        return float(val) if val != '' and val is not None else None
                    except (ValueError, TypeError):
                        return None
                return None
            source_totals = {
                'count': try_get(3) or try_get(idx_mfr),
                'qty': try_get(idx_qty),
                'in_amt': try_get(idx_in_amt),
                'retail_amt': try_get(idx_retail_amt)
            }
            continue

        original_count += 1
        name = raw_name
        spec = str(row[idx_spec]).strip() if idx_spec < len(row) else ''
        form = str(row[idx_form]).strip() if idx_form < len(row) else ''
        mfr = str(row[idx_mfr]).strip() if idx_mfr < len(row) else ''
        unit = str(row[idx_unit]).strip() if idx_unit < len(row) else ''

        def parse_float(val):
            try:
                if val == '' or val is None:
                    return 0.0
                return float(val)
            except (ValueError, TypeError):
                return 0.0

        qty = parse_float(row[idx_qty] if idx_qty < len(row) else 0)
        in_amt = parse_float(row[idx_in_amt] if idx_in_amt < len(row) else 0)
        retail_amt = parse_float(row[idx_retail_amt] if idx_retail_amt < len(row) else 0)
        stock_val = parse_float(row[idx_stock] if idx_stock < len(row) else 0)

        key = (name, spec, form, mfr)
        if key not in groups:
            groups[key] = {
                '药品名称': name,
                '规格': spec,
                '剂型': form,
                '制药厂': mfr,
                '单位': unit,
                '数量': 0.0,
                '进价金额': 0.0,
                '零价金额': 0.0,
                '库存': stock_val
            }

        groups[key]['数量'] += qty
        groups[key]['进价金额'] += in_amt
        groups[key]['零价金额'] += retail_amt

    # 排序规则：数量降序，数量相同时按药品名称GBK编码(拼音顺序)
    sorted_records = sorted(
        groups.values(),
        key=lambda x: (-x['数量'], pinyin_sort_key(x['药品名称']))
    )

    total_qty = source_totals['qty'] if source_totals and source_totals['qty'] is not None else sum(item['数量'] for item in sorted_records)
    total_in_amt = source_totals['in_amt'] if source_totals and source_totals['in_amt'] is not None else sum(item['进价金额'] for item in sorted_records)
    total_retail_amt = source_totals['retail_amt'] if source_totals and source_totals['retail_amt'] is not None else sum(item['零价金额'] for item in sorted_records)
    final_count = source_totals['count'] if source_totals and source_totals['count'] is not None else original_count

    totals = {
        'original_count': final_count,
        'unique_count': len(sorted_records),
        'total_qty': total_qty,
        'total_in_amt': total_in_amt,
        'total_retail_amt': total_retail_amt
    }

    return sorted_records, totals


def build_xls_styles():
    """构造 xlwt 的样式字典"""
    import xlwt

    # 基础边框
    borders = xlwt.Borders()
    borders.left = xlwt.Borders.THIN
    borders.right = xlwt.Borders.THIN
    borders.top = xlwt.Borders.THIN
    borders.bottom = xlwt.Borders.THIN

    # 对齐
    align_center = xlwt.Alignment()
    align_center.horz = xlwt.Alignment.HORZ_CENTER
    align_center.vert = xlwt.Alignment.VERT_CENTER

    align_left = xlwt.Alignment()
    align_left.horz = xlwt.Alignment.HORZ_LEFT
    align_left.vert = xlwt.Alignment.VERT_CENTER

    # 字体
    font_title = xlwt.Font()
    font_title.name = '宋体'
    font_title.height = 360  # 18pt

    font_body = xlwt.Font()
    font_body.name = '宋体'
    font_body.height = 220  # 11pt

    # 标题样式
    style_title = xlwt.XFStyle()
    style_title.font = font_title
    style_title.alignment = align_center

    # 表头样式
    style_header_left = xlwt.XFStyle()
    style_header_left.font = font_body
    style_header_left.alignment = align_left
    style_header_left.borders = borders

    style_header_center = xlwt.XFStyle()
    style_header_center.font = font_body
    style_header_center.alignment = align_center
    style_header_center.borders = borders

    # 数据行样式
    style_data_left = xlwt.XFStyle()
    style_data_left.font = font_body
    style_data_left.alignment = align_left
    style_data_left.borders = borders

    style_data_center = xlwt.XFStyle()
    style_data_center.font = font_body
    style_data_center.alignment = align_center
    style_data_center.borders = borders

    # 金额样式（带两位小数格式，但保留原精度）
    style_data_money = xlwt.XFStyle()
    style_data_money.font = font_body
    style_data_money.alignment = align_center
    style_data_money.borders = borders
    style_data_money.num_format_str = '0.00_'

    return {
        'title': style_title,
        'header_left': style_header_left,
        'header_center': style_header_center,
        'data_left': style_data_left,
        'data_center': style_data_center,
        'data_money': style_data_money
    }


def process_xls_file(file_path, output_path=None, sheet_name=None):
    """处理 .xls 格式文件"""
    import xlrd
    from xlutils.copy import copy

    rb = xlrd.open_workbook(file_path, formatting_info=True)
    s0 = rb.sheet_by_index(0)
    sheet0_name = s0.name

    # 读取第一个Sheet数据
    rows = []
    for r in range(s0.nrows):
        rows.append([s0.cell_value(r, c) for c in range(s0.ncols)])

    sorted_records, totals = process_sales_data(rows)

    # 复制工作簿以保留第一个Sheet
    wb = copy(rb)

    # 确定第二个Sheet的名称
    if sheet_name:
        target_sheet_name = sheet_name
    elif len(rb.sheet_names()) > 1:
        target_sheet_name = rb.sheet_names()[1]
    else:
        target_sheet_name = '销售明细'

    # 如果工作簿中已经有第二个Sheet，先干净删除它以便重建
    while len(wb._Workbook__worksheets) > 1:
        old_sheet = wb._Workbook__worksheets.pop(1)
        if hasattr(wb, '_Workbook__worksheet_idx_from_name'):
            wb._Workbook__worksheet_idx_from_name.pop(old_sheet.name.lower(), None)

    s1 = wb.add_sheet(target_sheet_name, cell_overwrite_ok=True)

    styles = build_xls_styles()

    # 1. 标题行 (Row 0)
    # 标题内容，例如 "2026.8月西药销售明细表"
    if '销售表' in sheet0_name:
        title_text = sheet0_name.replace('销售表', '销售明细表')
    else:
        title_text = f"{sheet0_name}销售明细表"

    s1.row(0).height = 540
    s1.write_merge(0, 0, 0, 8, title_text, styles['title'])

    # 2. 表头行 (Row 1)
    s1.row(1).height = 360
    for col_idx, col_name in enumerate(TARGET_HEADERS):
        style = styles['header_left'] if col_idx == 0 else styles['header_center']
        s1.write(1, col_idx, col_name, style)

    # 3. 数据行 (Row 2 .. N+1)
    current_row = 2
    for item in sorted_records:
        s1.row(current_row).height = 360
        s1.write(current_row, 0, item['药品名称'], styles['data_left'])
        s1.write(current_row, 1, item['规格'], styles['data_left'])
        s1.write(current_row, 2, item['剂型'], styles['data_center'])
        s1.write(current_row, 3, item['制药厂'], styles['data_left'])
        s1.write(current_row, 4, item['单位'], styles['data_center'])
        s1.write(current_row, 5, item['数量'], styles['data_center'])
        s1.write(current_row, 6, item['进价金额'], styles['data_money'])
        s1.write(current_row, 7, item['零价金额'], styles['data_center'])
        s1.write(current_row, 8, item['库存'], styles['data_center'])
        current_row += 1

    # 4. 合计行
    s1.row(current_row).height = 360
    s1.write(current_row, 0, '合计', styles['data_left'])
    s1.write(current_row, 1, '', styles['data_left'])
    s1.write(current_row, 2, '', styles['data_center'])
    s1.write(current_row, 3, float(totals['original_count']), styles['data_center'])
    s1.write(current_row, 4, '条', styles['data_center'])
    s1.write(current_row, 5, totals['total_qty'], styles['data_center'])
    s1.write(current_row, 6, totals['total_in_amt'], styles['data_center'])
    s1.write(current_row, 7, totals['total_retail_amt'], styles['data_center'])
    s1.write(current_row, 8, '', styles['data_center'])

    # 设置列宽
    for col_idx, width in COL_WIDTHS_XLS.items():
        s1.col(col_idx).width = width

    save_path = output_path or file_path
    wb.save(save_path)
    return target_sheet_name, totals, save_path


def process_xlsx_file(file_path, output_path=None, sheet_name=None):
    """处理 .xlsx 格式文件"""
    import openpyxl
    from openpyxl.styles import Font, Alignment, Border, Side

    wb = openpyxl.load_workbook(file_path)
    s0 = wb.worksheets[0]
    sheet0_name = s0.title

    # 读取第一个Sheet数据
    rows = []
    for r in s0.iter_rows(values_only=True):
        rows.append(list(r))

    sorted_records, totals = process_sales_data(rows)

    # 确定第二个Sheet名称
    if sheet_name:
        target_sheet_name = sheet_name
    elif len(wb.worksheets) > 1:
        target_sheet_name = wb.worksheets[1].title
    else:
        target_sheet_name = '销售明细'

    if target_sheet_name in wb.sheetnames:
        idx = wb.sheetnames.index(target_sheet_name)
        wb.remove(wb.worksheets[idx])
        s1 = wb.create_sheet(title=target_sheet_name, index=idx)
    else:
        s1 = wb.create_sheet(title=target_sheet_name, index=1)

    # 样式定义
    thin = Side(border_style="thin", color="000000")
    border_all = Border(top=thin, left=thin, right=thin, bottom=thin)

    font_title = Font(name='宋体', size=18, bold=False)
    font_body = Font(name='宋体', size=11, bold=False)

    align_center = Alignment(horizontal='center', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # 1. 标题行
    if '销售表' in sheet0_name:
        title_text = sheet0_name.replace('销售表', '销售明细表')
    else:
        title_text = f"{sheet0_name}销售明细表"

    s1.row_dimensions[1].height = 27
    s1.merge_cells(start_row=1, start_column=1, end_row=1, end_column=9)
    title_cell = s1.cell(row=1, column=1, value=title_text)
    title_cell.font = font_title
    title_cell.alignment = align_center

    # 2. 表头行
    s1.row_dimensions[2].height = 18
    for col_idx, col_name in enumerate(TARGET_HEADERS, start=1):
        cell = s1.cell(row=2, column=col_idx, value=col_name)
        cell.font = font_body
        cell.border = border_all
        cell.alignment = align_left if col_idx == 1 else align_center

    # 3. 数据行
    current_row = 3
    for item in sorted_records:
        s1.row_dimensions[current_row].height = 18
        row_vals = [
            (item['药品名称'], align_left),
            (item['规格'], align_left),
            (item['剂型'], align_center),
            (item['制药厂'], align_left),
            (item['单位'], align_center),
            (item['数量'], align_center),
            (item['进价金额'], align_center),
            (item['零价金额'], align_center),
            (item['库存'], align_center),
        ]
        for col_idx, (val, align) in enumerate(row_vals, start=1):
            cell = s1.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_body
            cell.border = border_all
            cell.alignment = align
        current_row += 1

    # 4. 合计行
    s1.row_dimensions[current_row].height = 18
    summary_vals = [
        ('合计', align_left),
        ('', align_left),
        ('', align_center),
        (totals['original_count'], align_center),
        ('条', align_center),
        (totals['total_qty'], align_center),
        (totals['total_in_amt'], align_center),
        (totals['total_retail_amt'], align_center),
        ('', align_center),
    ]
    for col_idx, (val, align) in enumerate(summary_vals, start=1):
        cell = s1.cell(row=current_row, column=col_idx, value=val)
        cell.font = font_body
        cell.border = border_all
        cell.alignment = align

    # 设置列宽
    for col_letter, width in COL_WIDTHS_XLSX.items():
        s1.column_dimensions[col_letter].width = width

    save_path = output_path or file_path
    wb.save(save_path)
    return target_sheet_name, totals, save_path


def main():
    parser = argparse.ArgumentParser(description='西药销售表第一个Sheet聚合生成第二个Sheet明细表')
    parser.add_argument('file_path', nargs='?', default=None, help='Excel文件路径（支持 .xls 和 .xlsx）')
    parser.add_argument('-o', '--output', default=None, help='输出文件路径（默认覆盖原文件或更新原文件）')
    parser.add_argument('-s', '--sheet-name', default=None, help='生成的目标Sheet名称（默认销售明细或保持原有）')

    args = parser.parse_args()

    # 如果未指定文件路径，自动寻找当前目录下的销售表文件
    file_path = args.file_path
    if not file_path:
        candidates = [f for f in os.listdir('.') if f.endswith(('.xls', '.xlsx')) and not f.startswith('~$') and not 'backup' in f]
        if candidates:
            # 优先选择包含销售字样的文件
            sales_files = [f for f in candidates if '销售' in f]
            file_path = sales_files[0] if sales_files else candidates[0]
        else:
            print("错误: 当前目录下未找到 Excel 文件 (.xls / .xlsx)，请指定文件路径。")
            sys.exit(1)

    if not os.path.exists(file_path):
        print(f"错误: 文件 '{file_path}' 不存在。")
        sys.exit(1)

    print(f"正在读取并处理文件: {file_path}")

    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.xls':
        sheet_name, totals, out_path = process_xls_file(file_path, args.output, args.sheet_name)
    elif ext == '.xlsx':
        sheet_name, totals, out_path = process_xlsx_file(file_path, args.output, args.sheet_name)
    else:
        print(f"不支持的文件类型: {ext}")
        sys.exit(1)

    print("\n====== 处理完成 ======")
    print(f"已生成目标Sheet: 【{sheet_name}】")
    print(f"输出保存文件: {out_path}")
    print(f"原始记录总数: {totals['original_count']} 条")
    print(f"聚合后品种数: {totals['unique_count']} 种")
    print(f"销售数量总和: {totals['total_qty']}")
    print(f"进价金额总计: {totals['total_in_amt']:.4f}")
    print(f"零价金额总计: {totals['total_retail_amt']:.2f}")


if __name__ == '__main__':
    main()
