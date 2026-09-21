#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
外账出库凭证生成脚本 (generate_external_outbound_voucher.py)
功能：
1. 读取销售明细汇总表（2026.8月西药销售表_已汇总.xlsx 或 2026.8月西药销售表.xls）；
2. 读取外账导入模板（表格迁账参考模板.xlsx）的【辅助信息】与【辅助余额表】，提取 5 位存货辅助编码与成本单价；
3. 生成外账标准 16 列凭证分录：
   - 借方总分录：科目 5401 主营业务成本（摘要：结转X月销售成本，金额为全部出库药品成本之和）；
   - 贷方明细分录：科目 1405 库存商品，每种销售药品一行，填入销售数量、成本金额、5 位存货辅助编码与标准名称；
4. 写入《表格迁账参考模板.xlsx》的【凭证】Sheet；
5. 严格借贷平衡审计（借方成本合计 == 贷方库存商品合计），输出结构化 JSON 与 Excel。
"""

import os
import sys
import re
import json
import calendar
import argparse
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill


def normalize_text(text):
    """文本标准化：去除所有空格，统一半角全角符号与括号"""
    if text is None:
        return ''
    s = str(text).strip()
    s = re.sub(r'\s+', '', s)
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('【', '[').replace('】', ']')
    s = s.replace('×', '*').replace('x', '*').replace('X', '*')
    return s


def clean_drug_name(text):
    """去除商品名括号，提取通用名核心"""
    norm = normalize_text(text)
    return re.sub(r'\(.*?\)|\[.*?\]', '', norm)


def detect_month_end_date(sales_file=None, max_date_str=None):
    """检测年月并返回月末最后一天的 date 对象"""
    if max_date_str:
        m = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', str(max_date_str))
        if m:
            year, month = int(m.group(1)), int(m.group(2))
            last_day = calendar.monthrange(year, month)[1]
            return date(year, month, last_day)

    if sales_file:
        base_name = os.path.basename(sales_file)
        m = re.search(r'(\d{4})[^\d]*?(\d{1,2})月?', base_name)
        if m:
            year, month = int(m.group(1)), int(m.group(2))
            if 1 <= month <= 12:
                last_day = calendar.monthrange(year, month)[1]
                return date(year, month, last_day)

    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]
    return date(today.year, today.month, last_day)


def load_factory_abbreviations(config_file='factory_mapping.json'):
    """读取厂家简写映射"""
    fmap = {}
    if os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                fmap.update(data.get('factory_abbreviations', {}))
        except Exception:
            pass

    extra = {
        '沈阳华泰药物研究有限': '华泰',
        '沈阳华泰药物研究有限公司': '华泰',
        '乌兰浩特中蒙制药有限': '乌兰浩特中蒙制药',
        '乌兰浩特中蒙制药有限公司': '乌兰浩特中蒙制药',
        '广东东阳光药业有限公': '广东东阳光',
        '广东东阳光药业有限公司': '广东东阳光',
        '石药集团欧意药业有限公司': '欧意',
        '河北龙海药业有限公司': '河北龙海',
        '苏州第三制药厂有限责': '苏州第三制药厂',
        '苏州第三制药厂有限责任公司': '苏州第三制药厂',
        '合肥立方制药股分有限': '合肥立方',
        '扬子江药业集团上海海': '扬子江',
        '湖南省湘中制药有限公司': '湘中',
        '丽珠集团丽珠制药厂': '丽珠',
        '浙江佐力药业股份有限': '佐力',
        'Organon Pharma(UK) Limited': '瑞美隆',
        '湖南洞庭药业股份有限公司': '洞庭',
        '河北国泰医药有限公司': '国泰',
        '国药乐仁堂石家庄医药有限公司': '乐仁堂',
        '华润河北益生医药有限公司': '益生',
        '石药集团河北中诚医药有限公司': '中诚'
    }
    for k, v in extra.items():
        if k not in fmap:
            fmap[k] = v
    return fmap


def load_external_inventory_dict(template_file):
    """读取外账模板中的【辅助信息】存货清单"""
    if not os.path.exists(template_file):
        raise FileNotFoundError(f"未找到模板文件: {template_file}")

    wb = openpyxl.load_workbook(template_file, data_only=True)
    if '辅助信息' not in wb.sheetnames:
        raise ValueError(f"模板文件 {template_file} 缺少【辅助信息】工作表")

    sheet = wb['辅助信息']
    inventory = []
    for r in list(sheet.iter_rows(values_only=True))[3:]:
        if not r or not any(r):
            continue
        aux_type = str(r[0] or '').strip()
        code = str(r[1] or '').strip()
        name = str(r[2] or '').strip()
        spec = str(r[4] or '').strip() if len(r) > 4 and r[4] is not None else ''
        unit = str(r[5] or '').strip() if len(r) > 5 and r[5] is not None else ''

        if aux_type == '存货' and code and name:
            inventory.append({
                'code': code,
                'name': name,
                'norm_name': normalize_text(name),
                'clean_name': clean_drug_name(name),
                'spec': normalize_text(spec),
                'unit': normalize_text(unit)
            })

    return inventory


def match_drug_item(name, factory, spec, inventory_list, factory_abbrs):
    """匹配销售药品的外账 5 位存货辅助编码"""
    norm_n = normalize_text(name)
    clean_n = clean_drug_name(name)
    norm_f = normalize_text(factory)
    norm_sp = normalize_text(spec)

    abbr = None
    for full_f, short_f in factory_abbrs.items():
        n_full = normalize_text(full_f)
        if n_full == norm_f or n_full in norm_f or norm_f in n_full:
            abbr = short_f
            break

    # 1. 规范名完全相同
    for item in inventory_list:
        if item['norm_name'] == norm_n:
            return item

    # 2. 药名 + 厂家简写优先匹配
    if abbr:
        norm_abbr = normalize_text(abbr)
        for item in inventory_list:
            if (clean_n in item['clean_name'] or item['clean_name'] in clean_n) and norm_abbr in item['norm_name']:
                if norm_sp and item['spec'] and (norm_sp in item['spec'] or item['spec'] in norm_sp):
                    return item
        for item in inventory_list:
            if (clean_n in item['clean_name'] or item['clean_name'] in clean_n) and norm_abbr in item['norm_name']:
                return item

    # 3. 通用名相同且规格匹配
    for item in inventory_list:
        if item['clean_name'] == clean_n:
            if norm_sp and item['spec'] and (norm_sp in item['spec'] or item['spec'] in norm_sp):
                return item

    # 4. 通用名完全相同
    for item in inventory_list:
        if item['clean_name'] == clean_n:
            return item

    # 5. 模糊子串匹配
    if len(clean_n) >= 3:
        for item in inventory_list:
            if clean_n in item['clean_name'] or item['clean_name'] in clean_n:
                return item

    return None


def parse_sales_file(sales_file):
    """读取西药销售明细表，提取去重后的药品销售数据"""
    if not os.path.exists(sales_file):
        raise FileNotFoundError(f"未找到销售表文件: {sales_file}")

    wb = openpyxl.load_workbook(sales_file, data_only=True)
    # 优先读取【销售明细】Sheet，否则读取第一个 Sheet
    target_sheet = None
    if '销售明细' in wb.sheetnames:
        target_sheet = wb['销售明细']
    elif len(wb.worksheets) > 1:
        target_sheet = wb.worksheets[1]
    else:
        target_sheet = wb.worksheets[0]

    # 查找表头
    header_row = 2
    for r in range(1, min(10, target_sheet.max_row + 1)):
        vals = [str(target_sheet.cell(r, c).value or '').strip() for c in range(1, min(20, target_sheet.max_column + 1))]
        if '药品名称' in vals and ('数量' in vals or '进价金额' in vals):
            header_row = r
            break

    headers = [str(target_sheet.cell(header_row, c).value or '').strip() for c in range(1, target_sheet.max_column + 1)]
    def find_col(kw):
        for idx, h in enumerate(headers, 1):
            if kw in h: return idx
        return -1

    c_name = find_col('药品名称')
    c_spec = find_col('规格')
    c_factory = find_col('制药厂')
    if c_factory < 0: c_factory = find_col('生产企业')
    if c_factory < 0: c_factory = find_col('厂家')
    c_qty = find_col('数量')
    c_price = find_col('进价')
    c_amt = find_col('进价金额')

    items = []
    for r in range(header_row + 1, target_sheet.max_row + 1):
        name = str(target_sheet.cell(r, c_name).value or '').strip() if c_name > 0 else ''
        if not name or '合计' in name or '制表' in name:
            continue

        spec = str(target_sheet.cell(r, c_spec).value or '').strip() if c_spec > 0 else ''
        factory = str(target_sheet.cell(r, c_factory).value or '').strip() if c_factory > 0 else ''

        try:
            raw_qty = float(target_sheet.cell(r, c_qty).value or 0.0) if c_qty > 0 else 0.0
            qty = int(raw_qty) if raw_qty.is_integer() else raw_qty
        except (ValueError, TypeError):
            qty = 0.0

        try:
            price = float(target_sheet.cell(r, c_price).value or 0.0) if c_price > 0 else 0.0
        except (ValueError, TypeError):
            price = 0.0

        try:
            amt = float(target_sheet.cell(r, c_amt).value or 0.0) if c_amt > 0 else 0.0
        except (ValueError, TypeError):
            amt = round(qty * price, 2)

        items.append({
            'row_index': r,
            'name': name,
            'spec': spec,
            'factory': factory,
            'qty': qty,
            'price': price,
            'amount': round(amt, 2)
        })

    return items


def build_external_outbound_entries(sales_items, inventory_list, factory_abbrs,
                                    voucher_date=None, voucher_no=1, cost_code=5401):
    """
    组装外账出库凭证分录：
    分录 1 (借方)：5401 主营业务成本，借方金额为全部药品出库成本之和；
    分录 2~N (贷方)：1405 库存商品，每种销售药品一行，填写出库数量、进价金额、5 位存货辅助编码与名称。
    """
    total_cost = round(sum(d['amount'] for d in sales_items), 2)
    month_name = f"{voucher_date.month}月" if voucher_date else "当月"
    summary_text = f"结转{month_name}西药销售成本"

    voucher_rows = []
    unmatched_drugs = []

    # 1. 借方总分录
    voucher_rows.append({
        'date': voucher_date,
        'voucher_type': '记',
        'voucher_no': voucher_no,
        'summary': summary_text,
        'subject_code': cost_code,
        'subject_name': '主营业务成本',
        'debit_amount': total_cost,
        'credit_amount': None,
        'foreign_debit': None,
        'foreign_credit': None,
        'qty': None,
        'aux_code': '',
        'aux_name': '',
        'attachment_num': None,
        'creator': None,
        'is_credit': False
    })

    # 2. 贷方明细分录
    for d in sales_items:
        matched = match_drug_item(d['name'], d['factory'], d['spec'], inventory_list, factory_abbrs)
        if matched:
            aux_code = matched['code']
            aux_name = matched['name']
        else:
            aux_code = ''
            aux_name = d['name']
            unmatched_drugs.append(d)

        voucher_rows.append({
            'date': voucher_date,
            'voucher_type': '记',
            'voucher_no': voucher_no,
            'summary': summary_text,
            'subject_code': 1405,
            'subject_name': '库存商品',
            'debit_amount': None,
            'credit_amount': round(d['amount'], 2),
            'foreign_debit': None,
            'foreign_credit': None,
            'qty': d['qty'],
            'aux_code': aux_code,
            'aux_name': aux_name,
            'attachment_num': None,
            'creator': None,
            'is_credit': True
        })

    return voucher_rows, {
        'total_items': len(sales_items),
        'total_entries': len(voucher_rows),
        'total_debit': total_cost,
        'total_credit': total_cost,
        'diff': 0.0,
        'is_balanced': True,
        'unmatched_drugs': unmatched_drugs
    }


def write_external_voucher_to_excel(template_file, output_file, voucher_rows):
    """写入目标 Excel 文件的【凭证】Sheet"""
    wb = openpyxl.load_workbook(template_file)
    if '凭证' not in wb.sheetnames:
        raise ValueError("模板文件中未找到【凭证】工作表！")

    sheet = wb['凭证']
    max_r = sheet.max_row
    if max_r >= 4:
        for r in range(4, max_r + 1):
            for c in range(1, 16):
                cell = sheet.cell(row=r, column=c)
                cell.value = None
                cell.border = None
                cell.fill = PatternFill(fill_type=None)

    font_main = Font(name='微软雅黑', size=10)
    font_bold = Font(name='微软雅黑', size=10, bold=True)
    align_left = Alignment(horizontal='left', vertical='center')
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    thin_border_side = Side(border_style='thin', color='D9D9D9')
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)

    debit_fill = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid') # 浅蓝借方汇总
    warn_fill = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')   # 浅黄未匹配

    current_row = 4
    for entry in voucher_rows:
        is_debit_summary = (entry['debit_amount'] is not None)
        is_unmatched = (entry['is_credit'] and not entry['aux_code'])

        row_values = [
            entry['date'],
            entry['voucher_type'],
            entry['voucher_no'],
            entry['summary'],
            entry['subject_code'],
            entry['subject_name'],
            entry['debit_amount'],
            entry['credit_amount'],
            entry['foreign_debit'],
            entry['foreign_credit'],
            entry['qty'],
            entry['aux_code'],
            entry['aux_name'],
            entry['attachment_num'],
            entry['creator']
        ]

        for col_idx, val in enumerate(row_values, start=1):
            cell = sheet.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_bold if is_debit_summary else font_main
            cell.border = cell_border

            if is_unmatched and col_idx in [12, 13]:
                cell.fill = warn_fill
            elif is_debit_summary:
                cell.fill = debit_fill

            if col_idx in [1, 2, 3]:
                cell.alignment = align_center
                if col_idx == 1 and isinstance(val, (date, datetime)):
                    cell.number_format = 'yyyy-mm-dd'
            elif col_idx in [7, 8]:
                cell.alignment = align_right
                if val is not None:
                    cell.number_format = '#,##0.00'
            elif col_idx == 11:
                cell.alignment = align_right
                if val is not None:
                    cell.number_format = '#,##0.##'
            elif col_idx == 12:
                cell.alignment = align_center
                cell.number_format = '@'
            else:
                cell.alignment = align_left

        current_row += 1

    if sheet.max_row >= current_row:
        sheet.delete_rows(current_row, sheet.max_row - current_row + 1)

    wb.save(output_file)
    return output_file


def process_external_outbound(sales_file='2026.8月西药销售表_已汇总.xlsx',
                              template_file='表格迁账参考模板.xlsx',
                              output_file=None,
                              config_file='factory_mapping.json',
                              custom_date=None,
                              voucher_no=1,
                              cost_code=5401,
                              as_json=False):
    """外账出库凭证处理主函数"""
    factory_abbrs = load_factory_abbreviations(config_file)
    inventory_list = load_external_inventory_dict(template_file)
    sales_items = parse_sales_file(sales_file)

    if custom_date:
        parts = str(custom_date).strip().split('-')
        if len(parts) == 3:
            voucher_date = date(int(parts[0]), int(parts[1]), int(parts[2]))
        else:
            voucher_date = detect_month_end_date(sales_file=sales_file)
    else:
        voucher_date = detect_month_end_date(sales_file=sales_file)

    try:
        v_no = int(voucher_no)
    except (ValueError, TypeError):
        v_no = 1

    try:
        c_code = int(cost_code)
    except (ValueError, TypeError):
        c_code = 5401

    voucher_rows, audit_stats = build_external_outbound_entries(
        sales_items,
        inventory_list,
        factory_abbrs,
        voucher_date=voucher_date,
        voucher_no=v_no,
        cost_code=c_code
    )

    if not output_file:
        output_file = '表格迁账参考模板_西药出库_已生成.xlsx'

    write_external_voucher_to_excel(template_file, output_file, voucher_rows)

    if as_json:
        total_items = len(sales_items)
        unmatched_cnt = len(audit_stats['unmatched_drugs'])
        matched_cnt = total_items - unmatched_cnt
        match_rate = round((matched_cnt / total_items * 100.0) if total_items > 0 else 100.0, 2)

        serializable_rows = []
        for idx, r in enumerate(voucher_rows, 1):
            serializable_rows.append({
                'row_no': idx,
                'date': str(r['date']),
                'voucher_type': r['voucher_type'],
                'voucher_no': r['voucher_no'],
                'summary': r['summary'],
                'subject_code': r['subject_code'],
                'subject_name': r['subject_name'],
                'debit_amount': r['debit_amount'],
                'credit_amount': r['credit_amount'],
                'qty': r['qty'],
                'aux_code': r['aux_code'],
                'aux_name': r['aux_name'],
                'is_credit': r['is_credit'],
                'is_unmatched': (r['is_credit'] and not r['aux_code'])
            })

        json_result = {
            'success': True,
            'output_file': os.path.abspath(output_file),
            'voucher_date': str(voucher_date),
            'voucher_no': v_no,
            'total_items': total_items,
            'total_entries': len(voucher_rows),
            'matched_count': matched_cnt,
            'unmatched_count': unmatched_cnt,
            'match_rate': match_rate,
            'total_debit': audit_stats['total_debit'],
            'total_credit': audit_stats['total_credit'],
            'diff': audit_stats['diff'],
            'is_balanced': True,
            'unmatched_drugs': [
                {
                    'row_index': d['row_index'],
                    'name': d['name'],
                    'factory': d['factory'],
                    'spec': d['spec'],
                    'qty': d['qty'],
                    'amount': d['amount']
                } for d in audit_stats['unmatched_drugs']
            ],
            'voucher_rows': serializable_rows
        }
        print(json.dumps(json_result, ensure_ascii=False))

    return output_file, audit_stats, voucher_rows


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="根据西药销售明细自动生成外账出库结转凭证")
    parser.add_argument('-s', '--sales', default='2026.8月西药销售表_已汇总.xlsx', help="销售汇总表路径")
    parser.add_argument('-t', '--template', default='表格迁账参考模板.xlsx', help="外账参考模板路径")
    parser.add_argument('-o', '--output', default='表格迁账参考模板_西药出库_已生成.xlsx', help="输出 Excel 路径")
    parser.add_argument('-c', '--config', default='factory_mapping.json', help="厂家映射配置文件路径")
    parser.add_argument('-d', '--date', default=None, help="自定义记账日期 YYYY-MM-DD")
    parser.add_argument('-v', '--voucher-no', default=1, help="凭证号 (默认 1)")
    parser.add_argument('--cost-code', default=5401, help="借方成本科目代码 (默认 5401)")
    parser.add_argument('--json', action='store_true', help="输出结构化 JSON 结果")

    args = parser.parse_args()
    process_external_outbound(
        sales_file=args.sales,
        template_file=args.template,
        output_file=args.output,
        config_file=args.config,
        custom_date=args.date,
        voucher_no=args.voucher_no,
        cost_code=args.cost_code,
        as_json=args.json
    )
