#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
石家庄心理医院 - 账实库存智能核对脚本 (compare_inventory.py)
功能：
1. 读取财务系统入库出库完毕后的【数量金额总账】（.xlsx）。
2. 读取库管系统导出的【西药 / 中药 / 耗材 库存汇总报表】（.xls 或 .xlsx）。
3. 支持西药、中药、耗材三库独立或联合批量比对。
4. 结合厂家映射与规格清洗，精准比对财务期末结存数量与库管在库数量。
5. 输出四维状态：完全一致、数量差异、仅财务有、仅库管有。
6. 自动生成标准《账实库存核对分析报告.xlsx》供财务与库管对账调账。
"""

import os
import sys
import re
import json
import argparse
from collections import OrderedDict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import xlrd
except ImportError:
    xlrd = None


def clean_text(text):
    """去除中英文括号、空格、斜杠、横杠、下划线并转小写"""
    return re.sub(r'[\s\(\)（）\-\*\/_]+', '', str(text or '')).lower()


def load_factory_config(config_file=None):
    """加载厂家映射与规则字典"""
    abbr_map = {}
    strict_drugs = []
    if config_file and os.path.exists(config_file):
        try:
            with open(config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                abbr_map = data.get('factory_abbreviations', {}) or data.get('vendor_abbr_map', {})
                strict_drugs = data.get('strict_vendor_suffix_drugs', [])
        except Exception:
            pass
    return abbr_map, strict_drugs


def load_ledger_inventory(ledger_path):
    """
    读取财务数量金额总账，提取期末结存信息
    返回值：按前缀分组的字典 { 'XY': [...], 'ZY': [...], 'HC': [...], 'all': [...] }
    """
    if not os.path.exists(ledger_path):
        raise FileNotFoundError(f"总账文件 '{ledger_path}' 不存在")

    wb = openpyxl.load_workbook(ledger_path, data_only=True)
    ws = wb.active

    items_by_cat = {
        'XY': [],  # 西药
        'ZY': [],  # 中药
        'HC': [],  # 耗材
        'OTHER': []
    }
    all_items = []

    # 表头位于第 3、4 行，数据从第 5 行开始 (第 5 行为 1201 合计)
    for r in range(6, ws.max_row + 1):
        code = str(ws.cell(r, 1).value or '').strip()
        name_full = str(ws.cell(r, 2).value or '').strip()
        unit = str(ws.cell(r, 3).value or '').strip()
        
        if not code.startswith('1201_'):
            continue

        qty_val = ws.cell(r, 17).value
        price_val = ws.cell(r, 18).value
        amt_val = ws.cell(r, 19).value

        try:
            qty = float(qty_val) if qty_val is not None and str(qty_val).strip() != '' else 0.0
        except:
            qty = 0.0
        try:
            price = float(price_val) if price_val is not None and str(price_val).strip() != '' else 0.0
        except:
            price = 0.0
        try:
            amt = float(amt_val) if amt_val is not None and str(amt_val).strip() != '' else 0.0
        except:
            amt = 0.0

        # 解析品名与规格：形如 "存货_阿普唑仑片 0.4mg*100" 或 "存货_鸡血藤（蕴德） 10g*1袋/袋"
        name_clean = name_full
        if name_clean.startswith('存货_'):
            name_clean = name_clean[3:]
        
        parts = name_clean.split(' ', 1)
        drug_name = parts[0].strip()
        spec = parts[1].strip() if len(parts) > 1 else ''

        record = {
            'code': code,
            'raw_name': name_full,
            'name': drug_name,
            'spec': spec,
            'unit': unit,
            'qty': round(qty, 2),
            'price': round(price, 4),
            'amt': round(amt, 2),
            'matched': False
        }

        all_items.append(record)
        if code.startswith('1201_XY'):
            items_by_cat['XY'].append(record)
        elif code.startswith('1201_ZY') or code.startswith('1201_KL'):
            # 中药房包含 1201_ZY (饮片) 和 1201_KL (配方颗粒)
            items_by_cat['ZY'].append(record)
        elif code.startswith('1201_HC'):
            items_by_cat['HC'].append(record)
        else:
            items_by_cat['OTHER'].append(record)

    return items_by_cat, all_items


def load_warehouse_inventory(file_path):
    """
    读取库管系统库存汇总报表（支持 .xls 与 .xlsx）
    表头规范：['药品编码', '数字编码', '药品名称', '规格', '剂型', '制药厂', '单位', '数量', '单价', '金额']
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"库管文件 '{file_path}' 不存在")

    ext = os.path.splitext(file_path)[1].lower()
    raw_rows = []

    if ext == '.xls':
        if xlrd is None:
            raise ImportError("读取 .xls 文件需安装 xlrd 模块")
        try:
            rb = xlrd.open_workbook(file_path)
            s0 = rb.sheet_by_index(0)
            for r in range(s0.nrows):
                raw_rows.append([s0.cell_value(r, c) for c in range(s0.ncols)])
        except Exception as e:
            # 探测是否为 xlsx 改名或 HTML 伪报表
            with open(file_path, 'rb') as f:
                head = f.read(512)
            if head.startswith(b'PK\x03\x04'):
                wb = openpyxl.load_workbook(file_path, data_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    raw_rows.append(list(row))
            elif b'<html' in head.lower() or b'<table' in head.lower() or b'<?xml' in head.lower():
                raise ValueError(f"文件 '{file_path}' 疑似为医院系统导出的 HTML/XML 网页伪表格，并非标准 Excel 二进制文件。\n【解决方法】：请用 WPS 或 Microsoft Excel 打开该报表，点击【文件】->【另存为】，选择【Excel 工作簿 (*.xlsx)】保存后再导入！") from e
            else:
                raise ValueError(f"打开 Excel 文件 '{file_path}' 失败: {e}\n提示：请先用 WPS 或 Excel 打开并“另存为”标准 .xlsx 格式后再导入。") from e
    elif ext == '.xlsx':
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(values_only=True):
                raw_rows.append(list(row))
        except Exception as e:
            if xlrd is not None:
                with open(file_path, 'rb') as f:
                    head = f.read(8)
                if head.startswith(b'\xD0\xCF\x11\xE0'):
                    rb = xlrd.open_workbook(file_path)
                    s0 = rb.sheet_by_index(0)
                    for r in range(s0.nrows):
                        raw_rows.append([s0.cell_value(r, c) for c in range(s0.ncols)])
                else:
                    raise
            else:
                raise
    else:
        raise ValueError(f"不支持的文件格式: {ext}")

    if not raw_rows:
        return []

    # 寻找表头行
    header_idx = -1
    col_map = {}
    for idx, r in enumerate(raw_rows[:5]):
        row_strs = [str(c or '').strip() for c in r]
        if '药品名称' in row_strs or '名称' in row_strs or '物资名称' in row_strs:
            header_idx = idx
            for c_idx, c_val in enumerate(row_strs):
                if c_val in ['药品名称', '名称', '物资名称']:
                    col_map['name'] = c_idx
                elif c_val in ['规格']:
                    col_map['spec'] = c_idx
                elif c_val in ['制药厂', '生产厂家', '厂家']:
                    col_map['factory'] = c_idx
                elif c_val in ['单位']:
                    col_map['unit'] = c_idx
                elif c_val in ['数量']:
                    col_map['qty'] = c_idx
                elif c_val in ['单价']:
                    col_map['price'] = c_idx
                elif c_val in ['金额']:
                    col_map['amt'] = c_idx
                elif c_val in ['药品编码', '物资编码', '编码']:
                    col_map['code'] = c_idx
            break

    if header_idx == -1 or 'name' not in col_map or 'qty' not in col_map:
        # 回退默认位置
        col_map = {'code': 0, 'name': 2, 'spec': 3, 'factory': 5, 'unit': 6, 'qty': 7, 'price': 8, 'amt': 9}
        header_idx = 0

    items = []
    for r in range(header_idx + 1, len(raw_rows)):
        row = raw_rows[r]
        name = str(row[col_map['name']] if col_map.get('name') is not None and len(row) > col_map['name'] else '').strip()
        # 跳过空行或合计行
        if not name or name in ['合计', '总计', '总合计'] or '合计' in name:
            continue

        spec = str(row[col_map['spec']] if col_map.get('spec') is not None and len(row) > col_map['spec'] else '').strip()
        factory = str(row[col_map['factory']] if col_map.get('factory') is not None and len(row) > col_map['factory'] else '').strip()
        unit = str(row[col_map['unit']] if col_map.get('unit') is not None and len(row) > col_map['unit'] else '').strip()
        code = str(row[col_map['code']] if col_map.get('code') is not None and len(row) > col_map['code'] else '').strip()

        qty_val = row[col_map['qty']] if len(row) > col_map['qty'] else 0.0
        price_val = row[col_map['price']] if col_map.get('price') is not None and len(row) > col_map['price'] else 0.0
        amt_val = row[col_map['amt']] if col_map.get('amt') is not None and len(row) > col_map['amt'] else 0.0

        try:
            qty = float(qty_val) if qty_val != '' and qty_val is not None else 0.0
        except:
            qty = 0.0
        try:
            price = float(price_val) if price_val != '' and price_val is not None else 0.0
        except:
            price = 0.0
        try:
            amt = float(amt_val) if amt_val != '' and amt_val is not None else 0.0
        except:
            amt = 0.0

        items.append({
            'code': code,
            'name': name,
            'spec': spec,
            'factory': factory,
            'unit': unit,
            'qty': round(qty, 2),
            'price': round(price, 4),
            'amt': round(amt, 2)
        })

    return items


def compare_single_category(ledger_candidates, wh_items, category_name, factory_config=None):
    """
    针对某一品类（如西药 XY、中药 ZY、耗材 HC）执行精细化比对
    """
    abbr_map, strict_drugs = factory_config or ({}, [])

    matched_records = []
    ledger_unmatched = []
    
    # 建立财务总账索引
    ledger_pool = [it.copy() for it in ledger_candidates]
    tcm_prefixes = ("制", "炒", "麸炒", "炙", "煅", "酒", "醋", "生", "清", "法", "姜", "焦", "蜜炙", "盐", "熟")

    def strip_tcm(name_str):
        for p in tcm_prefixes:
            if name_str.startswith(p) and len(name_str) > len(p):
                return name_str[len(p):]
        return name_str

    for wh_it in wh_items:
        c_wh_name = clean_text(wh_it['name'])
        c_wh_spec = clean_text(wh_it['spec'])
        c_wh_fac = clean_text(wh_it.get('factory', ''))

        # 智能提取厂家简称（三维容错：直接包含简称、包含全称、全称包含截断）
        factory_brief = ''
        if c_wh_fac:
            for full_f, brief_f in abbr_map.items():
                c_full = clean_text(full_f)
                c_brief = clean_text(brief_f)
                if c_wh_fac in c_full or c_full in c_wh_fac or (c_brief and c_brief in c_wh_fac):
                    factory_brief = brief_f
                    break

        best_match = None

        # 1. 优先全词 + 厂家括号简称匹配
        if factory_brief:
            target_with_fac = clean_text(f"{wh_it['name']}{factory_brief}")
            for l_it in ledger_pool:
                if l_it['matched']:
                    continue
                c_l_raw = clean_text(l_it['raw_name'])
                if target_with_fac in c_l_raw and (not c_wh_spec or c_wh_spec in c_l_raw):
                    best_match = l_it
                    break

        # 1.5 财务品名括号厂家反向智能解析
        if not best_match and c_wh_fac:
            for l_it in ledger_pool:
                if l_it['matched']:
                    continue
                m = re.search(r'[（\(]([^\)）]+)[）\)]', l_it['raw_name'])
                if m:
                    tag = clean_text(m.group(1))
                    l_base = clean_text(re.sub(r'[（\(][^\)）]+[）\)]', '', l_it['raw_name']).replace('存货_', '').split()[0])
                    if l_base == c_wh_name and tag in c_wh_fac:
                        best_match = l_it
                        break

        # 2. 检查严格厂家锁定药品（如海螵蛸），若未带指定厂家则禁止盲目回退
        is_strict = any(s in wh_it['name'] for s in strict_drugs)

        if not best_match and not (is_strict and not factory_brief):
            # 规格与名称全匹配
            for l_it in ledger_pool:
                if l_it['matched']:
                    continue
                c_l_raw = clean_text(l_it['raw_name'])
                c_l_name = clean_text(l_it['name'])
                if c_wh_name == c_l_name:
                    if c_wh_spec in c_l_raw or not c_wh_spec:
                        best_match = l_it
                        break

        # 3. 模糊前缀匹配
        if not best_match and not is_strict:
            for l_it in ledger_pool:
                if l_it['matched']:
                    continue
                c_l_raw = clean_text(l_it['raw_name'])
                if c_wh_name in c_l_raw and (c_wh_spec in c_l_raw or not c_wh_spec):
                    best_match = l_it
                    break

        # 4. 中药炮制前缀智能兼容匹配 (如 "吴茱萸" 匹配 "制吴茱萸")
        if not best_match and not is_strict:
            stripped_wh = strip_tcm(c_wh_name)
            for l_it in ledger_pool:
                if l_it['matched']:
                    continue
                c_l_name = clean_text(l_it['name'])
                stripped_l = strip_tcm(c_l_name)
                if stripped_wh == stripped_l or c_wh_name == stripped_l or stripped_wh == c_l_name:
                    best_match = l_it
                    break

        if best_match:
            best_match['matched'] = True
            diff_qty = round(best_match['qty'] - wh_it['qty'], 2)
            diff_amt = round(best_match['amt'] - wh_it['amt'], 2)
            status = 'EQUAL' if diff_qty == 0 else 'DIFF_QTY'
            
            matched_records.append({
                'category': category_name,
                'status': status,
                'status_desc': '数量完全吻合' if status == 'EQUAL' else '存在数量差异',
                'name': wh_it['name'],
                'spec': wh_it['spec'],
                'factory': wh_it['factory'],
                'unit': wh_it['unit'] or best_match['unit'],
                'ledger_code': best_match['code'],
                'ledger_name': best_match['raw_name'],
                'ledger_qty': best_match['qty'],
                'wh_qty': wh_it['qty'],
                'diff_qty': diff_qty,
                'ledger_price': best_match['price'],
                'wh_price': wh_it['price'],
                'ledger_amt': best_match['amt'],
                'wh_amt': wh_it['amt'],
                'diff_amt': diff_amt,
            })
        else:
            matched_records.append({
                'category': category_name,
                'status': 'WH_ONLY',
                'status_desc': '仅库管有/财务未建账',
                'name': wh_it['name'],
                'spec': wh_it['spec'],
                'factory': wh_it['factory'],
                'unit': wh_it['unit'],
                'ledger_code': '-',
                'ledger_name': '-',
                'ledger_qty': 0.0,
                'wh_qty': wh_it['qty'],
                'diff_qty': round(-wh_it['qty'], 2),
                'ledger_price': 0.0,
                'wh_price': wh_it['price'],
                'ledger_amt': 0.0,
                'wh_amt': wh_it['amt'],
                'diff_amt': round(-wh_it['amt'], 2),
            })

    # 4. 统计财务总账单边有结存、库管未列入的品种
    for l_it in ledger_pool:
        if not l_it['matched'] and (l_it['qty'] != 0 or l_it['amt'] != 0):
            matched_records.append({
                'category': category_name,
                'status': 'LEDGER_ONLY',
                'status_desc': '仅财务有结存/库管已空',
                'name': l_it['name'],
                'spec': l_it['spec'],
                'factory': '-',
                'unit': l_it['unit'],
                'ledger_code': l_it['code'],
                'ledger_name': l_it['raw_name'],
                'ledger_qty': l_it['qty'],
                'wh_qty': 0.0,
                'diff_qty': l_it['qty'],
                'ledger_price': l_it['price'],
                'wh_price': 0.0,
                'ledger_amt': l_it['amt'],
                'wh_amt': 0.0,
                'diff_amt': l_it['amt'],
            })

    # 统计分类摘要
    total_items = len(matched_records)
    equal_count = sum(1 for r in matched_records if r['status'] == 'EQUAL')
    diff_count = sum(1 for r in matched_records if r['status'] == 'DIFF_QTY')
    wh_only_count = sum(1 for r in matched_records if r['status'] == 'WH_ONLY')
    ledger_only_count = sum(1 for r in matched_records if r['status'] == 'LEDGER_ONLY')
    
    total_ledger_amt = round(sum(r['ledger_amt'] for r in matched_records), 2)
    total_wh_amt = round(sum(r['wh_amt'] for r in matched_records), 2)
    total_diff_amt = round(total_ledger_amt - total_wh_amt, 2)
    
    match_rate = round(equal_count / total_items * 100, 1) if total_items > 0 else 0.0

    return {
        'category': category_name,
        'summary': {
            'total_items': total_items,
            'equal_count': equal_count,
            'diff_count': diff_count,
            'wh_only_count': wh_only_count,
            'ledger_only_count': ledger_only_count,
            'match_rate': match_rate,
            'total_ledger_amt': total_ledger_amt,
            'total_wh_amt': total_wh_amt,
            'total_diff_amt': total_diff_amt,
        },
        'records': matched_records
    }


def generate_comparison_excel(audit_results, output_path):
    """
    生成高规格格式化的《账实库存核对分析报告.xlsx》
    """
    wb = openpyxl.Workbook()
    
    # 样式定义
    header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
    title_font = Font(name='微软雅黑', size=14, bold=True, color='1F2937')
    section_font = Font(name='微软雅黑', size=12, bold=True, color='0F766E')
    data_font = Font(name='微软雅黑', size=10)
    mono_font = Font(name='Consolas', size=10)
    
    fill_header = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
    fill_sub_header = PatternFill(start_color='334155', end_color='334155', fill_type='solid')
    fill_diff = PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid')      # 浅红高亮差异
    fill_warn = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')      # 浅橙高亮单边
    fill_equal = PatternFill(start_color='ECFDF5', end_color='ECFDF5', fill_type='solid')     # 浅绿正常
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    # ----------------------------------------------------
    # Sheet 1: 核对看板总览
    # ----------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = "账实核对看板"
    ws_dash.views.sheetView[0].showGridLines = True

    ws_dash.merge_cells('A1:H1')
    ws_dash['A1'] = "石家庄心理医院 - 财务总账与库管库存账实核对报告"
    ws_dash['A1'].font = title_font
    ws_dash['A1'].alignment = Alignment(vertical='center')
    ws_dash.row_dimensions[1].height = 40

    headers_dash = ['分类名称', '比对品规总数', '完全吻合数', '数量差异数', '仅库管有数', '仅财务有数', '数量一致率', '账实差额(元)']
    ws_dash.row_dimensions[3].height = 26
    for c_idx, h in enumerate(headers_dash, 1):
        cell = ws_dash.cell(row=3, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    curr_r = 4
    for res in audit_results:
        s = res['summary']
        row_vals = [
            res['category'],
            s['total_items'],
            s['equal_count'],
            s['diff_count'],
            s['wh_only_count'],
            s['ledger_only_count'],
            f"{s['match_rate']}%",
            f"¥ {s['total_diff_amt']:,.2f}"
        ]
        ws_dash.row_dimensions[curr_r].height = 22
        for c_idx, val in enumerate(row_vals, 1):
            cell = ws_dash.cell(row=curr_r, column=c_idx, value=val)
            cell.font = data_font
            cell.border = thin_border
            if c_idx == 1:
                cell.alignment = Alignment(horizontal='left', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='right', vertical='center')
                if c_idx == 7 and s['match_rate'] >= 90:
                    cell.fill = fill_equal
                if c_idx == 4 and s['diff_count'] > 0:
                    cell.fill = fill_diff
        curr_r += 1

    for col in ws_dash.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws_dash.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 4, 14)

    # ----------------------------------------------------
    # Sheet 2: 差异重点明细清单 (重点排查表)
    # ----------------------------------------------------
    ws_diff = wb.create_sheet(title="数量差异排查清单")
    ws_diff.views.sheetView[0].showGridLines = True

    diff_headers = [
        '分类', '核对状态', '药品/物资名称', '规格', '生产厂家', '单位',
        '财务科目编码', '财务总账数量', '库管在库数量', '数量差额',
        '财务结存金额', '库管在库金额', '金额差额'
    ]

    ws_diff.row_dimensions[1].height = 26
    for c_idx, h in enumerate(diff_headers, 1):
        cell = ws_diff.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    diff_row = 2
    for res in audit_results:
        for r in res['records']:
            if r['status'] == 'EQUAL':
                continue  # 差异表只留差异项与单边项

            ws_diff.row_dimensions[diff_row].height = 20
            vals = [
                r['category'],
                r['status_desc'],
                r['name'],
                r['spec'],
                r['factory'],
                r['unit'],
                r['ledger_code'],
                r['ledger_qty'],
                r['wh_qty'],
                r['diff_qty'],
                r['ledger_amt'],
                r['wh_amt'],
                r['diff_amt'],
            ]

            fill_to_use = fill_diff if r['status'] == 'DIFF_QTY' else fill_warn
            for c_idx, v in enumerate(vals, 1):
                cell = ws_diff.cell(row=diff_row, column=c_idx, value=v)
                cell.font = mono_font if c_idx in [8, 9, 10, 11, 12, 13] else data_font
                cell.border = thin_border
                cell.fill = fill_to_use
                if c_idx in [8, 9, 10, 11, 12, 13]:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            diff_row += 1

    for col in ws_diff.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws_diff.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)

    # ----------------------------------------------------
    # Sheet 3: 全品种对账明细底稿
    # ----------------------------------------------------
    ws_all = wb.create_sheet(title="全部品规对照底稿")
    ws_all.views.sheetView[0].showGridLines = True

    all_headers = diff_headers
    ws_all.row_dimensions[1].height = 26
    for c_idx, h in enumerate(all_headers, 1):
        cell = ws_all.cell(row=1, column=c_idx, value=h)
        cell.font = header_font
        cell.fill = fill_sub_header
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = thin_border

    all_row = 2
    for res in audit_results:
        for r in res['records']:
            ws_all.row_dimensions[all_row].height = 20
            vals = [
                r['category'],
                r['status_desc'],
                r['name'],
                r['spec'],
                r['factory'],
                r['unit'],
                r['ledger_code'],
                r['ledger_qty'],
                r['wh_qty'],
                r['diff_qty'],
                r['ledger_amt'],
                r['wh_amt'],
                r['diff_amt'],
            ]

            fill_to_use = fill_equal if r['status'] == 'EQUAL' else (fill_diff if r['status'] == 'DIFF_QTY' else fill_warn)
            for c_idx, v in enumerate(vals, 1):
                cell = ws_all.cell(row=all_row, column=c_idx, value=v)
                cell.font = mono_font if c_idx in [8, 9, 10, 11, 12, 13] else data_font
                cell.border = thin_border
                cell.fill = fill_to_use
                if c_idx in [8, 9, 10, 11, 12, 13]:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
            all_row += 1

    for col in ws_all.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        ws_all.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)

    wb.save(output_path)
    return output_path


def run_audit(ledger_file, west_wh_file=None, tcm_wh_file=None, hc_wh_file=None, config_file=None, output_file=None):
    """
    执行完整的账实多库对账审计主流程
    """
    factory_config = load_factory_config(config_file)
    ledger_by_cat, all_ledger = load_ledger_inventory(ledger_file)

    results = []

    # 1. 西药比对
    if west_wh_file and os.path.exists(west_wh_file):
        wh_west = load_warehouse_inventory(west_wh_file)
        res_west = compare_single_category(ledger_by_cat['XY'], wh_west, '西药房', factory_config)
        results.append(res_west)

    # 2. 中药比对
    if tcm_wh_file and os.path.exists(tcm_wh_file):
        wh_tcm = load_warehouse_inventory(tcm_wh_file)
        res_tcm = compare_single_category(ledger_by_cat['ZY'], wh_tcm, '中药房', factory_config)
        results.append(res_tcm)

    # 3. 耗材比对
    if hc_wh_file and os.path.exists(hc_wh_file):
        wh_hc = load_warehouse_inventory(hc_wh_file)
        res_hc = compare_single_category(ledger_by_cat['HC'], wh_hc, '耗材库', factory_config)
        results.append(res_hc)

    # 若未指定单独库管表，但指定了单个库管表且无法判定类别，自动全账比对
    if not results and west_wh_file and os.path.exists(west_wh_file):
        wh_general = load_warehouse_inventory(west_wh_file)
        res_gen = compare_single_category(all_ledger, wh_general, '综合库房', factory_config)
        results.append(res_gen)

    # 生成 Excel 报告
    if not output_file:
        base_dir = os.path.dirname(ledger_file) or '.'
        output_file = os.path.join(base_dir, '账实库存核对分析报告_已生成.xlsx')

    report_path = generate_comparison_excel(results, output_file)

    # 汇总总数据
    grand_total_items = sum(r['summary']['total_items'] for r in results)
    grand_equal = sum(r['summary']['equal_count'] for r in results)
    grand_diff = sum(r['summary']['diff_count'] for r in results)
    grand_wh_only = sum(r['summary']['wh_only_count'] for r in results)
    grand_ledger_only = sum(r['summary']['ledger_only_count'] for r in results)
    grand_diff_amt = round(sum(r['summary']['total_diff_amt'] for r in results), 2)
    overall_match_rate = round(grand_equal / grand_total_items * 100, 1) if grand_total_items > 0 else 0.0

    return {
        'success': True,
        'output_file': os.path.abspath(report_path),
        'overall': {
            'total_items': grand_total_items,
            'equal_count': grand_equal,
            'diff_count': grand_diff,
            'wh_only_count': grand_wh_only,
            'ledger_only_count': grand_ledger_only,
            'match_rate': overall_match_rate,
            'total_diff_amt': grand_diff_amt,
        },
        'categories': results
    }


def main():
    parser = argparse.ArgumentParser(description="财务数量金额总账与库管系统库存汇总核对工具")
    parser.add_argument("-l", "--ledger", required=True, help="财务系统数量金额总账表 (.xlsx)")
    parser.add_argument("-w", "--west", default=None, help="西药房库存汇总表 (.xls / .xlsx)")
    parser.add_argument("-t", "--tcm", default=None, help="中药房库存汇总表 (.xls / .xlsx)")
    parser.add_argument("-c", "--hc", default=None, help="耗材库库存汇总表 (.xls / .xlsx)")
    parser.add_argument("--config", default="factory_mapping.json", help="厂家映射配置文件")
    parser.add_argument("-o", "--output", default=None, help="输出 Excel 报告路径")

    args = parser.parse_args()

    res = run_audit(
        ledger_file=args.ledger,
        west_wh_file=args.west,
        tcm_wh_file=args.tcm,
        hc_wh_file=args.hc,
        config_file=args.config,
        output_file=args.output
    )

    print(json.dumps(res, ensure_ascii=False, indent=2))


if __name__ == '__main__':
    main()
