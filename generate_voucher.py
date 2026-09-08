#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
根据西药销售明细与总账表自动生成凭证导入模板
功能：
1. 读取销售表第二个Sheet（销售明细表）中整理好的药品销售数据（药品名称、规格、销售数量）。
2. 读取总账表（石家庄心理医院_数量金额总账），通过【存货_药品名称 规格】匹配对应的科目。
3. 从总账表提取单价与辅助账编码（例如 1201_XY0001 中的 XY0001）。
4. 按照会计凭证规范填充【凭证导入模板.xlsx】的【凭证模版】Sheet：
   - 日期：月末最后一天（例如 2026-08-31）
   - 凭证字：固定为“记”
   - 凭证号/附件数/摘要/科目代码/科目名称/借方金额：留空
   - 分录序号：从 1 递增
   - 贷方金额：数量 * 单价
   - 存货：填入辅助账编码（对应模板 Sheet 3 辅助核算编码）
   - 数量：销售数量
   - 单价：总账单价
   - 原币金额：数量 * 单价
   - 币别：固定为“RMB”
   - 汇率：固定为 1
"""

import os
import sys
import re
import json
import calendar
import argparse
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side


def load_factory_config(config_file='factory_mapping.json'):
    """
    读取厂家全称与简写映射配置文件
    """
    if not os.path.exists(config_file):
        return {
            'factory_abbreviations': {},
            'strict_vendor_suffix_drugs': [],
            'drug_code_overrides': {}
        }
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
            return {
                'factory_abbreviations': data.get('factory_abbreviations', {}),
                'strict_vendor_suffix_drugs': data.get('strict_vendor_suffix_drugs', []),
                'drug_code_overrides': data.get('drug_code_overrides', {})
            }
    except Exception as e:
        print(f"[!] 读取厂家映射文件 {config_file} 失败: {e}")
        return {
            'factory_abbreviations': {},
            'strict_vendor_suffix_drugs': [],
            'drug_code_overrides': {}
        }


def normalize_text(text):
    """文本标准化：去除所有空格，统一半角全角括号与符号"""
    if not text:
        return ''
    s = str(text).strip()
    s = re.sub(r'\s+', '', s)
    s = s.replace('（', '(').replace('）', ')')
    s = s.replace('【', '[').replace('】', ']')
    s = s.replace('×', '*').replace('x', '*').replace('X', '*')
    return s


def detect_month_end_date(sales_path, ledger_path):
    """
    自动检测目标年月并返回月末最后一天的日期字符串（YYYY-MM-DD）
    """
    # 优先从销售表文件名或总账表文件名提取，例如 "2026.8月", "202608期", "2026-08"
    combined_names = f"{os.path.basename(sales_path or '')} {os.path.basename(ledger_path or '')}"
    
    # 匹配 2026.8月 或 2026年8月
    m = re.search(r'(\d{4})[^\d]*?(\d{1,2})月', combined_names)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        last_day = calendar.monthrange(year, month)[1]
        return f"{year:04d}-{month:02d}-{last_day:02d}"

    # 匹配 202608期 或 202608
    m = re.search(r'(\d{4})(\d{2})期?', combined_names)
    if m:
        year, month = int(m.group(1)), int(m.group(2))
        if 1 <= month <= 12:
            last_day = calendar.monthrange(year, month)[1]
            return f"{year:04d}-{month:02d}-{last_day:02d}"

    # 默认当前时间所在月末
    today = date.today()
    last_day = calendar.monthrange(today.year, today.month)[1]
    return f"{today.year:04d}-{today.month:02d}-{last_day:02d}"


def load_sales_data(sales_file):
    """
    读取销售表的第二个 Sheet（销售明细）。
    如果第二个 Sheet 不存在，尝试读取第一个 Sheet 并进行聚合。
    返回: list of dict {'name': ..., 'spec': ..., 'qty': ..., 'unit': ...}
    """
    ext = os.path.splitext(sales_file)[1].lower()
    items = []

    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(sales_file)
        # 寻找销售明细 Sheet（通常是第 2 个 Sheet）
        sheet_idx = 1 if len(wb.sheet_names()) > 1 else 0
        s = wb.sheet_by_index(sheet_idx)
        sheet_name = s.name

        # 检查是否为明细表（表头通常在第 2 行或第 1 行）
        header_row = None
        for r in range(min(5, s.nrows)):
            row_vals = [str(s.cell_value(r, c)).strip() for c in range(s.ncols)]
            if '药品名称' in row_vals and '数量' in row_vals:
                header_row = r
                col_name = row_vals.index('药品名称')
                col_spec = row_vals.index('规格') if '规格' in row_vals else -1
                col_qty = row_vals.index('数量') if '数量' in row_vals else -1
                col_unit = row_vals.index('单位') if '单位' in row_vals else -1
                col_factory = row_vals.index('制药厂') if '制药厂' in row_vals else (row_vals.index('厂家') if '厂家' in row_vals else -1)
                break

        if header_row is None:
            raise ValueError(f"无法在销售表 '{sales_file}' 中找到包含【药品名称】和【数量】的表头行。")

        for r in range(header_row + 1, s.nrows):
            row_vals = [s.cell_value(r, c) for c in range(s.ncols)]
            raw_name = str(row_vals[col_name]).strip() if col_name >= 0 and col_name < len(row_vals) else ''
            if not raw_name or '合计' in raw_name:
                continue
            spec = str(row_vals[col_spec]).strip() if col_spec >= 0 and col_spec < len(row_vals) else ''
            unit = str(row_vals[col_unit]).strip() if col_unit >= 0 and col_unit < len(row_vals) else ''
            factory = str(row_vals[col_factory]).strip() if col_factory >= 0 and col_factory < len(row_vals) else ''
            try:
                raw_qty = float(row_vals[col_qty]) if col_qty >= 0 and col_qty < len(row_vals) else 0.0
                qty = int(raw_qty) if raw_qty.is_integer() else raw_qty
            except (ValueError, TypeError):
                qty = 0.0

            # 进价金额用于无总账单价时的可选回退
            col_in_amt = row_vals.index('进价金额') if '进价金额' in row_vals else -1
            try:
                in_amt = float(row_vals[col_in_amt]) if col_in_amt >= 0 and col_in_amt < len(row_vals) else 0.0
            except (ValueError, TypeError):
                in_amt = 0.0

            items.append({
                'name': raw_name,
                'spec': spec,
                'qty': qty,
                'unit': unit,
                'factory': factory,
                'in_price': round(in_amt / raw_qty, 4) if raw_qty > 0 else None
            })

    elif ext == '.xlsx':
        wb = openpyxl.load_workbook(sales_file, data_only=True)
        sheet_idx = 1 if len(wb.worksheets) > 1 else 0
        s = wb.worksheets[sheet_idx]
        sheet_name = s.title

        header_row = None
        for r in range(1, min(6, s.max_row + 1)):
            row_vals = [str(s.cell(r, c).value or '').strip() for c in range(1, s.max_column + 1)]
            if '药品名称' in row_vals and '数量' in row_vals:
                header_row = r
                col_name = row_vals.index('药品名称') + 1
                col_spec = row_vals.index('规格') + 1 if '规格' in row_vals else -1
                col_qty = row_vals.index('数量') + 1 if '数量' in row_vals else -1
                col_unit = row_vals.index('单位') + 1 if '单位' in row_vals else -1
                col_factory = row_vals.index('制药厂') + 1 if '制药厂' in row_vals else (row_vals.index('厂家') + 1 if '厂家' in row_vals else -1)
                break

        if header_row is None:
            raise ValueError(f"无法在销售表 '{sales_file}' 中找到包含【药品名称】和【数量】的表头行。")

        col_in_amt = -1
        if '进价金额' in [str(s.cell(header_row, c).value or '').strip() for c in range(1, s.max_column + 1)]:
            for c in range(1, s.max_column + 1):
                if str(s.cell(header_row, c).value or '').strip() == '进价金额':
                    col_in_amt = c
                    break

        for r in range(header_row + 1, s.max_row + 1):
            raw_name = str(s.cell(r, col_name).value or '').strip() if col_name > 0 else ''
            if not raw_name or '合计' in raw_name:
                continue
            spec = str(s.cell(r, col_spec).value or '').strip() if col_spec > 0 else ''
            unit = str(s.cell(r, col_unit).value or '').strip() if col_unit > 0 else ''
            factory = str(s.cell(r, col_factory).value or '').strip() if col_factory > 0 else ''
            try:
                raw_qty = float(s.cell(r, col_qty).value or 0.0) if col_qty > 0 else 0.0
                qty = int(raw_qty) if raw_qty.is_integer() else raw_qty
            except (ValueError, TypeError):
                qty = 0.0

            try:
                in_amt = float(s.cell(r, col_in_amt).value or 0.0) if col_in_amt > 0 else 0.0
            except (ValueError, TypeError):
                in_amt = 0.0

            items.append({
                'name': raw_name,
                'spec': spec,
                'qty': qty,
                'unit': unit,
                'factory': factory,
                'in_price': round(in_amt / raw_qty, 4) if raw_qty > 0 else None
            })
    else:
        raise ValueError(f"不支持的销售表文件格式: {ext}")

    return sheet_name, items


def load_ledger_and_aux_data(ledger_file, template_file):
    """
    读取总账表及凭证模板中的辅助核算项目数据，构建匹配索引
    返回: (ledger_exact, ledger_norm, aux_codes_set)
    """
    wb_ledger = openpyxl.load_workbook(ledger_file, data_only=True)
    s_ledger = wb_ledger.active

    # 寻找总账表的数据列位置
    # 通常 Col 1 是科目编码 (1201_XY0001)，Col 2 是科目名称 (存货_肉蔻五味丸 0.2g*100粒/盒)
    # 单价在期初/期末单价列（如 Col 6 / Col 18）
    ledger_exact = {}
    ledger_norm = {}

    for r in range(4, s_ledger.max_row + 1):
        c1 = str(s_ledger.cell(r, 1).value or '').strip()
        c2 = str(s_ledger.cell(r, 2).value or '').strip()
        if not c2:
            continue

        # 提取辅助核算编码 (例如 1201_XY0001 -> XY0001)
        aux_code = c1.split('_')[-1] if '_' in c1 else c1

        # 提取单价 (优先取第 18 列【期末结存单价】，其次取第 6 列【期初单价】)
        def parse_price(val):
            try:
                p = float(val) if val is not None and str(val).strip() != '' else 0.0
                return p if p > 0.0 else None
            except (ValueError, TypeError):
                return None

        price = parse_price(s_ledger.cell(r, 18).value)
        if price is None:
            price = parse_price(s_ledger.cell(r, 6).value)

        item_info = {
            'subject_code': c1,
            'aux_code': aux_code,
            'subject_name': c2,
            'price': price
        }

        ledger_exact[c2] = item_info
        ledger_norm[normalize_text(c2)] = item_info

    # 读取凭证导入模板中的【辅助核算项目数据】(Sheet 3)
    aux_codes_set = set()
    aux_name_map = {}
    if os.path.exists(template_file):
        wb_tmpl = openpyxl.load_workbook(template_file, data_only=True)
        if '辅助核算项目数据' in wb_tmpl.sheetnames:
            s_aux = wb_tmpl['辅助核算项目数据']
            for r in range(3, s_aux.max_row + 1):
                code = str(s_aux.cell(r, 2).value or '').strip()
                name = str(s_aux.cell(r, 3).value or '').strip()
                if code:
                    aux_codes_set.add(code)
                if name:
                    aux_name_map[normalize_text(name)] = code

    return ledger_exact, ledger_norm, aux_codes_set, aux_name_map


def match_drug_item(item, ledger_exact, ledger_norm, aux_name_map, factory_config=None):
    """
    为销售药品匹配总账科目与辅助核算编码
    规则：优先按 厂家/供应商 简称进行带括号标识匹配（如 鸡血藤（蕴德）、海螵蛸（蕴德））
    若属于严格厂家限定品种（如 海螵蛸），不回退到无厂家的老科目；
    非严格品种且无其他厂家冲突时，回退到标准科目匹配。
    """
    name = item['name']
    spec = item['spec']
    factory = item.get('factory') or item.get('supplier') or ''

    if factory_config is None:
        factory_config = {}
    factory_abbrs = factory_config.get('factory_abbreviations', {})
    strict_drugs = factory_config.get('strict_vendor_suffix_drugs', [])
    overrides = factory_config.get('drug_code_overrides', {})

    abbr = None
    if factory:
        norm_fac = normalize_text(factory)
        for full_f, short_f in factory_abbrs.items():
            norm_k = normalize_text(full_f)
            if norm_k == norm_fac or norm_k in norm_fac or norm_fac in norm_k:
                abbr = short_f
                break

    all_vendor_tags = set(factory_abbrs.values()) | {'国瑞堂', '国松堂', '盛方', '神农', '蕴德'}
    other_vendor_tags = {v for v in all_vendor_tags if v != abbr}

    # 2. 如果存在厂家简写，构造带厂家后缀的目标药名优先匹配
    if abbr:
        target_name_with_vendor = f"{name}（{abbr}）"
        target_name_half = f"{name}({abbr})"

        for cand_name in [target_name_with_vendor, target_name_half, name]:
            if cand_name in overrides and overrides[cand_name]:
                return {
                    'subject_code': '',
                    'aux_code': overrides[cand_name],
                    'subject_name': target_name_with_vendor,
                    'price': None
                }

        for cand_vendor in [target_name_with_vendor, target_name_half]:
            c1 = f"存货_{cand_vendor} {spec}".strip()
            if c1 in ledger_exact:
                return ledger_exact[c1]
            c2 = f"存货_{cand_vendor}".strip()
            if c2 in ledger_exact:
                return ledger_exact[c2]
            
            n1 = normalize_text(f"存货_{cand_vendor}{spec}")
            if n1 in ledger_norm:
                return ledger_norm[n1]
            n2 = normalize_text(f"存货_{cand_vendor}")
            if n2 in ledger_norm:
                return ledger_norm[n2]

        for cand_vendor in [target_name_with_vendor, target_name_half]:
            norm_tv = normalize_text(cand_vendor)
            if norm_tv in aux_name_map:
                return {
                    'subject_code': '',
                    'aux_code': aux_name_map[norm_tv],
                    'subject_name': target_name_with_vendor,
                    'price': None
                }

        if name in strict_drugs:
            item['target_name'] = target_name_with_vendor
            return None

        norm_name = normalize_text(name)
        has_other_vendor_conflict = False
        for k_norm in ledger_norm.keys():
            if norm_name in k_norm:
                if any(normalize_text(other_v) in k_norm for other_v in other_vendor_tags):
                    has_other_vendor_conflict = True
                    break
        if has_other_vendor_conflict:
            item['target_name'] = target_name_with_vendor
            return None

    # 3. 常规通用名匹配
    cand1 = f"存货_{name} {spec}".strip()
    if cand1 in ledger_exact:
        return ledger_exact[cand1]

    cand2 = f"存货_{name}".strip()
    if cand2 in ledger_exact:
        return ledger_exact[cand2]

    norm1 = normalize_text(f"存货_{name}{spec}")
    if norm1 in ledger_norm:
        return ledger_norm[norm1]

    norm2 = normalize_text(f"存货_{name}")
    if norm2 in ledger_norm:
        return ledger_norm[norm2]

    norm_name = normalize_text(name)
    norm_spec = normalize_text(spec)
    for k_norm, val in ledger_norm.items():
        if norm_name in k_norm:
            if any(normalize_text(other_v) in k_norm for other_v in other_vendor_tags):
                continue
            if not norm_spec or norm_spec in k_norm:
                return val

    if norm_name in aux_name_map:
        return {
            'subject_code': '',
            'aux_code': aux_name_map[norm_name],
            'subject_name': name,
            'price': None
        }

    return None


def generate_voucher_file(sales_file, ledger_file, template_file, output_file=None, voucher_date=None, fallback_price=False, config_file='factory_mapping.json'):
    """
    主生成函数：生成凭证导入模板
    """
    # 0. 读取厂家映射配置
    factory_config = load_factory_config(config_file)
    if factory_config.get('factory_abbreviations'):
        print(f"[*] 成功加载厂家全名与简写映射表: 共 {len(factory_config['factory_abbreviations'])} 个映射项")

    # 1. 自动计算或使用指定凭证日期
    target_date = voucher_date or detect_month_end_date(sales_file, ledger_file)

    # 2. 读取销售明细
    sheet_name, sales_items = load_sales_data(sales_file)
    print(f"[*] 成功读取销售明细: 来自【{os.path.basename(sales_file)}】的【{sheet_name}】")
    print(f"[*] 共读取到 {len(sales_items)} 条药品明细记录")

    # 3. 读取总账与辅助核算字典
    ledger_exact, ledger_norm, aux_codes_set, aux_name_map = load_ledger_and_aux_data(ledger_file, template_file)
    print(f"[*] 成功读取总账科目: 共 {len(ledger_exact)} 个科目项目")

    # 4. 加载凭证导入模板
    wb_out = openpyxl.load_workbook(template_file)
    if '凭证模版' not in wb_out.sheetnames:
        raise ValueError(f"凭证导入模板中未找到【凭证模版】工作表！")
    s_voucher = wb_out['凭证模版']

    # 清空原有数据行（保留第 1 行表头）
    if s_voucher.max_row > 1:
        s_voucher.delete_rows(2, s_voucher.max_row)

    # 样式定义
    thin = Side(border_style="thin", color="000000")
    border_all = Border(top=thin, left=thin, right=thin, bottom=thin)
    font_body = Font(name='宋体', size=10, bold=False)
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')

    # 5. 插入第 1 条分录：空行用于填写借方金额（分录序号为 1）
    debit_row_values = [
        target_date,      # 1. 日期: 月末最后一天
        '记',             # 2. 凭证字: 固定为“记”
        None,             # 3. 凭证号: 留空
        None,             # 4. 附件数: 留空
        1,                # 5. 分录序号: 1
        None,             # 6. 摘要: 留空待填
        None,             # 7. 科目代码: 留空待填
        None,             # 8. 科目名称: 留空待填
        None,             # 9. 借方金额: 留空待填
        None,             # 10. 贷方金额: 留空
        None,             # 11. 客户
        None,             # 12. 供应商
        None,             # 13. 职员
        None,             # 14. 项目
        None,             # 15. 部门
        None,             # 16. 存货
        None,             # 17. 是否限定
        None,             # 18. 自定义类别
        None,             # 19. 自定义编码
        None,             # 20. 自定义类别1
        None,             # 21. 自定义编码1
        None,             # 22. 数量
        None,             # 23. 单价
        None,             # 24. 原币金额
        'RMB',            # 25. 币别: RMB
        1,                # 26. 汇率: 1
    ]

    s_voucher.row_dimensions[2].height = 19
    for col_idx, val in enumerate(debit_row_values, start=1):
        cell = s_voucher.cell(row=2, column=col_idx, value=val)
        cell.font = font_body
        cell.border = border_all
        if col_idx in [9, 10, 24]: # 金额列（借方金额、贷方金额、原币金额）
            cell.alignment = align_right
            cell.number_format = '#,##0.00'
        elif col_idx in [22, 23]:
            cell.alignment = align_right
    # 6. 逐行匹配并填充贷方明细分录（从分录序号 2 开始递增）
    matched_count = 0
    unmatched_items = []
    total_credit_amt = 0.0
    total_qty = 0.0

    current_row = 3
    for entry_seq, item in enumerate(sales_items, start=2):
        matched = match_drug_item(item, ledger_exact, ledger_norm, aux_name_map, factory_config)

        qty = item['qty']
        total_qty += qty

        if matched and matched.get('aux_code'):
            aux_code = matched['aux_code']
            price = matched.get('price')
            matched_count += 1
        else:
            aux_code = ''
            price = item.get('in_price') if fallback_price else None
            unmatched_items.append(item)

        if price is not None:
            amt = round(qty * price, 2)
            total_credit_amt += amt
            amt_val = amt
            price_val = price
        else:
            amt_val = None
            price_val = None

        # 构造各列值 (共 26 列)
        row_values = [
            target_date,      # 1. 日期: 月末最后一天
            '记',             # 2. 凭证字: 固定为“记”
            None,             # 3. 凭证号: 为空
            None,             # 4. 附件数: 为空
            entry_seq,        # 5. 分录序号: 从 2 递增
            None,             # 6. 摘要: 为空
            None,             # 7. 科目代码: 为空
            None,             # 8. 科目名称: 为空
            None,             # 9. 借方金额: 为空
            amt_val,          # 10. 贷方金额: 数量 * 单价
            None,             # 11. 客户
            None,             # 12. 供应商
            None,             # 13. 职员
            None,             # 14. 项目
            None,             # 15. 部门
            aux_code,         # 16. 存货: 辅助核算编码
            None,             # 17. 是否限定
            None,             # 18. 自定义类别
            None,             # 19. 自定义编码
            None,             # 20. 自定义类别1
            None,             # 21. 自定义编码1
            qty,              # 22. 数量: 销售数量
            price_val,        # 23. 单价: 总账加权单价
            amt_val,          # 24. 原币金额: 销售贷方金额
            'RMB',            # 25. 币别: 固定为 RMB
            1,                # 26. 汇率: 固定为 1
        ]

        s_voucher.row_dimensions[current_row].height = 19
        for col_idx, val in enumerate(row_values, start=1):
            cell = s_voucher.cell(row=current_row, column=col_idx, value=val)
            cell.font = font_body
            cell.border = border_all
            if col_idx in [9, 10, 24]:
                cell.alignment = align_right
                cell.number_format = '#,##0.00'
            elif col_idx in [22, 23]:
                cell.alignment = align_right
            else:
                cell.alignment = align_center

        current_row += 1

    # 保存生成的文件
    save_path = output_file or '凭证导入模板_已生成.xlsx'
    wb_out.save(save_path)

    # 7. 打印总结报表
    print("\n" + "="*50)
    print("           凭证导入模板生成完成报告")
    print("="*50)
    print(f"凭证入账日期: {target_date}")
    print(f"生成分录总数: {len(sales_items) + 1} 行（借方 1 行 + 贷方明细 {len(sales_items)} 行）")
    print(f"匹配成功总数: {matched_count} 行")
    print(f"未匹配/需关注: {len(unmatched_items)} 行")
    print(f"销售数量总和: {total_qty:,.2f}")
    print(f"贷方金额总计: {total_credit_amt:,.2f} 元")
    print(f"输出保存文件: {save_path}")

    if unmatched_items:
        print("\n[⚠️ 注意] 以下药品在总账中未匹配到科目或编码，存货编码/单价暂已留空，请在系统中补录核实：")
        for u in unmatched_items:
            show_name = u.get('target_name') or u['name']
            print(f"  - 药品: {show_name} | 规格: {u['spec']} | 销售数量: {u['qty']}")

    return save_path


def main():
    parser = argparse.ArgumentParser(description='根据西药/中药销售明细与总账表自动生成凭证导入模板')
    parser.add_argument('-s', '--sales', default=None, help='销售表文件路径（默认自动查找西药销售表.xls/.xlsx）')
    parser.add_argument('-l', '--ledger', default=None, help='总账表文件路径（默认自动查找数量金额总账*.xlsx）')
    parser.add_argument('-t', '--template', default='凭证导入模板.xlsx', help='凭证导入模板路径（默认当前目录凭证导入模板.xlsx）')
    parser.add_argument('-o', '--output', default=None, help='输出凭证文件路径（默认 凭证导入模板_已生成.xlsx）')
    parser.add_argument('-d', '--date', default=None, help='指定凭证日期（格式 YYYY-MM-DD，默认自动推导月末最后一天）')
    parser.add_argument('--fallback-price', action='store_true', help='未匹配到总账单价时，是否自动回退使用销售表进价单价')
    parser.add_argument('-c', '--config', default='factory_mapping.json', help='厂家全称与简写映射配置文件（默认 factory_mapping.json）')

    args = parser.parse_args()

    # 自动搜索销售表
    sales_file = args.sales
    if not sales_file:
        cands = [f for f in os.listdir('.') if f.endswith(('.xls', '.xlsx')) and '销售' in f and not f.startswith('~$') and 'backup' not in f]
        if cands:
            sales_file = cands[0]
        else:
            print("错误: 未找到销售表文件，请使用 -s 指定销售表路径。")
            sys.exit(1)

    # 自动搜索总账表
    ledger_file = args.ledger
    if not ledger_file:
        cands = [f for f in os.listdir('.') if f.endswith('.xlsx') and '总账' in f and not f.startswith('~$')]
        if cands:
            ledger_file = cands[0]
        else:
            print("错误: 未找到数量金额总账表，请使用 -l 指定总账表路径。")
            sys.exit(1)

    # 检查模板文件
    template_file = args.template
    if not os.path.exists(template_file):
        print(f"错误: 凭证导入模板 '{template_file}' 不存在。")
        sys.exit(1)

    print(f"使用的销售表: {sales_file}")
    print(f"使用的总账表: {ledger_file}")
    print(f"使用的模板表: {template_file}")
    print(f"使用的映射表: {args.config}")

    generate_voucher_file(
        sales_file=sales_file,
        ledger_file=ledger_file,
        template_file=template_file,
        output_file=args.output,
        voucher_date=args.date,
        fallback_price=args.fallback_price,
        config_file=args.config
    )


if __name__ == '__main__':
    main()
