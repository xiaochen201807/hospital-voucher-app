#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Tauri 桌面端自动化统一调度引擎 (CLI / JSON API)
为桌面 UI 提供四大核心业务支持：
1. 销售明细去重汇总 (process_sales)
2. 销售出库凭证生成 (generate_voucher)
3. 药品入库凭证生成 (generate_inbound_voucher)
4. 厂家与编码映射字典管理 (factory_mapping)
5. 本地目录候选文件自动扫描 (scan)
"""

import os
import sys
import json
import argparse
import glob

# 智能判定运行根目录（兼容 PyInstaller 独立二进制与开发调试）
IS_FROZEN = getattr(sys, 'frozen', False)
BASE_DIR = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
PARENT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))

for d in [BASE_DIR, PARENT_DIR, os.path.abspath(os.path.dirname(__file__)), os.getcwd()]:
    if d and os.path.exists(d) and d not in sys.path:
        sys.path.insert(0, d)

# 显式导入 Excel 处理库
try:
    import openpyxl
    import xlrd
    import xlwt
    import xlutils
    import xlutils.copy
except ImportError:
    pass

# 显式导入核心业务模块（逐一导入，避免连锁失效）
try:
    import process_sales
except Exception as e:
    print(f"[Warning] 导入 process_sales 异常: {e}", file=sys.stderr)

try:
    import generate_voucher
except Exception as e:
    print(f"[Warning] 导入 generate_voucher 异常: {e}", file=sys.stderr)

try:
    import generate_inbound_voucher
except Exception as e:
    print(f"[Warning] 导入 generate_inbound_voucher 异常: {e}", file=sys.stderr)


def cmd_scan(args):
    """扫描指定目录或工作区中的候选 Excel 文件"""
    scan_dir = args.dir or PARENT_DIR
    res = {
        'sales_files': [],
        'inbound_files': [],
        'ledger_files': [],
        'template_files': [],
        'all_excel': []
    }
    if not os.path.exists(scan_dir):
        return {'success': False, 'error': f"目录 '{scan_dir}' 不存在"}

    for f in sorted(os.listdir(scan_dir)):
        if f.startswith('~$') or not f.endswith(('.xlsx', '.xls')):
            continue
        full_path = os.path.join(scan_dir, f)
        if '已生成' in f or 'backup' in f:
            continue
        
        res['all_excel'].append({'name': f, 'path': full_path})
        if '销售' in f:
            res['sales_files'].append({'name': f, 'path': full_path})
        if '入库' in f and '模板' not in f:
            res['inbound_files'].append({'name': f, 'path': full_path})
        if '总账' in f:
            res['ledger_files'].append({'name': f, 'path': full_path})
        if '模板' in f:
            res['template_files'].append({'name': f, 'path': full_path})

    return {'success': True, 'data': res}


def find_config_path(custom_path=None):
    """智能查找 factory_mapping.json 配置文件路径"""
    if custom_path and os.path.exists(custom_path):
        return custom_path
    for p in [
        os.path.join(os.getcwd(), 'factory_mapping.json'),
        os.path.join(BASE_DIR, 'factory_mapping.json'),
        os.path.join(PARENT_DIR, 'factory_mapping.json'),
    ]:
        if os.path.exists(p):
            return os.path.abspath(p)
    return os.path.abspath(os.path.join(BASE_DIR, 'factory_mapping.json'))


def cmd_process_sales(args):
    """执行销售汇总去重处理"""
    if 'process_sales' not in globals() or process_sales is None:
        return {'success': False, 'error': "核心模块 process_sales 加载失败，请检查安装包完整性"}

    sales_file = args.file
    if not sales_file or not os.path.exists(sales_file):
        return {'success': False, 'error': f"销售表文件 '{sales_file}' 不存在"}

    ext = os.path.splitext(sales_file)[1].lower()
    output_path = args.output
    sheet_name = args.sheet_name or '销售明细'

    try:
        if ext == '.xls':
            target_sheet_name, totals, final_out = process_sales.process_xls_file(sales_file, output_path, sheet_name)
        elif ext == '.xlsx':
            target_sheet_name, totals, final_out = process_sales.process_xlsx_file(sales_file, output_path, sheet_name)
        else:
            return {'success': False, 'error': f"不支持的文件类型: {ext}"}

        return {
            'success': True,
            'output_file': os.path.abspath(final_out),
            'sheet_name': target_sheet_name,
            'totals': {
                'original_count': totals['original_count'],
                'unique_count': totals['unique_count'],
                'total_qty': totals['total_qty'],
                'total_in_amt': totals['total_in_amt'],
                'total_retail_amt': totals['total_retail_amt']
            }
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}


def cmd_generate_voucher(args):
    """执行销售出库凭证生成"""
    if 'generate_voucher' not in globals() or generate_voucher is None:
        return {'success': False, 'error': "核心模块 generate_voucher 加载失败，请检查安装包完整性"}

    sales_file = args.sales
    ledger_file = args.ledger
    template_file = args.template
    config_file = find_config_path(args.config)

    if not sales_file or not os.path.exists(sales_file):
        return {'success': False, 'error': f"销售明细表 '{sales_file}' 不存在"}
    if not ledger_file or not os.path.exists(ledger_file):
        return {'success': False, 'error': f"总账表 '{ledger_file}' 不存在"}
    if not template_file or not os.path.exists(template_file):
        return {'success': False, 'error': f"凭证模板 '{template_file}' 不存在"}

    output_file = args.output or os.path.join(os.path.dirname(sales_file), '凭证导入模板_已生成.xlsx')

    try:
        # 加载配置
        factory_config = generate_voucher.load_factory_config(config_file)
        
        # 计算日期与明细
        target_date = args.date or generate_voucher.detect_month_end_date(sales_file, ledger_file)
        sheet_name, sales_items = generate_voucher.load_sales_data(sales_file)
        ledger_exact, ledger_norm, aux_codes_set, aux_name_map = generate_voucher.load_ledger_and_aux_data(ledger_file, template_file)

        # 执行凭证文件生成
        final_path = generate_voucher.generate_voucher_file(
            sales_file=sales_file,
            ledger_file=ledger_file,
            template_file=template_file,
            output_file=output_file,
            voucher_date=target_date,
            fallback_price=args.fallback_price,
            config_file=config_file
        )

        # 统计匹配状态
        matched_count = 0
        unmatched_items = []
        total_credit_amt = 0.0
        total_qty = 0.0

        for item in sales_items:
            matched = generate_voucher.match_drug_item(item, ledger_exact, ledger_norm, aux_name_map, factory_config)
            qty = item['qty']
            total_qty += qty
            if matched and matched.get('aux_code'):
                price = matched.get('price')
                matched_count += 1
            else:
                price = item.get('in_price') if args.fallback_price else None
                unmatched_items.append({
                    'name': item['name'],
                    'target_name': item.get('target_name') or item['name'],
                    'spec': item.get('spec', ''),
                    'factory': item.get('factory', ''),
                    'qty': item.get('qty', 0),
                    'in_price': item.get('in_price')
                })
            if price is not None:
                total_credit_amt += round(qty * price, 2)

        return {
            'success': True,
            'output_file': os.path.abspath(final_path),
            'voucher_date': target_date,
            'total_items': len(sales_items),
            'matched_count': matched_count,
            'unmatched_count': len(unmatched_items),
            'match_rate': round(matched_count / len(sales_items) * 100, 1) if sales_items else 0,
            'total_qty': round(total_qty, 2),
            'total_credit_amt': round(total_credit_amt, 2),
            'unmatched_items': unmatched_items
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}


def cmd_generate_inbound_voucher(args):
    """执行药品入库凭证生成 (西药/中药)"""
    if 'generate_inbound_voucher' not in globals() or generate_inbound_voucher is None:
        return {'success': False, 'error': "核心模块 generate_inbound_voucher 加载失败，请检查安装包完整性"}

    inbound_file = args.inbound
    ledger_file = args.ledger
    template_file = args.template
    config_file = find_config_path(args.config)

    if not inbound_file or not os.path.exists(inbound_file):
        return {'success': False, 'error': f"入库单文件 '{inbound_file}' 不存在"}
    if not ledger_file or not os.path.exists(ledger_file):
        return {'success': False, 'error': f"总账表 '{ledger_file}' 不存在"}
    if not template_file or not os.path.exists(template_file):
        return {'success': False, 'error': f"凭证模板 '{template_file}' 不存在"}

    output_file = args.output
    if not output_file:
        base_dir = os.path.dirname(inbound_file) or '.'
        is_zy = '中药' in os.path.basename(inbound_file)
        out_name = '凭证导入模板_中药入库_已生成.xlsx' if is_zy else '凭证导入模板_入库_已生成.xlsx'
        output_file = os.path.join(base_dir, out_name)

    try:
        factory_config = generate_inbound_voucher.load_factory_config(config_file)
        sheet_name, inbound_items, max_date_str = generate_inbound_voucher.load_inbound_data(inbound_file)
        target_date = args.date or generate_inbound_voucher.detect_month_end_date(inbound_file, ledger_file, max_date_str)
        ledger_exact, ledger_norm, aux_codes_set, aux_name_map, supplier_dict = generate_inbound_voucher.load_ledger_and_aux_data(ledger_file, template_file)

        final_path = generate_inbound_voucher.generate_inbound_voucher_file(
            inbound_file=inbound_file,
            ledger_file=ledger_file,
            template_file=template_file,
            output_file=output_file,
            voucher_date=target_date,
            voucher_no=args.voucher_no,
            config_file=config_file
        )

        # 统计供应商分组与借贷总额
        from collections import OrderedDict
        grouped = OrderedDict()
        matched_count = 0
        unmatched_items = []
        total_debit_amt = 0.0
        total_credit_amt = 0.0
        total_qty = 0.0

        for item in inbound_items:
            sup = item['supplier'] or '未知供应商'
            if sup not in grouped:
                grouped[sup] = []
            grouped[sup].append(item)

            matched = generate_inbound_voucher.match_drug_item(item, ledger_exact, ledger_norm, aux_name_map, factory_config)
            qty = item['qty']
            amt = round(item['in_amt'], 2)
            total_qty += qty
            total_debit_amt += amt

            if matched and matched.get('aux_code'):
                matched_count += 1
            else:
                unmatched_items.append({
                    'supplier': sup,
                    'name': item['name'],
                    'target_name': item.get('target_name') or item['name'],
                    'spec': item['spec'],
                    'qty': item['qty'],
                    'in_price': item['in_price'],
                    'in_amt': amt
                })

        suppliers_summary = []
        for sup_name, items in grouped.items():
            sup_code, _ = generate_inbound_voucher.match_supplier_code(sup_name, supplier_dict)
            brief = generate_inbound_voucher.get_supplier_brief(sup_name)
            sub_total = round(sum(it['in_amt'] for it in items), 2)
            total_credit_amt += sub_total
            suppliers_summary.append({
                'supplier': sup_name,
                'code': sup_code,
                'brief': brief,
                'count': len(items),
                'amount': sub_total
            })

        is_balanced = round(total_debit_amt, 2) == round(total_credit_amt, 2)

        return {
            'success': True,
            'output_file': os.path.abspath(final_path),
            'voucher_date': target_date,
            'total_items': len(inbound_items),
            'total_entries': len(inbound_items) + len(grouped),
            'matched_count': matched_count,
            'unmatched_count': len(unmatched_items),
            'match_rate': round(matched_count / len(inbound_items) * 100, 1) if inbound_items else 0,
            'total_qty': round(total_qty, 2),
            'total_debit_amt': round(total_debit_amt, 2),
            'total_credit_amt': round(total_credit_amt, 2),
            'is_balanced': is_balanced,
            'suppliers_summary': suppliers_summary,
            'unmatched_items': unmatched_items
        }
    except Exception as e:
        return {'success': False, 'error': str(e)}


def cmd_get_config(args):
    """读取 factory_mapping.json"""
    config_file = args.config or os.path.join(PARENT_DIR, 'factory_mapping.json')
    if not os.path.exists(config_file):
        return {'success': False, 'error': f"配置文件 '{config_file}' 不存在"}
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            data = json.load(f)
        return {'success': True, 'data': data, 'config_file': os.path.abspath(config_file)}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def cmd_save_config(args):
    """保存 factory_mapping.json"""
    config_file = args.config or os.path.join(PARENT_DIR, 'factory_mapping.json')
    try:
        new_data = json.loads(args.data)
        with open(config_file, 'w', encoding='utf-8') as f:
            json.dump(new_data, f, ensure_ascii=False, indent=2)
        return {'success': True, 'message': '配置保存成功', 'config_file': os.path.abspath(config_file)}
    except Exception as e:
        return {'success': False, 'error': str(e)}


def main():
    parser = argparse.ArgumentParser(description='Tauri 桌面端自动化统一调度引擎')
    subparsers = parser.add_subparsers(dest='action', help='子命令动作')

    # scan
    p_scan = subparsers.add_parser('scan')
    p_scan.add_argument('--dir', default=None)

    # process_sales
    p_sales = subparsers.add_parser('process_sales')
    p_sales.add_argument('--file', required=True)
    p_sales.add_argument('--output', default=None)
    p_sales.add_argument('--sheet-name', default=None)

    # generate_voucher (outbound)
    p_out = subparsers.add_parser('generate_voucher')
    p_out.add_argument('--sales', required=True)
    p_out.add_argument('--ledger', required=True)
    p_out.add_argument('--template', required=True)
    p_out.add_argument('--output', default=None)
    p_out.add_argument('--date', default=None)
    p_out.add_argument('--fallback-price', action='store_true')
    p_out.add_argument('--config', default=None)

    # generate_inbound_voucher (inbound)
    p_in = subparsers.add_parser('generate_inbound_voucher')
    p_in.add_argument('--inbound', required=True)
    p_in.add_argument('--ledger', required=True)
    p_in.add_argument('--template', required=True)
    p_in.add_argument('--output', default=None)
    p_in.add_argument('--date', default=None)
    p_in.add_argument('--voucher-no', default=None)
    p_in.add_argument('--config', default=None)

    # get_config
    p_get_cfg = subparsers.add_parser('get_config')
    p_get_cfg.add_argument('--config', default=None)

    # save_config
    p_save_cfg = subparsers.add_parser('save_config')
    p_save_cfg.add_argument('--data', required=True)
    p_save_cfg.add_argument('--config', default=None)

    args = parser.parse_args()

    dispatch = {
        'scan': cmd_scan,
        'process_sales': cmd_process_sales,
        'generate_voucher': cmd_generate_voucher,
        'generate_inbound_voucher': cmd_generate_inbound_voucher,
        'get_config': cmd_get_config,
        'save_config': cmd_save_config,
    }

    if args.action in dispatch:
        res = dispatch[args.action](args)
        # 始终以单行纯 JSON 输出，供 Rust / Tauri 安全无误解析
        print(json.dumps(res, ensure_ascii=False))
    else:
        print(json.dumps({'success': False, 'error': f"未知命令动作: {args.action}"}, ensure_ascii=False))


if __name__ == '__main__':
    main()
