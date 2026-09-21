use crate::core::config::{clean_text, ConfigData};
use crate::core::excel_utils::{
    amount_or_product, cents_to_currency, collect_source_dates, currency_cents, find_header_row,
    is_summary_row, optional_col, read_sheet_rows, required_col, row_as_f64, row_as_string,
};
use crate::core::ledger::load_ledger_entries;
use crate::core::matching::{build_ledger_candidates, resolve_drug_match};
use crate::core::models::{ConfirmedLedgerMapping, UnmatchedDrug};
use crate::core::voucher_writer::{
    copy_template_sheets, load_template_sheets, set_default_voucher_column_widths,
    write_voucher_headers, write_voucher_line, VoucherFormats, VoucherLine,
};
use rust_xlsxwriter::Workbook;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::{Path, PathBuf};

pub use crate::core::supplier::{get_supplier_brief, load_supplier_dict, match_supplier_code};

#[derive(Debug, Serialize, Deserialize)]
pub struct SupplierSummary {
    pub supplier: String,
    pub code: String,
    pub brief: String,
    pub count: usize,
    pub amount: f64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct InboundVoucherResult {
    pub success: bool,
    pub output_file: String,
    pub voucher_date: String,
    pub total_items: usize,
    pub total_entries: usize,
    pub matched_count: usize,
    pub unmatched_count: usize,
    pub match_rate: f64,
    pub total_qty: f64,
    pub total_debit_amt: f64,
    pub total_credit_amt: f64,
    pub is_balanced: bool,
    pub suppliers_summary: Vec<SupplierSummary>,
    pub unmatched_items: Vec<UnmatchedDrug>,
    #[serde(default)]
    pub error: Option<String>,
}

#[derive(Debug, Clone)]
struct InboundItem {
    supplier: String,
    name: String,
    spec: String,
    factory: String,
    qty: f64,
    price: f64,
    amount: f64,
    amount_cents: i64,
    #[allow(dead_code)]
    unit: String,
}

#[derive(Debug, Clone)]
struct SupplierGroup {
    supplier: String,
    code: String,
    brief: String,
    item_indices: Vec<usize>,
    amount_cents: i64,
}

/// 按供应商编码建立分组；未匹配到编码时才退回清洗后的供应商名称。
///
/// 入库单中的供应商名称可能同时出现简称、全称或系统截断值。只用原始字符串
/// 分组会把同一供应商拆成多张汇总分录，先解析标准编码再分组才能保证“一供应商一汇总”。
fn build_supplier_groups(
    items: &[InboundItem],
    supplier_dict: &HashMap<String, String>,
) -> Vec<SupplierGroup> {
    let mut groups = Vec::new();
    let mut group_indices: HashMap<String, usize> = HashMap::new();

    for (item_idx, item) in items.iter().enumerate() {
        let (code, matched_supplier) = match_supplier_code(&item.supplier, supplier_dict);
        let group_key = if !code.is_empty() {
            format!("code:{}", code)
        } else {
            format!("name:{}", clean_text(&item.supplier))
        };
        let display_supplier = if !matched_supplier.trim().is_empty() {
            matched_supplier
        } else {
            item.supplier.clone()
        };

        let group_idx = if let Some(group_idx) = group_indices.get(&group_key) {
            *group_idx
        } else {
            let group_idx = groups.len();
            groups.push(SupplierGroup {
                supplier: display_supplier.clone(),
                code: code.clone(),
                brief: get_supplier_brief(&display_supplier),
                item_indices: Vec::new(),
                amount_cents: 0,
            });
            group_indices.insert(group_key, group_idx);
            group_idx
        };

        let group = &mut groups[group_idx];
        group.item_indices.push(item_idx);
        group.amount_cents += item.amount_cents;
    }

    groups
}

/// 执行生成药房入库凭证
#[allow(clippy::too_many_arguments)]
pub fn generate_inbound_voucher(
    inbound_path: &str,
    ledger_path: &str,
    template_path: &str,
    custom_output: Option<&str>,
    target_date_opt: Option<&str>,
    voucher_no: Option<&str>,
    config: &ConfigData,
    confirmed_mappings: Option<&[ConfirmedLedgerMapping]>,
) -> Result<InboundVoucherResult, String> {
    let in_p = Path::new(inbound_path);
    let ledger_p = Path::new(ledger_path);
    let tmpl_p = Path::new(template_path);

    if !in_p.exists() {
        return Err(format!("入库单文件 '{:?}' 不存在", in_p));
    }
    if !ledger_p.exists() {
        return Err(format!("总账文件 '{:?}' 不存在", ledger_p));
    }
    if !tmpl_p.exists() {
        return Err(format!("凭证模板文件 '{:?}' 不存在", tmpl_p));
    }

    // 1. 读取总账建立字典
    let ledger_entries = load_ledger_entries(ledger_p)?;

    // 2. 只读取一次模板，同时得到供应商字典和附表快照。
    let tmpl_sheets = load_template_sheets(tmpl_p)?;
    let supplier_dict = load_supplier_dict(&tmpl_sheets);

    // 3. 读取入库单明细
    let (_, rows) = read_sheet_rows(in_p, &["入库", "明细"], "入库单")?;
    if rows.len() < 2 {
        return Err("入库单中无有效数据".into());
    }

    let h_idx = find_header_row(&rows, &["药品名称", "品名", "商品名称"], 10, "入库单")?;
    let header = &rows[h_idx];
    let col_name = required_col(header, &["药品名称", "品名", "商品名称"], "药品名称")?;
    let col_spec = optional_col(header, &["规格"], col_name + 1);
    let col_factory = optional_col(header, &["制药厂", "生产厂家", "厂家"], col_spec + 1);
    let col_sup = optional_col(header, &["供货单位", "供应商", "单位名称"], col_factory + 1);
    let col_unit = optional_col(header, &["单位"], col_sup + 1);
    let col_qty = optional_col(header, &["数量", "入库数量"], col_unit + 1);
    let col_price = optional_col(header, &["进价", "成本单价", "单价"], col_qty + 1);
    let col_amt = optional_col(header, &["金额", "进价金额", "入库金额"], col_price + 1);
    let col_date = crate::core::excel_utils::find_col_idx(header, &["入库日期", "日期"]);

    let is_tcm = inbound_path.contains("中药");
    let subject_code_debit = "1201";
    let subject_code_credit = "220201";

    let v_no_str = voucher_no.unwrap_or("").trim();

    let mut items = Vec::new();
    let mut total_qty = 0.0;

    for row in rows.iter().skip(h_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let name = row_as_string(row, col_name);
        if is_summary_row(&name) {
            continue;
        }

        let spec = row_as_string(row, col_spec);
        let factory = row_as_string(row, col_factory);
        let supplier = row_as_string(row, col_sup);
        let unit = row_as_string(row, col_unit);
        let qty = row_as_f64(row, col_qty);
        let price = row_as_f64(row, col_price);
        let amt = amount_or_product(row_as_f64(row, col_amt), qty, price);
        let amount_cents = currency_cents(amt);

        total_qty += qty;

        items.push(InboundItem {
            supplier,
            name,
            spec,
            factory,
            qty,
            price,
            amount: cents_to_currency(amount_cents),
            amount_cents,
            unit,
        });
    }

    let source_dates = collect_source_dates(&rows, h_idx + 1, col_name, col_date);
    let voucher_date = crate::core::config::detect_voucher_date_with_source_dates(
        target_date_opt,
        &[inbound_path, ledger_path],
        &source_dates,
    );

    // 按供应商标准编码建立分组，保证每个供应商只有一个汇总分录，且明细连续。
    let supplier_groups = build_supplier_groups(&items, &supplier_dict);

    let mut suppliers_summary = Vec::new();
    let total_debit_cents: i64 = items.iter().map(|item| item.amount_cents).sum();
    let total_credit_cents: i64 = supplier_groups.iter().map(|group| group.amount_cents).sum();

    for group in &supplier_groups {
        let c_amt = cents_to_currency(group.amount_cents);

        suppliers_summary.push(SupplierSummary {
            supplier: group.supplier.clone(),
            code: group.code.clone(),
            brief: group.brief.clone(),
            count: group.item_indices.len(),
            amount: c_amt,
        });
    }

    let total_debit = cents_to_currency(total_debit_cents);
    let total_credit = cents_to_currency(total_credit_cents);
    let is_balanced = total_debit_cents == total_credit_cents;

    // 4. 生成凭证 Excel
    let mut out_wb = Workbook::new();
    let ws = out_wb.add_worksheet();
    ws.set_name("凭证模版").map_err(|e| e.to_string())?;
    let formats = VoucherFormats::new("#,##0.00");
    write_voucher_headers(ws, &formats)?;

    let mut row_idx: u32 = 1;
    let mut seq = 1;
    let mut matched_count = 0;
    let mut unmatched_items = Vec::new();

    // 写入“供应商明细借方分录 -> 该供应商贷方汇总分录”，再处理下一个供应商。
    for (group_idx, group) in supplier_groups.iter().enumerate() {
        for item_idx in &group.item_indices {
            let it = &items[*item_idx];
            let matched = resolve_drug_match(
                *item_idx,
                &it.name,
                &it.spec,
                &it.factory,
                &ledger_entries,
                config,
                confirmed_mappings,
            );
            let aux_code = if let Some(m) = matched {
                matched_count += 1;
                m.aux_code
            } else {
                let reason = if config
                    .strict_vendor_suffix_drugs
                    .iter()
                    .any(|d| clean_text(d) == clean_text(&it.name))
                {
                    "严格锁定厂家药品，总账未查到匹配厂家科目".to_string()
                } else {
                    "总账中未查到对应存货编码".to_string()
                };
                unmatched_items.push(UnmatchedDrug {
                    id: *item_idx,
                    name: it.name.clone(),
                    target_name: it.name.clone(),
                    spec: it.spec.clone(),
                    factory: it.factory.clone(),
                    supplier: it.supplier.clone(),
                    qty: it.qty,
                    price: it.price,
                    in_price: it.price,
                    amount: it.amount,
                    in_amt: it.amount,
                    reason,
                    candidates: build_ledger_candidates(&it.name, &ledger_entries),
                });
                String::new()
            };
            let summary_text = if !group.brief.is_empty() {
                group.brief.clone()
            } else {
                format!("入库-{}", it.name)
            };
            let line = VoucherLine {
                date: &voucher_date,
                voucher_no: (!v_no_str.is_empty()).then_some(v_no_str),
                seq,
                summary: Some(&summary_text),
                subject_code: Some(subject_code_debit),
                debit_amount: Some(it.amount),
                credit_amount: None,
                supplier_code: None,
                inventory_code: (!aux_code.is_empty()).then_some(aux_code.as_str()),
                qty: Some(it.qty),
                price: Some(it.price),
                original_amount: Some(it.amount),
            };
            write_voucher_line(ws, row_idx, &line, &formats)?;

            row_idx += 1;
            seq += 1;
        }

        // 写入当前供应商的贷方汇总分录。
        let sup = &suppliers_summary[group_idx];
        let line = VoucherLine {
            date: &voucher_date,
            voucher_no: (!v_no_str.is_empty()).then_some(v_no_str),
            seq,
            summary: Some(&sup.brief),
            subject_code: Some(subject_code_credit),
            debit_amount: None,
            credit_amount: Some(sup.amount),
            supplier_code: (!sup.code.is_empty()).then_some(sup.code.as_str()),
            inventory_code: None,
            qty: None,
            price: None,
            original_amount: None,
        };
        write_voucher_line(ws, row_idx, &line, &formats)?;

        row_idx += 1;
        seq += 1;
    }

    set_default_voucher_column_widths(
        ws,
        &[
            (0, 14.0),
            (5, 20.0),
            (7, 16.0),
            (8, 15.0),
            (9, 15.0),
            (11, 14.0),
            (15, 14.0),
        ],
    )?;
    copy_template_sheets(&mut out_wb, &tmpl_sheets)?;

    let out_path_buf = if let Some(co) = custom_output {
        PathBuf::from(co)
    } else {
        let parent = in_p.parent().unwrap_or_else(|| Path::new("."));
        let name_prefix = if is_tcm {
            "凭证导入模板_中药入库_已生成.xlsx"
        } else {
            "凭证导入模板_入库_已生成.xlsx"
        };
        parent.join(name_prefix)
    };

    out_wb
        .save(&out_path_buf)
        .map_err(|e| format!("保存入库凭证 Excel 失败: {}", e))?;

    let total_items = items.len();
    let total_entries = (seq - 1) as usize;
    let match_rate = if total_items > 0 {
        ((matched_count as f64 / total_items as f64) * 1000.0).round() / 10.0
    } else {
        0.0
    };

    Ok(InboundVoucherResult {
        success: true,
        output_file: out_path_buf.to_string_lossy().to_string(),
        voucher_date,
        total_items,
        total_entries,
        matched_count,
        unmatched_count: unmatched_items.len(),
        match_rate,
        total_qty: (total_qty * 100.0).round() / 100.0,
        total_debit_amt: total_debit,
        total_credit_amt: total_credit,
        is_balanced,
        suppliers_summary,
        unmatched_items,
        error: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::excel_utils::{cell_as_string, open_excel};
    use calamine::Reader;

    fn project_file(name: &str) -> String {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(name)
            .to_string_lossy()
            .to_string()
    }

    #[test]
    fn supplier_brief_contains_the_suffix_exactly_once() {
        assert_eq!(get_supplier_brief("河北蕴德药业有限公司"), "河北蕴德到货");
        assert_eq!(get_supplier_brief(""), "药品到货");
    }

    #[test]
    fn supplier_name_variants_are_merged_by_supplier_code() {
        let supplier_dict =
            HashMap::from([("004".to_string(), "河北国泰医药有限责任公司".to_string())]);
        let items = vec![
            InboundItem {
                supplier: "河北国泰".to_string(),
                name: "药品A".to_string(),
                spec: String::new(),
                factory: String::new(),
                qty: 1.0,
                price: 1.0,
                amount: 1.0,
                amount_cents: 100,
                unit: "盒".to_string(),
            },
            InboundItem {
                supplier: "河北国泰医药有限责任公司".to_string(),
                name: "药品B".to_string(),
                spec: String::new(),
                factory: String::new(),
                qty: 2.0,
                price: 1.0,
                amount: 2.0,
                amount_cents: 200,
                unit: "盒".to_string(),
            },
        ];

        let groups = build_supplier_groups(&items, &supplier_dict);

        assert_eq!(groups.len(), 1);
        assert_eq!(groups[0].code, "004");
        assert_eq!(groups[0].item_indices, vec![0, 1]);
        assert_eq!(groups[0].amount_cents, 300);
    }

    #[test]
    #[ignore = "依赖未提交的业务 Excel 样例；本地有样例时使用 cargo test --lib -- --ignored"]
    fn real_inbound_workbook_uses_business_date_and_correct_summary() {
        let (config, _) = crate::core::config::load_config(None);
        let output = std::env::temp_dir().join(format!(
            "desktop_app_inbound_test_{}.xlsx",
            std::process::id()
        ));
        let output_string = output.to_string_lossy().to_string();
        let inbound_path = project_file("西药-药品入库单.xlsx");
        let ledger_path = project_file("石家庄心理医院_数量金额总账_20260903150759.xlsx");
        let template_path = project_file("凭证导入模板-入库.xlsx");

        let result = generate_inbound_voucher(
            &inbound_path,
            &ledger_path,
            &template_path,
            Some(&output_string),
            None,
            Some("20"),
            &config,
            None,
        )
        .expect("真实入库样例应能生成凭证");

        assert_eq!(result.voucher_date, "2026-08-31");
        assert!(result.is_balanced);
        assert_eq!(
            currency_cents(result.total_debit_amt),
            currency_cents(result.total_credit_amt)
        );
        assert!(result
            .suppliers_summary
            .iter()
            .all(|summary| summary.amount >= 0.0));

        let mut workbook = open_excel(&output).expect("应能重新打开生成的入库凭证");
        let sheet_name = workbook
            .sheet_names()
            .first()
            .cloned()
            .expect("生成结果应包含凭证工作表");
        let range = workbook
            .worksheet_range(&sheet_name)
            .expect("应能读取生成的凭证工作表");
        let mut current_supplier: Option<String> = None;
        for row in range.rows().skip(1) {
            if row.len() > 5 {
                assert_eq!(cell_as_string(&row[1]), "记");
                assert_eq!(cell_as_string(&row[2]), "20");
                assert!(!cell_as_string(&row[5]).contains("到货到货"));

                match cell_as_string(&row[6]).as_str() {
                    "1201" => {
                        let summary = cell_as_string(&row[5]);
                        if let Some(expected) = &current_supplier {
                            assert_eq!(expected, &summary, "同一供应商的明细必须连续");
                        } else {
                            current_supplier = Some(summary);
                        }
                    }
                    "220201" => {
                        let summary = cell_as_string(&row[5]);
                        assert_eq!(
                            current_supplier.as_deref(),
                            Some(summary.as_str()),
                            "供应商汇总必须紧跟明细"
                        );
                        current_supplier = None;
                    }
                    _ => {}
                }
            }
        }
        assert!(
            current_supplier.is_none(),
            "每个供应商的明细都必须有对应汇总"
        );

        std::fs::remove_file(output).expect("应清理入库测试输出");
    }
}
