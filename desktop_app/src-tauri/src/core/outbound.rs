use crate::core::config::{clean_text, ConfigData};
use crate::core::excel_utils::{
    cents_to_currency, collect_source_dates, currency_cents, find_exact_col_idx, find_header_row,
    is_effectively_zero, is_summary_row, optional_col, read_sheet_rows_prefer, required_col,
    round_currency, row_as_f64, row_as_string,
};
use crate::core::voucher_writer::{
    copy_template_sheets, load_template_sheets, set_default_voucher_column_widths,
    write_voucher_headers, write_voucher_line, VoucherFormats, VoucherLine,
};
use rust_xlsxwriter::Workbook;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::{Path, PathBuf};

// 保留旧的模块路径，避免外部调用方因公共类型迁移而立即失效。
pub use crate::core::ledger::load_ledger_entries;
use crate::core::matching::{build_ledger_candidates, resolve_drug_match};
pub use crate::core::matching::{
    extract_name_and_vendor_tag, match_drug, match_drug_with_method, strip_tcm_prefix,
};
pub use crate::core::models::{ConfirmedLedgerMapping, LedgerEntry, UnmatchedDrug};

/// 计算销售出库的基础单价。
///
/// 先按正常单价计算本次出库的基础金额；只有在本次出库累计数量恰好清空该存货时，
/// 才会在后续步骤将总账金额与基础金额之间的差额摊到最后一笔出库明细。
fn effective_outbound_unit_price(
    ledger_entry: &LedgerEntry,
    sales_price: f64,
    fallback_price: bool,
) -> f64 {
    if ledger_entry.price > 0.0 {
        ledger_entry.price
    } else if fallback_price {
        sales_price
    } else {
        0.0
    }
}

#[derive(Debug)]
struct OutboundDetailLine {
    aux_code: String,
    qty: f64,
    unit_price: f64,
    amount_cents: i64,
}

#[derive(Debug, Clone, Copy)]
struct LedgerBalance {
    qty: f64,
    amount_cents: i64,
}

fn aggregate_outbound_totals(lines: &[OutboundDetailLine]) -> HashMap<String, (f64, i64)> {
    let mut outbound_totals: HashMap<String, (f64, i64)> = HashMap::new();
    for line in lines {
        if line.aux_code.is_empty() {
            continue;
        }

        let total = outbound_totals
            .entry(line.aux_code.clone())
            .or_insert((0.0, 0));
        total.0 += line.qty;
        total.1 += line.amount_cents;
    }
    outbound_totals
}

/// 出库数量不能超过第二次总账导出的期末数量。
///
/// 这个校验把“总账文件选错、销售表重复导入、库存不足”等问题挡在凭证生成之前，
/// 避免程序生成一张数量已经超过账面结存的错误凭证。
fn validate_outbound_quantities(
    lines: &[OutboundDetailLine],
    balances_by_aux_code: &HashMap<String, LedgerBalance>,
) -> Result<(), String> {
    let outbound_totals = aggregate_outbound_totals(lines);
    let mut over_issued = Vec::new();

    for (aux_code, (outbound_qty, _)) in outbound_totals {
        let Some(balance) = balances_by_aux_code.get(&aux_code) else {
            continue;
        };
        if outbound_qty - balance.qty > 1e-6 {
            over_issued.push(format!(
                "{}（出库 {:.4} > 总账结存 {:.4}）",
                aux_code, outbound_qty, balance.qty
            ));
        }
    }

    if over_issued.is_empty() {
        return Ok(());
    }

    let over_issued_count = over_issued.len();
    let displayed: Vec<String> = over_issued.into_iter().take(10).collect();
    let suffix = if displayed.len() < over_issued_count {
        format!("；另有 {} 个科目", over_issued_count - displayed.len())
    } else {
        String::new()
    };
    Err(format!(
        "本次出库数量超过所选总账的期末结存，未生成凭证：{}{}。请确认使用的是入库凭证导入后的第二次总账，并检查销售表是否重复。",
        displayed.join("；"),
        suffix
    ))
}

fn depletion_counts(
    lines: &[OutboundDetailLine],
    balances_by_aux_code: &HashMap<String, LedgerBalance>,
) -> (usize, usize) {
    let outbound_totals = aggregate_outbound_totals(lines);
    let mut fully_depleted_count = 0;
    let mut partially_depleted_count = 0;

    for (aux_code, (outbound_qty, _)) in outbound_totals {
        let Some(balance) = balances_by_aux_code.get(&aux_code) else {
            continue;
        };
        if is_effectively_zero(balance.qty) || is_effectively_zero(outbound_qty) {
            continue;
        }
        if is_effectively_zero(outbound_qty - balance.qty) {
            fully_depleted_count += 1;
        } else if outbound_qty < balance.qty {
            partially_depleted_count += 1;
        }
    }

    (fully_depleted_count, partially_depleted_count)
}

/// 计算本次出库后恰好清空的存货尾差。
///
/// 总账文件是本次出库前的余额快照。只有当本次出库数量累计等于该快照的
/// 期末数量时，才说明本次出库会把存货数量变为 0。此时应让该存货的出库
/// 合计金额等于总账快照金额，尾差为“总账金额 - 出库基础金额”。总账中
/// 历史遗留的“数量已为 0、金额仍非 0”记录不会因为自身数量为 0 被处理。
fn build_full_depletion_tails(
    lines: &[OutboundDetailLine],
    balances_by_aux_code: &HashMap<String, LedgerBalance>,
) -> HashMap<String, i64> {
    let outbound_totals = aggregate_outbound_totals(lines);

    let mut tails_by_aux_code = HashMap::new();
    for (aux_code, balance) in balances_by_aux_code {
        let Some((outbound_qty, outbound_amount_cents)) = outbound_totals.get(aux_code) else {
            continue;
        };

        if is_effectively_zero(balance.qty)
            || is_effectively_zero(*outbound_qty)
            || !is_effectively_zero(*outbound_qty - balance.qty)
        {
            continue;
        }

        let tail_cents = balance.amount_cents - *outbound_amount_cents;
        if tail_cents != 0 {
            tails_by_aux_code.insert(aux_code.clone(), tail_cents);
        }
    }

    tails_by_aux_code
}

fn apply_full_depletion_tails(
    lines: &mut [OutboundDetailLine],
    tails_by_aux_code: &HashMap<String, i64>,
) -> i64 {
    let mut last_line_by_aux_code = HashMap::new();
    for (index, line) in lines.iter().enumerate() {
        if !line.aux_code.is_empty() && !is_effectively_zero(line.qty) {
            last_line_by_aux_code.insert(line.aux_code.clone(), index);
        }
    }

    let mut applied_cents = 0_i64;
    for (aux_code, tail_cents) in tails_by_aux_code {
        if let Some(index) = last_line_by_aux_code.get(aux_code) {
            lines[*index].amount_cents += *tail_cents;
            applied_cents += *tail_cents;
        }
    }
    applied_cents
}

#[derive(Debug, Serialize, Deserialize)]
pub struct OutboundVoucherResult {
    pub success: bool,
    pub output_file: String,
    pub voucher_date: String,
    pub total_items: usize,
    pub matched_count: usize,
    pub unmatched_count: usize,
    pub match_rate: f64,
    pub total_qty: f64,
    pub total_credit_amt: f64,
    pub fully_depleted_count: usize,
    pub partially_depleted_count: usize,
    pub tail_adjustment_amt: f64,
    pub unmatched_items: Vec<UnmatchedDrug>,
    #[serde(default)]
    pub error: Option<String>,
}

/// 执行生成销售出库凭证
#[allow(clippy::too_many_arguments)]
pub fn generate_outbound_voucher(
    sales_path: &str,
    ledger_path: &str,
    template_path: &str,
    custom_output: Option<&str>,
    target_date_opt: Option<&str>,
    fallback_price: bool,
    config: &ConfigData,
    confirmed_mappings: Option<&[ConfirmedLedgerMapping]>,
) -> Result<OutboundVoucherResult, String> {
    let sales_p = Path::new(sales_path);
    let ledger_p = Path::new(ledger_path);
    let tmpl_p = Path::new(template_path);

    if !sales_p.exists() {
        return Err(format!("销售表 '{:?}' 不存在", sales_p));
    }
    if !ledger_p.exists() {
        return Err(format!("总账表 '{:?}' 不存在", ledger_p));
    }
    if !tmpl_p.exists() {
        return Err(format!("凭证模板文件 '{:?}' 不存在", tmpl_p));
    }

    // 1. 读取总账建立字典
    let ledger_entries = load_ledger_entries(ledger_p)?;

    // 2. 读取销售汇总表明细
    let (_, rows) =
        read_sheet_rows_prefer(sales_p, &["销售明细", "汇总"], &["汇总", "销售"], "销售表")?;
    if rows.len() < 2 {
        return Err("销售表中无有效数据".into());
    }

    let h_idx = find_header_row(&rows, &["药品名称", "品名"], 10, "销售表")?;
    let header = &rows[h_idx];
    let col_name = required_col(header, &["药品名称", "品名"], "药品名称")?;
    let col_spec = optional_col(header, &["规格"], col_name + 1);
    let col_factory = optional_col(header, &["制药厂", "生产厂家", "厂家"], col_spec + 2);
    let col_qty = optional_col(header, &["数量", "实发数量"], col_spec + 4);
    let col_price = find_exact_col_idx(header, &["进价", "成本单价", "销售进价"]);
    let col_cost_amt = find_exact_col_idx(header, &["进价金额", "成本金额"]);
    let col_date = crate::core::excel_utils::find_col_idx(header, &["销售日期", "日期"]);
    let source_dates = collect_source_dates(&rows, h_idx + 1, col_name, col_date);

    // 凭证日期动态推断
    let voucher_date = crate::core::config::detect_voucher_date_with_source_dates(
        target_date_opt,
        &[sales_path, ledger_path],
        &source_dates,
    );

    // 读取原凭证模板中所有 Sheet 数据。
    let tmpl_sheets = load_template_sheets(tmpl_p)?;

    // 3. 开始使用 rust_xlsxwriter 构造凭证导入表
    let mut wb_out = Workbook::new();
    let ws_voucher = wb_out.add_worksheet();
    ws_voucher.set_name("凭证模版").map_err(|e| e.to_string())?;

    let formats = VoucherFormats::new("#,##0");
    write_voucher_headers(ws_voucher, &formats)?;

    // 逐行匹配贷方明细。先暂存明细，待全部行匹配完成后，才能按本次出库累计数量
    // 判断哪些存货恰好清零，并把对应尾差准确摊到最后一笔出库商品。
    let mut matched_count = 0;
    let mut unmatched_items = Vec::new();
    let mut total_qty = 0.0;
    let mut detail_lines = Vec::new();
    let mut balances_by_aux_code = HashMap::new();

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
        let qty = row_as_f64(row, col_qty);
        let raw_price = col_price
            .map(|col| row_as_f64(row, col))
            .filter(|price| !is_effectively_zero(*price))
            .or_else(|| {
                col_cost_amt.and_then(|col| {
                    (!is_effectively_zero(qty)).then_some(row_as_f64(row, col) / qty)
                })
            })
            .unwrap_or(0.0);
        let item_id = detail_lines.len();

        total_qty += qty;

        let matched = resolve_drug_match(
            item_id,
            &name,
            &spec,
            &factory,
            &ledger_entries,
            config,
            confirmed_mappings,
        );

        let (aux_code, unit_price) = if let Some(m) = matched {
            matched_count += 1;
            let p = effective_outbound_unit_price(&m, raw_price, fallback_price);
            if !m.aux_code.is_empty() {
                balances_by_aux_code
                    .entry(m.aux_code.clone())
                    .or_insert(LedgerBalance {
                        qty: m.end_qty,
                        amount_cents: currency_cents(m.end_amount),
                    });
            }
            (m.aux_code, p)
        } else {
            let reason = if config
                .strict_vendor_suffix_drugs
                .iter()
                .any(|d| clean_text(d) == clean_text(&name))
            {
                "严格锁定厂家药品，总账中未找到匹配的厂家细分科目".to_string()
            } else {
                "总账中未检索到匹配的存货科目编码".to_string()
            };
            let amt = round_currency(qty * raw_price);
            unmatched_items.push(UnmatchedDrug {
                id: item_id,
                name: name.clone(),
                target_name: name.clone(),
                spec: spec.clone(),
                factory: factory.clone(),
                supplier: factory.clone(),
                qty,
                price: raw_price,
                in_price: raw_price,
                amount: amt,
                in_amt: amt,
                reason,
                candidates: build_ledger_candidates(&name, &ledger_entries),
            });
            let p = if fallback_price { raw_price } else { 0.0 };
            (String::new(), p)
        };

        detail_lines.push(OutboundDetailLine {
            aux_code,
            qty,
            unit_price,
            amount_cents: currency_cents(qty * unit_price),
        });
    }

    validate_outbound_quantities(&detail_lines, &balances_by_aux_code)?;

    // 只对本次出库后恰好变为 0 的存货，在该存货最后一笔出库明细上处理尾差，
    // 不影响该存货前面的正常出库金额，也不处理历史遗留的零数量余额。
    let tails_by_aux_code = build_full_depletion_tails(&detail_lines, &balances_by_aux_code);
    let tail_adjustment_cents = apply_full_depletion_tails(&mut detail_lines, &tails_by_aux_code);
    let (fully_depleted_count, partially_depleted_count) =
        depletion_counts(&detail_lines, &balances_by_aux_code);
    let total_credit_cents: i64 = detail_lines.iter().map(|line| line.amount_cents).sum();
    let total_credit_amt = cents_to_currency(total_credit_cents);

    for (index, detail) in detail_lines.iter().enumerate() {
        let credit_amt = cents_to_currency(detail.amount_cents);
        let line = VoucherLine {
            date: &voucher_date,
            voucher_no: None,
            seq: (index + 2) as u32,
            summary: None,
            subject_code: None,
            debit_amount: None,
            credit_amount: Some(credit_amt),
            supplier_code: None,
            inventory_code: (!detail.aux_code.is_empty()).then_some(detail.aux_code.as_str()),
            qty: Some(detail.qty),
            price: Some(detail.unit_price),
            original_amount: Some(credit_amt),
        };
        write_voucher_line(ws_voucher, (index + 2) as u32, &line, &formats)?;
    }

    // 第 1 行 (分录序号 1): 借方平衡行。科目仍由财务按业务性质填写，金额自动
    // 等于所有贷方明细按分取整后的合计，避免尾差导致导入凭证借贷不平。
    let opening_line = VoucherLine {
        date: &voucher_date,
        voucher_no: None,
        seq: 1,
        summary: None,
        subject_code: None,
        debit_amount: Some(total_credit_amt),
        credit_amount: None,
        supplier_code: None,
        inventory_code: None,
        qty: None,
        price: None,
        original_amount: None,
    };
    write_voucher_line(ws_voucher, 1, &opening_line, &formats)?;

    set_default_voucher_column_widths(ws_voucher, &[(0, 14.0), (9, 15.0), (15, 14.0), (23, 15.0)])?;
    copy_template_sheets(&mut wb_out, &tmpl_sheets)?;

    // 确定输出路径
    let out_path_buf = if let Some(co) = custom_output {
        PathBuf::from(co)
    } else {
        let parent = sales_p.parent().unwrap_or_else(|| Path::new("."));
        parent.join("凭证导入模板_已生成.xlsx")
    };

    wb_out
        .save(&out_path_buf)
        .map_err(|e| format!("保存出库凭证 Excel 失败: {}", e))?;

    let total_items = detail_lines.len();
    let match_rate = if total_items > 0 {
        ((matched_count as f64 / total_items as f64) * 1000.0).round() / 10.0
    } else {
        0.0
    };

    Ok(OutboundVoucherResult {
        success: true,
        output_file: out_path_buf.to_string_lossy().to_string(),
        voucher_date,
        total_items,
        matched_count,
        unmatched_count: unmatched_items.len(),
        match_rate,
        total_qty: round_currency(total_qty),
        total_credit_amt,
        fully_depleted_count,
        partially_depleted_count,
        tail_adjustment_amt: cents_to_currency(tail_adjustment_cents),
        unmatched_items,
        error: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use calamine::Reader;

    fn project_file(name: &str) -> String {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(name)
            .to_string_lossy()
            .to_string()
    }

    fn ledger_entry(drug_name: &str, spec: &str) -> LedgerEntry {
        LedgerEntry {
            code: "1201_TEST".to_string(),
            aux_code: "TEST".to_string(),
            name_full: format!("存货_{} {}", drug_name, spec),
            drug_name: drug_name.to_string(),
            spec: spec.to_string(),
            price: 1.0,
            end_qty: 1.0,
            end_amount: 0.0,
        }
    }

    #[test]
    fn full_depletion_adjusts_last_line_to_the_ledger_amount() {
        let mut entry = ledger_entry("阿莫西林", "0.25g*24粒/盒");
        entry.end_qty = 266.0;
        entry.price = 0.80;
        entry.end_amount = 212.70;

        let unit_price = effective_outbound_unit_price(&entry, 0.793, true);
        assert_eq!(unit_price, 0.80);

        let mut lines = vec![OutboundDetailLine {
            aux_code: entry.aux_code.clone(),
            qty: 266.0,
            unit_price,
            amount_cents: currency_cents(266.0 * unit_price),
        }];
        let balances = HashMap::from([(
            entry.aux_code.clone(),
            LedgerBalance {
                qty: entry.end_qty,
                amount_cents: currency_cents(entry.end_amount),
            },
        )]);
        let tails = build_full_depletion_tails(&lines, &balances);
        let applied = apply_full_depletion_tails(&mut lines, &tails);

        assert_eq!(applied, -10);
        assert_eq!(lines[0].amount_cents, 21270);
    }

    #[test]
    fn full_depletion_tail_is_applied_only_to_the_last_outbound_line() {
        let aux_code = "TEST".to_string();
        let mut lines = vec![
            OutboundDetailLine {
                aux_code: aux_code.clone(),
                qty: 3.0,
                unit_price: 1.0,
                amount_cents: 300,
            },
            OutboundDetailLine {
                aux_code: aux_code.clone(),
                qty: 2.0,
                unit_price: 1.0,
                amount_cents: 200,
            },
        ];
        let tails = HashMap::from([(aux_code, 7_i64)]);

        assert_eq!(apply_full_depletion_tails(&mut lines, &tails), 7);
        assert_eq!(lines[0].amount_cents, 300);
        assert_eq!(lines[1].amount_cents, 207);
    }

    #[test]
    fn partial_depletion_does_not_receive_tail() {
        let mut entry = ledger_entry("阿莫西林", "0.25g*24粒/盒");
        entry.end_qty = 266.0;
        entry.price = 0.80;
        entry.end_amount = 212.70;

        let lines = vec![OutboundDetailLine {
            aux_code: entry.aux_code.clone(),
            qty: 265.0,
            unit_price: entry.price,
            amount_cents: currency_cents(265.0 * entry.price),
        }];
        let balances = HashMap::from([(
            entry.aux_code,
            LedgerBalance {
                qty: entry.end_qty,
                amount_cents: currency_cents(entry.end_amount),
            },
        )]);

        assert!(build_full_depletion_tails(&lines, &balances).is_empty());
    }

    #[test]
    fn historical_zero_balance_does_not_receive_tail_without_full_depletion() {
        let mut entry = ledger_entry("阿莫西林", "0.25g*24粒/盒");
        entry.end_qty = 0.0;
        entry.price = 0.42;
        entry.end_amount = -0.1;

        let lines = vec![OutboundDetailLine {
            aux_code: entry.aux_code.clone(),
            qty: 100.0,
            unit_price: entry.price,
            amount_cents: currency_cents(100.0 * entry.price),
        }];
        let balances = HashMap::from([(
            entry.aux_code,
            LedgerBalance {
                qty: entry.end_qty,
                amount_cents: currency_cents(entry.end_amount),
            },
        )]);

        assert!(build_full_depletion_tails(&lines, &balances).is_empty());
    }

    #[test]
    fn over_issued_quantity_is_rejected_before_voucher_generation() {
        let aux_code = "TEST".to_string();
        let lines = vec![OutboundDetailLine {
            aux_code: aux_code.clone(),
            qty: 11.0,
            unit_price: 1.0,
            amount_cents: 1100,
        }];
        let balances = HashMap::from([(
            aux_code.clone(),
            LedgerBalance {
                qty: 10.0,
                amount_cents: 1000,
            },
        )]);

        let error =
            validate_outbound_quantities(&lines, &balances).expect_err("超出结存的出库必须被拦截");
        assert!(error.contains("TEST"));
        assert!(error.contains("出库 11.0000 > 总账结存 10.0000"));
    }

    #[test]
    fn strict_drug_does_not_fall_back_to_generic_entry_without_vendor() {
        let mut config = ConfigData::default();
        config.strict_vendor_suffix_drugs.push("海螵蛸".to_string());
        let ledger = vec![ledger_entry("海螵蛸", "1克*1000克/袋")];

        assert!(match_drug("海螵蛸", "1克*1000克/袋", "", &ledger, &config).is_none());
    }

    #[test]
    fn strict_drug_matches_only_the_configured_vendor_entry() {
        let mut config = ConfigData::default();
        config.strict_vendor_suffix_drugs.push("海螵蛸".to_string());
        config
            .factory_abbreviations
            .insert("河北蕴德药业有限公司".to_string(), "蕴德".to_string());
        let ledger = vec![ledger_entry("海螵蛸蕴德", "1克*1000克/袋")];

        let matched = match_drug(
            "海螵蛸",
            "1克*1000克/袋",
            "河北蕴德药业有限公司",
            &ledger,
            &config,
        )
        .expect("应命中带厂家后缀的科目");
        assert_eq!(matched.drug_name, "海螵蛸蕴德");
    }

    #[test]
    fn truncated_vendor_matches_ledger_bracket_entry() {
        let mut config = ConfigData::default();
        config
            .factory_abbreviations
            .insert("河北国瑞堂药业有限公司".to_string(), "国瑞堂".to_string());

        let ledger = vec![
            ledger_entry("麸炒苍术（国松堂）", "1克*1000克/袋"),
            ledger_entry("麸炒苍术（国瑞堂）", "1克*1000克/袋"),
        ];

        // 库管系统导出被截断为 "河北国瑞堂药业有限公"
        let matched = match_drug(
            "麸炒苍术",
            "1克*1000克/袋",
            "河北国瑞堂药业有限公",
            &ledger,
            &config,
        )
        .expect("截断厂家名称必须能识别并命中对应国瑞堂条目");

        assert_eq!(matched.drug_name, "麸炒苍术（国瑞堂）");
    }

    #[test]
    fn tcm_processing_prefix_tolerance_matches() {
        let config = ConfigData::default();
        let ledger = vec![ledger_entry("制吴茱萸", "1克*1000克/袋")];

        let matched = match_drug("吴茱萸", "1克*1000克/袋", "", &ledger, &config)
            .expect("库管通用名应能兼容匹配财务炮制前缀");

        assert_eq!(matched.drug_name, "制吴茱萸");
    }

    #[test]
    fn real_outbound_workbook_still_generates_successfully() {
        let (config, _) = crate::core::config::load_config(None);
        let output = std::env::temp_dir().join(format!(
            "desktop_app_outbound_test_{}.xlsx",
            std::process::id()
        ));
        let output_string = output.to_string_lossy().to_string();
        let raw_sales_path = project_file("2026.8月西药销售表.xls");
        let sales_output = std::env::temp_dir().join(format!(
            "desktop_app_sales_for_outbound_test_{}.xlsx",
            std::process::id()
        ));
        let sales_output_string = sales_output.to_string_lossy().to_string();
        let sales_result = crate::core::sales::process_sales_file(
            &raw_sales_path,
            Some(&sales_output_string),
            None,
        )
        .expect("出库集成测试应先生成销售汇总表");
        assert_eq!(sales_result.totals.unique_count, 54);

        let source_ledger_path = project_file("石家庄心理医院_数量金额总账_20260903150759.xlsx");
        let ledger_output = std::env::temp_dir().join(format!(
            "desktop_app_sufficient_ledger_for_outbound_test_{}.xlsx",
            std::process::id()
        ));
        let ledger_output_string = ledger_output.to_string_lossy().to_string();
        let source_entries = load_ledger_entries(Path::new(&source_ledger_path))
            .expect("出库集成测试应能读取总账样例");
        let mut ledger_wb = Workbook::new();
        let ledger_ws = ledger_wb.add_worksheet();
        ledger_ws
            .set_name("数量金额总账")
            .expect("应能创建测试总账工作表");
        for (row_idx, entry) in source_entries.iter().enumerate() {
            ledger_ws
                .write_string(row_idx as u32, 0, &entry.code)
                .expect("应能写入测试科目编码");
            ledger_ws
                .write_string(row_idx as u32, 1, &entry.name_full)
                .expect("应能写入测试科目名称");
            ledger_ws
                .write_number(row_idx as u32, 16, 1_000_000.0)
                .expect("应能写入测试期末数量");
            ledger_ws
                .write_number(row_idx as u32, 17, entry.price)
                .expect("应能写入测试期末单价");
            ledger_ws
                .write_number(row_idx as u32, 18, 1_000_000.0 * entry.price)
                .expect("应能写入测试期末金额");
        }
        ledger_wb.save(&ledger_output).expect("应能保存测试总账");

        let template_path = project_file("凭证导入模板.xlsx");

        let result = generate_outbound_voucher(
            &sales_output_string,
            &ledger_output_string,
            &template_path,
            Some(&output_string),
            Some("2026-08-31"),
            true,
            &config,
            None,
        )
        .expect("真实销售样例应能生成凭证");

        assert_eq!(result.voucher_date, "2026-08-31");
        assert_eq!(result.total_items, 54);
        assert!(output.exists());

        let mut output_wb =
            crate::core::excel_utils::open_excel(&output).expect("生成文件应能重新打开");
        let output_range = output_wb
            .worksheet_range("凭证模版")
            .expect("生成文件应包含凭证模版工作表");
        let output_rows: Vec<_> = output_range.rows().collect();
        let debit_cents = output_rows
            .get(1)
            .and_then(|row| row.get(8))
            .map(crate::core::excel_utils::cell_as_f64)
            .map(currency_cents)
            .unwrap_or_default();
        let credit_cents: i64 = output_rows
            .iter()
            .skip(2)
            .filter_map(|row| row.get(9))
            .map(crate::core::excel_utils::cell_as_f64)
            .map(currency_cents)
            .sum();

        assert_eq!(debit_cents, currency_cents(result.total_credit_amt));
        assert_eq!(debit_cents, credit_cents);
        std::fs::remove_file(output).expect("应清理出库测试输出");
        std::fs::remove_file(sales_output).expect("应清理销售汇总测试输出");
        std::fs::remove_file(ledger_output).expect("应清理测试总账输出");
    }
}
