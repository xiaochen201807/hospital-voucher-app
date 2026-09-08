use crate::core::config::{clean_text, ConfigData};
use crate::core::excel_utils::{
    cents_to_currency, collect_source_dates, currency_cents, find_header_row, is_effectively_zero,
    is_summary_row, optional_col, read_sheet_rows, required_col, round_currency, row_as_f64,
    row_as_string,
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
pub use crate::core::matching::{
    extract_name_and_vendor_tag, match_drug, match_drug_with_method, strip_tcm_prefix,
};
pub use crate::core::models::{LedgerEntry, UnmatchedDrug};

/// 计算销售出库的基础单价。
///
/// 期末数量为 0 时仍按正常单价计算本次出库金额；该科目期末金额中的历史尾差
/// 由 `apply_zero_ending_tails` 摊到该科目的最后一笔出库明细，不能把整笔出库金额清零。
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

fn apply_zero_ending_tails(
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
    pub unmatched_items: Vec<UnmatchedDrug>,
    #[serde(default)]
    pub error: Option<String>,
}

/// 执行生成销售出库凭证
pub fn generate_outbound_voucher(
    sales_path: &str,
    ledger_path: &str,
    template_path: &str,
    custom_output: Option<&str>,
    target_date_opt: Option<&str>,
    fallback_price: bool,
    config: &ConfigData,
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
    let (_, rows) = read_sheet_rows(sales_p, &["汇总", "销售"], "销售表")?;
    if rows.len() < 2 {
        return Err("销售表中无有效数据".into());
    }

    let h_idx = find_header_row(&rows, &["药品名称", "品名"], 10, "销售表")?;
    let header = &rows[h_idx];
    let col_name = required_col(header, &["药品名称", "品名"], "药品名称")?;
    let col_spec = optional_col(header, &["规格"], col_name + 1);
    let col_factory = optional_col(header, &["制药厂", "生产厂家", "厂家"], col_spec + 2);
    let col_qty = optional_col(header, &["数量", "实发数量"], col_spec + 4);
    let col_price = optional_col(header, &["进价", "成本单价", "销售进价"], col_qty + 1);
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

    // 逐行匹配贷方明细。先暂存明细，待全部行匹配完成后，才能把每个零结存科目
    // 的尾差准确摊到该科目的最后一笔出库商品。
    let mut matched_count = 0;
    let mut unmatched_items = Vec::new();
    let mut total_qty = 0.0;
    let mut detail_lines = Vec::new();
    let mut tails_by_aux_code = HashMap::new();

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
        let raw_price = row_as_f64(row, col_price);

        total_qty += qty;

        let matched = match_drug(&name, &spec, &factory, &ledger_entries, config);

        let (aux_code, unit_price) = if let Some(m) = matched {
            matched_count += 1;
            let p = effective_outbound_unit_price(&m, raw_price, fallback_price);
            if is_effectively_zero(m.end_qty) && !m.aux_code.is_empty() {
                tails_by_aux_code.insert(m.aux_code.clone(), currency_cents(m.end_amount));
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

    // 对每个期末数量为 0 的存货，只在该存货最后一笔出库明细上处理期末历史尾差，
    // 不影响该存货前面的正常出库金额。
    apply_zero_ending_tails(&mut detail_lines, &tails_by_aux_code);
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
    fn zero_ending_quantity_keeps_base_price_and_records_tail() {
        let mut entry = ledger_entry("阿莫西林", "0.25g*24粒/盒");
        entry.end_qty = 0.0;
        entry.price = 0.42;
        entry.end_amount = 5.21;

        let unit_price = effective_outbound_unit_price(&entry, 0.18, true);
        assert_eq!(unit_price, 0.42);

        let mut lines = vec![OutboundDetailLine {
            aux_code: entry.aux_code.clone(),
            qty: 100.0,
            unit_price,
            amount_cents: currency_cents(100.0 * unit_price),
        }];
        let tails = HashMap::from([(entry.aux_code.clone(), currency_cents(entry.end_amount))]);
        let applied = apply_zero_ending_tails(&mut lines, &tails);

        assert_eq!(applied, 521);
        assert_eq!(lines[0].amount_cents, 4721);
    }

    #[test]
    fn zero_ending_tail_is_applied_only_to_the_last_outbound_line() {
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

        assert_eq!(apply_zero_ending_tails(&mut lines, &tails), 7);
        assert_eq!(lines[0].amount_cents, 300);
        assert_eq!(lines[1].amount_cents, 207);
    }

    #[test]
    fn tiny_ending_quantity_residual_also_receives_tail() {
        let mut entry = ledger_entry("阿莫西林", "0.25g*24粒/盒");
        entry.end_qty = -1e-10;
        entry.price = 0.42;
        entry.end_amount = -0.1;

        assert_eq!(effective_outbound_unit_price(&entry, 0.18, true), 0.42);
        assert_eq!(currency_cents(entry.end_amount), -10);
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
        let sales_path = project_file("2026.8月西药销售表_已汇总.xlsx");
        let ledger_path = project_file("石家庄心理医院_数量金额总账_20260903150759.xlsx");
        let template_path = project_file("凭证导入模板.xlsx");

        let result = generate_outbound_voucher(
            &sales_path,
            &ledger_path,
            &template_path,
            Some(&output_string),
            None,
            true,
            &config,
        )
        .expect("真实销售样例应能生成凭证");

        assert_eq!(result.voucher_date, "2026-08-31");
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
    }
}
