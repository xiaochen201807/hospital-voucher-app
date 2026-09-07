use crate::core::config::{clean_text, ConfigData};
use crate::core::excel_utils::{cell_as_f64, cell_as_string, find_col_idx, open_excel};
use calamine::Reader;
use rust_xlsxwriter::{Format, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::path::{Path, PathBuf};

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct UnmatchedDrug {
    pub name: String,
    pub target_name: String,
    pub spec: String,
    pub factory: String,
    pub supplier: String,
    pub qty: f64,
    pub price: f64,
    pub in_price: f64,
    pub amount: f64,
    pub in_amt: f64,
    pub reason: String,
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

#[derive(Debug, Clone)]
pub struct LedgerEntry {
    pub code: String,      // 如 1201_XY0001
    pub aux_code: String,  // 如 XY0001
    pub name_full: String, // 如 存货_西地兰注射液 0.4mg*2ml*10支/盒
    pub drug_name: String, // 清洗后的药名
    pub spec: String,
    pub price: f64,
    pub end_qty: f64,
}

/// 读取财务总账建立存货字典
pub fn load_ledger_entries(ledger_path: &Path) -> Result<Vec<LedgerEntry>, String> {
    let mut wb = open_excel(ledger_path)?;
    let sheet_name = wb
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "总账中无有效工作表".to_string())?;

    let range = wb
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取总账失败: {}", e))?;

    let mut entries = Vec::new();

    for row in range.rows() {
        if row.len() < 4 {
            continue;
        }
        let code = cell_as_string(&row[0]);
        if !code.starts_with("1201_") {
            continue;
        }

        let name_full = cell_as_string(&row[1]);
        let aux_code = code.replace("1201_", "");

        // 解析药名与规格（通常格式为 "存货_药名 规格" 或 "存货_药名"）
        let raw_name = name_full.trim_start_matches("存货_").trim();
        let parts: Vec<&str> = raw_name.splitn(2, ' ').collect();
        let drug_name = parts[0].to_string();
        let spec = if parts.len() > 1 { parts[1].to_string() } else { String::new() };

        // 结存数量位于第 17 列 (0-indexed: 16)，单价位于第 18 列 (17)
        let end_qty = if row.len() > 16 { cell_as_f64(&row[16]) } else { 0.0 };
        let price = if row.len() > 17 { cell_as_f64(&row[17]) } else { 0.0 };

        entries.push(LedgerEntry {
            code,
            aux_code,
            name_full,
            drug_name,
            spec,
            price,
            end_qty,
        });
    }

    Ok(entries)
}

/// 匹配药品存货科目
pub fn match_drug(
    drug_name: &str,
    spec: &str,
    factory: &str,
    ledger: &[LedgerEntry],
    config: &ConfigData,
) -> Option<LedgerEntry> {
    let clean_drug = clean_text(drug_name);
    let clean_sp = clean_text(spec);

    // 提取厂家简称
    let mut vendor_suffix = String::new();
    if !factory.is_empty() {
        for (full, brief) in &config.factory_abbreviations {
            if factory.contains(full) || clean_text(factory).contains(&clean_text(full)) {
                vendor_suffix = brief.clone();
                break;
            }
        }
    }

    // 0. 特殊手动指定覆盖（支持带厂家后缀的复合药名以及裸药名）
    let mut override_cands = Vec::new();
    if !vendor_suffix.is_empty() {
        override_cands.push(clean_text(&format!("{}({})", drug_name, vendor_suffix)));
        override_cands.push(clean_text(&format!("{}（{}）", drug_name, vendor_suffix)));
        override_cands.push(clean_text(&format!("{}{}", drug_name, vendor_suffix)));
    }
    if !factory.is_empty() {
        override_cands.push(clean_text(&format!("{}({})", drug_name, factory)));
        override_cands.push(clean_text(&format!("{}（{}）", drug_name, factory)));
    }
    if !spec.is_empty() {
        override_cands.push(clean_text(&format!("{}{}", drug_name, spec)));
    }
    override_cands.push(clean_drug.clone());

    for (k, v) in &config.drug_code_overrides {
        let clean_k = clean_text(k);
        if override_cands.iter().any(|c| c == &clean_k) {
            let trimmed_v = v.trim();
            // 如果配置为空字符串 ""，代表该品规明确需拦截人工建档，立即返回 None 阻断穿透！
            if trimmed_v.is_empty() {
                return None;
            }
            if let Some(target) = ledger.iter().find(|e| e.code == trimmed_v || e.aux_code == trimmed_v) {
                return Some(target.clone());
            } else {
                // 手动指定了编码但在总账中未找到对应科目，同样安全返回 None，严禁盲目穿透
                return None;
            }
        }
    }

    // 严格厂家后缀保护（如海螵蛸）
    let is_strict_drug = config
        .strict_vendor_suffix_drugs
        .iter()
        .any(|d| clean_text(d) == clean_drug);

    // 优先级 1: 全名 + 厂家全词精准匹配
    if !vendor_suffix.is_empty() {
        let expected_with_vendor = format!("{}{}", clean_drug, clean_text(&vendor_suffix));
        for e in ledger {
            let e_clean = clean_text(&e.drug_name);
            if e_clean == expected_with_vendor {
                return Some(e.clone());
            }
        }
    }

    // 严格厂家药品只允许命中“药名 + 厂家”的科目；厂家为空、未配置或未命中时，
    // 都不能回退到同名通用科目，否则会造成串户。
    if is_strict_drug {
        return None;
    }

    // 优先级 2: 全名 + 规格精准一致
    for e in ledger {
        let e_name = clean_text(&e.drug_name);
        let e_spec = clean_text(&e.spec);
        if e_name == clean_drug && !clean_sp.is_empty() && e_spec == clean_sp {
            return Some(e.clone());
        }
    }

    // 优先级 3: 仅全名完全一致（若有多条候选，必须通过规格明确排他匹配，严禁盲目取第一条）
    let candidates: Vec<&LedgerEntry> = ledger
        .iter()
        .filter(|e| clean_text(&e.drug_name) == clean_drug)
        .collect();

    if candidates.len() == 1 {
        // 如果是严格厂家保护药，且总账无厂家条目，禁止回退到无厂家的通用条目
        if is_strict_drug {
            return None;
        }
        return Some(candidates[0].clone());
    } else if candidates.len() > 1 {
        // 尝试规格精准匹配或包含比对
        let mut matched_cand = None;
        let mut match_count = 0;
        for c in &candidates {
            let c_sp = clean_text(&c.spec);
            if !c_sp.is_empty() && !clean_sp.is_empty() && (clean_sp == c_sp || clean_sp.contains(&c_sp) || c_sp.contains(&clean_sp)) {
                matched_cand = Some((*c).clone());
                match_count += 1;
            }
        }
        // 只有当规格能唯一确切命中时才返回，存在歧义时返回 None 要求人工确认
        if match_count == 1 {
            return matched_cand;
        }
        return None;
    }

    None
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
    let mut sales_wb = open_excel(sales_p)?;
    let sheet_names = sales_wb.sheet_names();
    let s_name = sheet_names
        .iter()
        .find(|s| s.contains("汇总") || s.contains("销售"))
        .unwrap_or(&sheet_names[0]);

    let range = sales_wb
        .worksheet_range(s_name)
        .map_err(|e| format!("读取销售表失败: {}", e))?;

    let rows: Vec<Vec<calamine::Data>> = range.rows().map(|r| r.to_vec()).collect();
    if rows.len() < 2 {
        return Err("销售表中无有效数据".into());
    }

    let mut h_idx = 0;
    for (i, r) in rows.iter().take(10).enumerate() {
        if find_col_idx(r, &["药品名称", "品名"]).is_some() {
            h_idx = i;
            break;
        }
    }

    let header = &rows[h_idx];
    let col_name = find_col_idx(header, &["药品名称", "品名"]).unwrap();
    let col_spec = find_col_idx(header, &["规格"]).unwrap_or(col_name + 1);
    let col_factory = find_col_idx(header, &["制药厂", "生产厂家", "厂家"]).unwrap_or(col_spec + 2);
    let col_qty = find_col_idx(header, &["数量", "实发数量"]).unwrap_or(col_spec + 4);
    let col_price = find_col_idx(header, &["进价", "成本单价", "销售进价"]).unwrap_or(col_qty + 1);
    let col_date = find_col_idx(header, &["销售日期", "日期"]);

    let mut source_dates = Vec::new();
    if let Some(date_col) = col_date {
        for row in rows.iter().skip(h_idx + 1) {
            let name = cell_as_string(row.get(col_name).unwrap_or(&calamine::Data::Empty));
            if name.is_empty() || name.contains("合计") || name.contains("总计") {
                continue;
            }
            if let Some(date_cell) = row.get(date_col) {
                let date_value = cell_as_string(date_cell);
                if !date_value.is_empty() {
                    source_dates.push(date_value);
                }
            }
        }
    }

    // 凭证日期动态推断
    let voucher_date = crate::core::config::detect_voucher_date_with_source_dates(
        target_date_opt,
        &[sales_path, ledger_path],
        &source_dates,
    );

    // 读取原凭证模板中所有 Sheet 数据。
    let mut tmpl_sheets = Vec::new();
    let mut tmpl_wb = open_excel(tmpl_p)?;
    for s in tmpl_wb.sheet_names() {
        let range = tmpl_wb
            .worksheet_range(&s)
            .map_err(|e| format!("读取凭证模板工作表 '{}' 失败: {}", s, e))?;
        let r_rows: Vec<Vec<calamine::Data>> = range.rows().map(|r| r.to_vec()).collect();
        tmpl_sheets.push((s.to_string(), r_rows));
    }

    // 3. 开始使用 rust_xlsxwriter 构造凭证导入表
    let mut wb_out = Workbook::new();
    let ws_voucher = wb_out.add_worksheet();
    ws_voucher.set_name("凭证模版").map_err(|e| e.to_string())?;

    // 样式定义
    let fmt_header = Format::new()
        .set_bold()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_text = Format::new()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let _fmt_left = Format::new()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Left)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_money = Format::new()
        .set_font_size(10)
        .set_num_format("#,##0.00")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_qty = Format::new()
        .set_font_size(10)
        .set_num_format("#,##0")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    // 写入第 0 行标准表头 (26 列)
    let template_headers = [
        "日期", "凭证字", "凭证号", "附件数", "分录序号", "摘要", "科目代码", "科目名称",
        "借方金额", "贷方金额", "客户", "供应商", "职员", "项目", "部门", "存货",
        "是否限定", "自定义类别", "自定义编码", "自定义类别1", "自定义编码1", "数量",
        "单价", "原币金额", "币别", "汇率",
    ];

    ws_voucher.set_row_height(0, 26).map_err(|e| e.to_string())?;
    for (c, h) in template_headers.iter().enumerate() {
        ws_voucher
            .write_string_with_format(0, c as u16, *h, &fmt_header)
            .map_err(|e| e.to_string())?;
    }

    // 第 1 行 (分录序号 1): 借方待填空行
    ws_voucher.set_row_height(1, 20).map_err(|e| e.to_string())?;
    ws_voucher.write_string_with_format(1, 0, &voucher_date, &fmt_text).map_err(|e| e.to_string())?;
    ws_voucher.write_string_with_format(1, 1, "记", &fmt_text).map_err(|e| e.to_string())?;
    ws_voucher.write_number_with_format(1, 4, 1.0, &fmt_text).map_err(|e| e.to_string())?;
    ws_voucher.write_string_with_format(1, 24, "RMB", &fmt_text).map_err(|e| e.to_string())?;
    ws_voucher.write_number_with_format(1, 25, 1.0, &fmt_text).map_err(|e| e.to_string())?;
    for c in 0..26 {
        if c != 0 && c != 1 && c != 4 && c != 24 && c != 25 {
            ws_voucher.write_blank(1, c as u16, &fmt_text).map_err(|e| e.to_string())?;
        }
    }

    // 逐行匹配并填充贷方明细行 (分录序号从 2 递增)
    let mut matched_count = 0;
    let mut unmatched_items = Vec::new();
    let mut total_credit_amt = 0.0;
    let mut total_qty = 0.0;
    let mut current_row: u32 = 2;
    let mut entry_seq = 2;

    for row in rows.iter().skip(h_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let name = cell_as_string(row.get(col_name).unwrap_or(&calamine::Data::Empty));
        if name.is_empty() || name.contains("合计") || name.contains("总计") {
            continue;
        }

        let spec = row.get(col_spec).map(cell_as_string).unwrap_or_default();
        let factory = row.get(col_factory).map(cell_as_string).unwrap_or_default();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let raw_price = row.get(col_price).map(cell_as_f64).unwrap_or(0.0);

        total_qty += qty;

        let matched = match_drug(&name, &spec, &factory, &ledger_entries, config);

        let (aux_code, unit_price) = if let Some(m) = matched {
            matched_count += 1;
            let p = if m.price > 0.0 { m.price } else if fallback_price { raw_price } else { 0.0 };
            (m.aux_code, p)
        } else {
            let reason = if config.strict_vendor_suffix_drugs.iter().any(|d| clean_text(d) == clean_text(&name)) {
                "严格锁定厂家药品，总账中未找到匹配的厂家细分科目".to_string()
            } else {
                "总账中未检索到匹配的存货科目编码".to_string()
            };
            let amt = (qty * raw_price * 100.0).round() / 100.0;
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

        let credit_amt = (qty * unit_price * 100.0).round() / 100.0;
        total_credit_amt += credit_amt;

        ws_voucher.set_row_height(current_row, 20).map_err(|e| e.to_string())?;

        ws_voucher.write_string_with_format(current_row, 0, &voucher_date, &fmt_text).map_err(|e| e.to_string())?;
        ws_voucher.write_string_with_format(current_row, 1, "记", &fmt_text).map_err(|e| e.to_string())?;
        ws_voucher.write_blank(current_row, 2, &fmt_text).map_err(|e| e.to_string())?;
        ws_voucher.write_blank(current_row, 3, &fmt_text).map_err(|e| e.to_string())?;
        ws_voucher.write_number_with_format(current_row, 4, entry_seq as f64, &fmt_text).map_err(|e| e.to_string())?;

        for c in 5..9 {
            ws_voucher.write_blank(current_row, c as u16, &fmt_text).map_err(|e| e.to_string())?;
        }

        // 贷方金额
        ws_voucher.write_number_with_format(current_row, 9, credit_amt, &fmt_money).map_err(|e| e.to_string())?;

        for c in 10..15 {
            ws_voucher.write_blank(current_row, c as u16, &fmt_text).map_err(|e| e.to_string())?;
        }

        // 存货代码
        if !aux_code.is_empty() {
            ws_voucher.write_string_with_format(current_row, 15, &aux_code, &fmt_text).map_err(|e| e.to_string())?;
        } else {
            ws_voucher.write_blank(current_row, 15, &fmt_text).map_err(|e| e.to_string())?;
        }

        for c in 16..21 {
            ws_voucher.write_blank(current_row, c as u16, &fmt_text).map_err(|e| e.to_string())?;
        }

        // 数量、单价、原币金额
        ws_voucher.write_number_with_format(current_row, 21, qty, &fmt_qty).map_err(|e| e.to_string())?;
        ws_voucher.write_number_with_format(current_row, 22, unit_price, &fmt_money).map_err(|e| e.to_string())?;
        ws_voucher.write_number_with_format(current_row, 23, credit_amt, &fmt_money).map_err(|e| e.to_string())?;
        ws_voucher.write_string_with_format(current_row, 24, "RMB", &fmt_text).map_err(|e| e.to_string())?;
        ws_voucher.write_number_with_format(current_row, 25, 1.0, &fmt_text).map_err(|e| e.to_string())?;

        current_row += 1;
        entry_seq += 1;
    }

    // 设置自适应列宽
    for c in 0..26 {
        ws_voucher.set_column_width(c, 13).map_err(|e| e.to_string())?;
    }
    ws_voucher.set_column_width(0, 14).map_err(|e| e.to_string())?;
    ws_voucher.set_column_width(9, 15).map_err(|e| e.to_string())?;
    ws_voucher.set_column_width(15, 14).map_err(|e| e.to_string())?;
    ws_voucher.set_column_width(23, 15).map_err(|e| e.to_string())?;

    // 复制原模板中的其他附表（例如“辅助核算”、“科目代码”等工作表）的单元格数据。
    // 凭证主表由程序重新生成；rust_xlsxwriter 无法保留原工作簿的全部样式/合并/验证元数据。
    for (s_name, s_rows) in tmpl_sheets {
        if s_name.contains("凭证") || s_name.contains("模版") {
            continue;
        }
        let ws_extra = wb_out.add_worksheet();
        ws_extra.set_name(&s_name).map_err(|e| e.to_string())?;
        for (r_idx, row) in s_rows.iter().enumerate() {
            for (c_idx, cell) in row.iter().enumerate() {
                match cell {
                    calamine::Data::Float(f) => {
                        ws_extra
                            .write_number(r_idx as u32, c_idx as u16, *f)
                            .map_err(|e| e.to_string())?;
                    }
                    calamine::Data::Int(i) => {
                        ws_extra
                            .write_number(r_idx as u32, c_idx as u16, *i as f64)
                            .map_err(|e| e.to_string())?;
                    }
                    calamine::Data::String(s) => {
                        ws_extra
                            .write_string(r_idx as u32, c_idx as u16, s)
                            .map_err(|e| e.to_string())?;
                    }
                    calamine::Data::Bool(b) => {
                        ws_extra
                            .write_boolean(r_idx as u32, c_idx as u16, *b)
                            .map_err(|e| e.to_string())?;
                    }
                    calamine::Data::DateTime(_)
                    | calamine::Data::DateTimeIso(_)
                    | calamine::Data::DurationIso(_)
                    | calamine::Data::Error(_) => {
                        let value = cell_as_string(cell);
                        if !value.is_empty() {
                            ws_extra
                                .write_string(r_idx as u32, c_idx as u16, &value)
                                .map_err(|e| e.to_string())?;
                        }
                    }
                    calamine::Data::Empty => {}
                }
            }
        }
    }

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

    let total_items = (entry_seq - 2) as usize;
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
        total_qty: (total_qty * 100.0).round() / 100.0,
        total_credit_amt: (total_credit_amt * 100.0).round() / 100.0,
        unmatched_items,
        error: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn ledger_entry(drug_name: &str, spec: &str) -> LedgerEntry {
        LedgerEntry {
            code: "1201_TEST".to_string(),
            aux_code: "TEST".to_string(),
            name_full: format!("存货_{} {}", drug_name, spec),
            drug_name: drug_name.to_string(),
            spec: spec.to_string(),
            price: 1.0,
            end_qty: 1.0,
        }
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
    fn real_outbound_workbook_still_generates_successfully() {
        let (config, _) = crate::core::config::load_config(None);
        let output = std::env::temp_dir().join(format!(
            "desktop_app_outbound_test_{}.xlsx",
            std::process::id()
        ));
        let output_string = output.to_string_lossy().to_string();

        let result = generate_outbound_voucher(
            "../../2026.8月西药销售表_已汇总.xlsx",
            "../../石家庄心理医院_数量金额总账_20260907173619.xlsx",
            "../../凭证导入模板.xlsx",
            Some(&output_string),
            None,
            true,
            &config,
        )
        .expect("真实销售样例应能生成凭证");

        assert_eq!(result.voucher_date, "2026-08-31");
        assert!(output.exists());
        std::fs::remove_file(output).expect("应清理出库测试输出");
    }
}
