use crate::core::config::{clean_text, ConfigData};
use crate::core::excel_utils::{cell_as_f64, cell_as_string, find_col_idx, open_excel};
use crate::core::outbound::{load_ledger_entries, match_drug, UnmatchedDrug};
use calamine::Reader;
use rust_xlsxwriter::{Format, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::path::{Path, PathBuf};

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
    #[allow(dead_code)]
    unit: String,
}

/// 提取供应商简称摘要，例如 "河北蕴德药业有限公司" -> "河北蕴德到货"
pub fn get_supplier_brief(sup_name: &str) -> String {
    if sup_name.is_empty() {
        return "药品到货".to_string();
    }
    let cleaned = sup_name
        .replace("股份有限公司", "")
        .replace("医药有限公司", "")
        .replace("药业有限公司", "")
        .replace("有限责任公司", "")
        .replace("有限公司", "")
        .replace("责任公司", "");
    format!("{}到货", cleaned.trim())
}

/// 从凭证模板中提取供应商往来字典 (code -> name)
pub fn load_supplier_dict(template_path: &Path) -> HashMap<String, String> {
    let mut dict = HashMap::new();
    if let Ok(mut wb) = open_excel(template_path) {
        if let Some(sheet_name) = wb.sheet_names().iter().find(|s| s.contains("辅助核算")) {
            if let Ok(range) = wb.worksheet_range(sheet_name) {
                for row in range.rows() {
                    if row.len() >= 3 {
                        let cat = cell_as_string(&row[0]);
                        if cat.contains("供应商") {
                            let code = cell_as_string(&row[1]);
                            let name = cell_as_string(&row[2]);
                            if !code.is_empty() && !name.is_empty() {
                                dict.insert(code, name);
                            }
                        }
                    }
                }
            }
        }
    }

    // 内置常见兜底
    if dict.is_empty() {
        dict.insert("001".into(), "国药乐仁堂医药有限公司".into());
        dict.insert("002".into(), "华润益生制药有限公司".into());
        dict.insert("004".into(), "河北国泰医药有限责任公司".into());
        dict.insert("005".into(), "河北蕴德药业有限公司".into());
        dict.insert("007".into(), "石药集团中诚医药字号".into());
    }

    dict
}

/// 匹配供应商编码
pub fn match_supplier_code(sup_name: &str, dict: &HashMap<String, String>) -> (String, String) {
    let clean_sup = clean_text(sup_name);

    // 1. 完全一致匹配
    for (code, name) in dict {
        if clean_text(name) == clean_sup {
            return (code.clone(), name.clone());
        }
    }

    // 2. 互相包含比对
    for (code, name) in dict {
        let n_clean = clean_text(name);
        if n_clean.contains(&clean_sup) || clean_sup.contains(&n_clean) {
            return (code.clone(), name.clone());
        }
    }

    // 3. 去除通用后缀比对
    let core_sup = sup_name
        .replace("股份有限公司", "")
        .replace("医药有限公司", "")
        .replace("药业有限公司", "")
        .replace("有限责任公司", "")
        .replace("有限公司", "");
    let core_clean = clean_text(&core_sup);

    if !core_clean.is_empty() {
        for (code, name) in dict {
            if clean_text(name).contains(&core_clean) {
                return (code.clone(), name.clone());
            }
        }
    }

    (String::new(), sup_name.to_string())
}

/// 执行生成药房入库凭证
pub fn generate_inbound_voucher(
    inbound_path: &str,
    ledger_path: &str,
    template_path: &str,
    custom_output: Option<&str>,
    target_date_opt: Option<&str>,
    voucher_no: Option<&str>,
    config: &ConfigData,
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

    // 2. 读取供应商字典与模板所有 Sheet
    let supplier_dict = load_supplier_dict(tmpl_p);
    let mut tmpl_sheets = Vec::new();
    let mut tmpl_wb = open_excel(tmpl_p)?;
    for s in tmpl_wb.sheet_names() {
        let range = tmpl_wb
            .worksheet_range(&s)
            .map_err(|e| format!("读取凭证模板工作表 '{}' 失败: {}", s, e))?;
        let r_rows: Vec<Vec<calamine::Data>> = range.rows().map(|r| r.to_vec()).collect();
        tmpl_sheets.push((s.to_string(), r_rows));
    }

    // 3. 读取入库单明细
    let mut in_wb = open_excel(in_p)?;
    let sheet_names = in_wb.sheet_names();
    let s_name = sheet_names
        .iter()
        .find(|s| s.contains("入库") || s.contains("明细"))
        .unwrap_or(&sheet_names[0]);

    let range = in_wb
        .worksheet_range(s_name)
        .map_err(|e| format!("读取入库单失败: {}", e))?;

    let rows: Vec<Vec<calamine::Data>> = range.rows().map(|r| r.to_vec()).collect();
    if rows.len() < 2 {
        return Err("入库单中无有效数据".into());
    }

    let mut h_idx = 0;
    for (i, r) in rows.iter().take(10).enumerate() {
        if find_col_idx(r, &["药品名称", "品名", "商品名称"]).is_some() {
            h_idx = i;
            break;
        }
    }

    let header = &rows[h_idx];
    let col_name = find_col_idx(header, &["药品名称", "品名", "商品名称"]).unwrap();
    let col_spec = find_col_idx(header, &["规格"]).unwrap_or(col_name + 1);
    let col_factory = find_col_idx(header, &["制药厂", "生产厂家", "厂家"]).unwrap_or(col_spec + 1);
    let col_sup = find_col_idx(header, &["供货单位", "供应商", "单位名称"]).unwrap_or(col_factory + 1);
    let col_unit = find_col_idx(header, &["单位"]).unwrap_or(col_sup + 1);
    let col_qty = find_col_idx(header, &["数量", "入库数量"]).unwrap_or(col_unit + 1);
    let col_price = find_col_idx(header, &["进价", "成本单价", "单价"]).unwrap_or(col_qty + 1);
    let col_amt = find_col_idx(header, &["金额", "进价金额", "入库金额"]).unwrap_or(col_price + 1);
    let col_date = find_col_idx(header, &["入库日期", "日期"]);

    let is_tcm = inbound_path.contains("中药");
    let subject_code_debit = "1201";
    let subject_code_credit = "220201";

    let v_no_str = voucher_no.unwrap_or("").trim();

    let mut items = Vec::new();
    let mut source_dates = Vec::new();
    let mut total_debit = 0.0;
    let mut total_qty = 0.0;

    for row in rows.iter().skip(h_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let name = cell_as_string(row.get(col_name).unwrap_or(&calamine::Data::Empty));
        if name.is_empty() || name.contains("合计") || name.contains("总计") {
            continue;
        }

        if let Some(date_col) = col_date {
            if let Some(date_cell) = row.get(date_col) {
                let date_value = cell_as_string(date_cell);
                if !date_value.is_empty() {
                    source_dates.push(date_value);
                }
            }
        }

        let spec = row.get(col_spec).map(cell_as_string).unwrap_or_default();
        let factory = row.get(col_factory).map(cell_as_string).unwrap_or_default();
        let supplier = row.get(col_sup).map(cell_as_string).unwrap_or_default();
        let unit = row.get(col_unit).map(cell_as_string).unwrap_or_default();

        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let price = row.get(col_price).map(cell_as_f64).unwrap_or(0.0);
        let mut amt = row.get(col_amt).map(cell_as_f64).unwrap_or(0.0);
        if amt == 0.0 && qty > 0.0 && price > 0.0 {
            amt = (qty * price * 100.0).round() / 100.0;
        }

        total_debit += amt;
        total_qty += qty;

        items.push(InboundItem {
            supplier,
            name,
            spec,
            factory,
            qty,
            price,
            amount: amt,
            unit,
        });
    }

    let voucher_date = crate::core::config::detect_voucher_date_with_source_dates(
        target_date_opt,
        &[inbound_path, ledger_path],
        &source_dates,
    );

    // 按入库单中首次出现的顺序建立供应商分组，保证每个供应商的明细紧跟其汇总分录。
    let mut supplier_groups: Vec<(String, Vec<usize>, f64)> = Vec::new();
    let mut supplier_group_indices = HashMap::new();
    for (item_idx, it) in items.iter().enumerate() {
        let group_idx = if let Some(group_idx) = supplier_group_indices.get(&it.supplier) {
            *group_idx
        } else {
            let group_idx = supplier_groups.len();
            supplier_groups.push((it.supplier.clone(), Vec::new(), 0.0));
            supplier_group_indices.insert(it.supplier.clone(), group_idx);
            group_idx
        };
        supplier_groups[group_idx].1.push(item_idx);
        supplier_groups[group_idx].2 += it.amount;
    }

    let mut suppliers_summary = Vec::new();
    let mut total_credit = 0.0;

    for (sup, item_indices, amt) in &supplier_groups {
        let (code, _) = match_supplier_code(sup, &supplier_dict);
        let brief = get_supplier_brief(sup);
        let c_amt = (*amt * 100.0).round() / 100.0;
        total_credit += c_amt;

        suppliers_summary.push(SupplierSummary {
            supplier: sup.clone(),
            code,
            brief,
            count: item_indices.len(),
            amount: c_amt,
        });
    }

    total_debit = (total_debit * 100.0).round() / 100.0;
    total_credit = (total_credit * 100.0).round() / 100.0;
    let is_balanced = (total_debit - total_credit).abs() < 0.01;

    // 4. 生成凭证 Excel
    let mut out_wb = Workbook::new();
    let ws = out_wb.add_worksheet();
    ws.set_name("凭证模版").map_err(|e| e.to_string())?;

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

    let fmt_left = Format::new()
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
        .set_num_format("#,##0.00")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let template_headers = [
        "日期", "凭证字", "凭证号", "附件数", "分录序号", "摘要", "科目代码", "科目名称",
        "借方金额", "贷方金额", "客户", "供应商", "职员", "项目", "部门", "存货",
        "是否限定", "自定义类别", "自定义编码", "自定义类别1", "自定义编码1", "数量",
        "单价", "原币金额", "币别", "汇率",
    ];

    ws.set_row_height(0, 26).map_err(|e| e.to_string())?;
    for (c, h) in template_headers.iter().enumerate() {
        ws.write_string_with_format(0, c as u16, *h, &fmt_header)
            .map_err(|e| e.to_string())?;
    }

    let mut row_idx: u32 = 1;
    let mut seq = 1;
    let mut matched_count = 0;
    let mut unmatched_items = Vec::new();

    // 写入“供应商明细借方分录 -> 该供应商贷方汇总分录”，再处理下一个供应商。
    for (group_idx, (_, item_indices, _)) in supplier_groups.iter().enumerate() {
        for item_idx in item_indices {
            let it = &items[*item_idx];
            let matched = match_drug(&it.name, &it.spec, &it.factory, &ledger_entries, config);
            let aux_code = if let Some(m) = matched {
                matched_count += 1;
                m.aux_code
            } else {
                let reason = if config.strict_vendor_suffix_drugs.iter().any(|d| clean_text(d) == clean_text(&it.name)) {
                    "严格锁定厂家药品，总账未查到匹配厂家科目".to_string()
                } else {
                    "总账中未查到对应存货编码".to_string()
                };
                unmatched_items.push(UnmatchedDrug {
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
                });
                String::new()
            };

            ws.set_row_height(row_idx, 20).map_err(|e| e.to_string())?;

            ws.write_string_with_format(row_idx, 0, &voucher_date, &fmt_text).map_err(|e| e.to_string())?;
            ws.write_string_with_format(row_idx, 1, "记", &fmt_text).map_err(|e| e.to_string())?;
            if !v_no_str.is_empty() {
                ws.write_string_with_format(row_idx, 2, v_no_str, &fmt_text).map_err(|e| e.to_string())?;
            } else {
                ws.write_blank(row_idx, 2, &fmt_text).map_err(|e| e.to_string())?;
            }
            ws.write_blank(row_idx, 3, &fmt_text).map_err(|e| e.to_string())?;
            ws.write_number_with_format(row_idx, 4, seq as f64, &fmt_text).map_err(|e| e.to_string())?;

            let sup_brief = get_supplier_brief(&it.supplier);
            let summary_text = if !sup_brief.is_empty() {
                sup_brief
            } else {
                format!("入库-{}", it.name)
            };
            ws.write_string_with_format(row_idx, 5, &summary_text, &fmt_left).map_err(|e| e.to_string())?;
            ws.write_string_with_format(row_idx, 6, subject_code_debit, &fmt_text).map_err(|e| e.to_string())?;
            ws.write_blank(row_idx, 7, &fmt_left).map_err(|e| e.to_string())?;
            ws.write_number_with_format(row_idx, 8, it.amount, &fmt_money).map_err(|e| e.to_string())?;
            ws.write_blank(row_idx, 9, &fmt_text).map_err(|e| e.to_string())?;

            for c in 10..15 {
                ws.write_blank(row_idx, c as u16, &fmt_text).map_err(|e| e.to_string())?;
            }

            if !aux_code.is_empty() {
                ws.write_string_with_format(row_idx, 15, &aux_code, &fmt_text).map_err(|e| e.to_string())?;
            } else {
                ws.write_blank(row_idx, 15, &fmt_text).map_err(|e| e.to_string())?;
            }

            for c in 16..21 {
                ws.write_blank(row_idx, c as u16, &fmt_text).map_err(|e| e.to_string())?;
            }

            ws.write_number_with_format(row_idx, 21, it.qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws.write_number_with_format(row_idx, 22, it.price, &fmt_money).map_err(|e| e.to_string())?;
            ws.write_number_with_format(row_idx, 23, it.amount, &fmt_money).map_err(|e| e.to_string())?;
            ws.write_string_with_format(row_idx, 24, "RMB", &fmt_text).map_err(|e| e.to_string())?;
            ws.write_number_with_format(row_idx, 25, 1.0, &fmt_text).map_err(|e| e.to_string())?;

            row_idx += 1;
            seq += 1;
        }

        // 写入当前供应商的贷方汇总分录。
        let sup = &suppliers_summary[group_idx];
        ws.set_row_height(row_idx, 20).map_err(|e| e.to_string())?;

        ws.write_string_with_format(row_idx, 0, &voucher_date, &fmt_text).map_err(|e| e.to_string())?;
        ws.write_string_with_format(row_idx, 1, "记", &fmt_text).map_err(|e| e.to_string())?;
        if !v_no_str.is_empty() {
            ws.write_string_with_format(row_idx, 2, v_no_str, &fmt_text).map_err(|e| e.to_string())?;
        } else {
            ws.write_blank(row_idx, 2, &fmt_text).map_err(|e| e.to_string())?;
        }
        ws.write_blank(row_idx, 3, &fmt_text).map_err(|e| e.to_string())?;
        ws.write_number_with_format(row_idx, 4, seq as f64, &fmt_text).map_err(|e| e.to_string())?;

        let sup_summary = sup.brief.clone();
        ws.write_string_with_format(row_idx, 5, &sup_summary, &fmt_left).map_err(|e| e.to_string())?;
        ws.write_string_with_format(row_idx, 6, subject_code_credit, &fmt_text).map_err(|e| e.to_string())?;
        ws.write_blank(row_idx, 7, &fmt_left).map_err(|e| e.to_string())?;
        ws.write_blank(row_idx, 8, &fmt_text).map_err(|e| e.to_string())?;
        ws.write_number_with_format(row_idx, 9, sup.amount, &fmt_money).map_err(|e| e.to_string())?;

        ws.write_blank(row_idx, 10, &fmt_text).map_err(|e| e.to_string())?;

        // 供应商辅助核算代码
        if !sup.code.is_empty() {
            ws.write_string_with_format(row_idx, 11, &sup.code, &fmt_text).map_err(|e| e.to_string())?;
        } else {
            ws.write_blank(row_idx, 11, &fmt_text).map_err(|e| e.to_string())?;
        }

        for c in 12..24 {
            ws.write_blank(row_idx, c as u16, &fmt_text).map_err(|e| e.to_string())?;
        }

        ws.write_string_with_format(row_idx, 24, "RMB", &fmt_text).map_err(|e| e.to_string())?;
        ws.write_number_with_format(row_idx, 25, 1.0, &fmt_text).map_err(|e| e.to_string())?;

        row_idx += 1;
        seq += 1;
    }

    // 设置列宽
    for c in 0..26 {
        ws.set_column_width(c, 13).map_err(|e| e.to_string())?;
    }
    ws.set_column_width(0, 14).map_err(|e| e.to_string())?;
    ws.set_column_width(5, 20).map_err(|e| e.to_string())?;
    ws.set_column_width(7, 16).map_err(|e| e.to_string())?;
    ws.set_column_width(8, 15).map_err(|e| e.to_string())?;
    ws.set_column_width(9, 15).map_err(|e| e.to_string())?;
    ws.set_column_width(11, 14).map_err(|e| e.to_string())?;
    ws.set_column_width(15, 14).map_err(|e| e.to_string())?;

    // 复制原模板中的其他附表（例如“辅助核算”、“科目代码”等工作表）的单元格数据。
    // 凭证主表由程序重新生成；rust_xlsxwriter 无法保留原工作簿的全部样式/合并/验证元数据。
    for (s_name, s_rows) in tmpl_sheets {
        if s_name.contains("凭证") || s_name.contains("模版") {
            continue;
        }
        let ws_extra = out_wb.add_worksheet();
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

    let out_path_buf = if let Some(co) = custom_output {
        PathBuf::from(co)
    } else {
        let parent = in_p.parent().unwrap_or_else(|| Path::new("."));
        let name_prefix = if is_tcm { "凭证导入模板_中药入库_已生成.xlsx" } else { "凭证导入模板_入库_已生成.xlsx" };
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
        )
        .expect("真实入库样例应能生成凭证");

        assert_eq!(result.voucher_date, "2026-08-31");

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
                        assert_eq!(current_supplier.as_deref(), Some(summary.as_str()), "供应商汇总必须紧跟明细");
                        current_supplier = None;
                    }
                    _ => {}
                }
            }
        }
        assert!(current_supplier.is_none(), "每个供应商的明细都必须有对应汇总");

        std::fs::remove_file(output).expect("应清理入库测试输出");
    }
}
