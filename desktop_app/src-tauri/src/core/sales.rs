use crate::core::excel_utils::{cell_as_f64, cell_as_string, find_col_idx, open_excel};
use calamine::Reader;
use rust_xlsxwriter::{Format, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct SalesTotals {
    pub original_count: usize,
    pub unique_count: usize,
    pub total_qty: f64,
    pub total_in_amt: f64,
    pub total_retail_amt: f64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct SalesProcessResponse {
    pub success: bool,
    pub output_file: String,
    pub target_sheet: String,
    pub totals: SalesTotals,
    #[serde(default)]
    pub error: Option<String>,
}

#[derive(Debug, Clone)]
pub struct DrugSaleItem {
    pub name: String,
    pub spec: String,
    pub dosage: String,
    pub factory: String,
    pub unit: String,
    pub qty: f64,
    pub cost_amt: f64,
    pub retail_amt: f64,
    pub stock: f64,
}

pub fn process_sales_file(
    input_path: &str,
    custom_output: Option<&str>,
    sheet_name_hint: Option<&str>,
) -> Result<SalesProcessResponse, String> {
    let in_p = Path::new(input_path);
    if !in_p.exists() {
        return Err(format!("销售明细文件 '{:?}' 不存在", in_p));
    }

    let mut workbook = open_excel(in_p)?;
    let sheet_names = workbook.sheet_names();
    if sheet_names.is_empty() {
        return Err("Excel 文件中未发现任何工作表".into());
    }

    // 默认读取第 0 个原始 Sheet
    let sheet0_name = sheet_names[0].clone();
    let range0 = workbook
        .worksheet_range(&sheet0_name)
        .map_err(|e| format!("读取工作表 '{}' 失败: {}", sheet0_name, e))?;

    let rows0: Vec<Vec<calamine::Data>> = range0.rows().map(|r| r.to_vec()).collect();
    if rows0.len() < 2 {
        return Err("源数据行数过少，无有效明细数据".into());
    }

    // 寻找表头所在行（包含“药品名称”或“品名”）
    let mut header_idx = None;
    for (i, row) in rows0.iter().take(10).enumerate() {
        if find_col_idx(row, &["药品名称", "品名", "商品名称"]).is_some() {
            header_idx = Some(i);
            break;
        }
    }

    let header_row_idx = header_idx.ok_or_else(|| "未能在前10行中找到包含'药品名称'的有效表头".to_string())?;
    let header = &rows0[header_row_idx];

    let col_name = find_col_idx(header, &["药品名称", "品名", "商品名称"]).unwrap();
    let col_spec = find_col_idx(header, &["规格"]).unwrap_or(col_name + 1);
    let col_dosage = find_col_idx(header, &["剂型"]).unwrap_or(col_spec + 1);
    let col_factory = find_col_idx(header, &["制药厂", "生产厂家", "厂家"]).unwrap_or(col_dosage + 1);
    let col_unit = find_col_idx(header, &["单位"]).unwrap_or(col_factory + 1);
    let col_qty = find_col_idx(header, &["数量", "实发数量", "发药数量", "销售数量"]).unwrap_or(col_unit + 1);
    let col_cost_amt = find_col_idx(header, &["进价金额", "成本金额"]).unwrap_or(col_qty + 1);
    let col_retail_amt = find_col_idx(header, &["零价金额", "零售金额", "售价金额"]).unwrap_or(col_cost_amt + 1);
    let col_stock = find_col_idx(header, &["库存"]).unwrap_or(col_retail_amt + 1);

    // 以 (name, spec, dosage, factory) 4项元组为精确去重分组 Key
    let mut groups: BTreeMap<(String, String, String, String), DrugSaleItem> = BTreeMap::new();
    let mut original_count = 0;
    let mut source_total_qty: Option<f64> = None;
    let mut source_total_cost: Option<f64> = None;
    let mut source_total_retail: Option<f64> = None;

    for row in rows0.iter().skip(header_row_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let raw_name = cell_as_string(row.get(col_name).unwrap_or(&calamine::Data::Empty));
        if raw_name.is_empty() {
            continue;
        }

        // 捕捉原表可能已有的合计行
        if raw_name.contains("合计") || raw_name.contains("总计") {
            if let Some(c_qty) = row.get(col_qty) {
                source_total_qty = Some(cell_as_f64(c_qty));
            }
            if let Some(c_cost) = row.get(col_cost_amt) {
                source_total_cost = Some(cell_as_f64(c_cost));
            }
            if let Some(c_ret) = row.get(col_retail_amt) {
                source_total_retail = Some(cell_as_f64(c_ret));
            }
            continue;
        }

        original_count += 1;
        let name = raw_name;
        let spec = row.get(col_spec).map(cell_as_string).unwrap_or_default();
        let dosage = row.get(col_dosage).map(cell_as_string).unwrap_or_default();
        let factory = row.get(col_factory).map(cell_as_string).unwrap_or_default();
        let unit = row.get(col_unit).map(cell_as_string).unwrap_or_default();

        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let cost_amt = row.get(col_cost_amt).map(cell_as_f64).unwrap_or(0.0);
        let retail_amt = row.get(col_retail_amt).map(cell_as_f64).unwrap_or(0.0);
        let stock = row.get(col_stock).map(cell_as_f64).unwrap_or(0.0);

        let key = (name.clone(), spec.clone(), dosage.clone(), factory.clone());
        groups
            .entry(key)
            .and_modify(|e| {
                e.qty += qty;
                e.cost_amt += cost_amt;
                e.retail_amt += retail_amt;
                if e.stock == 0.0 && stock > 0.0 {
                    e.stock = stock;
                }
            })
            .or_insert(DrugSaleItem {
                name,
                spec,
                dosage,
                factory,
                unit,
                qty,
                cost_amt,
                retail_amt,
                stock,
            });
    }

    // 排序规则：按数量降序，数量相同按名称排序
    let mut sorted_records: Vec<DrugSaleItem> = groups.into_values().collect();
    sorted_records.sort_by(|a, b| {
        b.qty
            .partial_cmp(&a.qty)
            .unwrap_or(std::cmp::Ordering::Equal)
            .then_with(|| a.name.cmp(&b.name))
    });

    let sum_qty: f64 = sorted_records.iter().map(|i| i.qty).sum();
    let sum_cost: f64 = sorted_records.iter().map(|i| i.cost_amt).sum();
    let sum_retail: f64 = sorted_records.iter().map(|i| i.retail_amt).sum();

    let final_qty = source_total_qty.unwrap_or(sum_qty);
    let final_cost = source_total_cost.unwrap_or(sum_cost);
    let final_retail = source_total_retail.unwrap_or(sum_retail);

    let totals = SalesTotals {
        original_count,
        unique_count: sorted_records.len(),
        total_qty: (final_qty * 100.0).round() / 100.0,
        total_in_amt: (final_cost * 100.0).round() / 100.0,
        total_retail_amt: (final_retail * 100.0).round() / 100.0,
    };

    // 输出目标 Sheet 名称（默认为“销售明细”）
    let target_sheet_name = if let Some(h) = sheet_name_hint {
        if !h.trim().is_empty() {
            h.to_string()
        } else {
            "销售明细".to_string()
        }
    } else {
        "销售明细".to_string()
    };

    // 输出目标文件路径
    let out_path_buf = if let Some(co) = custom_output {
        if !co.trim().is_empty() {
            PathBuf::from(co)
        } else {
            let parent = in_p.parent().unwrap_or_else(|| Path::new("."));
            let stem = in_p.file_stem().and_then(|s| s.to_str()).unwrap_or("销售表");
            parent.join(format!("{}_已汇总.xlsx", stem))
        }
    } else {
        let parent = in_p.parent().unwrap_or_else(|| Path::new("."));
        let stem = in_p.file_stem().and_then(|s| s.to_str()).unwrap_or("销售表");
        parent.join(format!("{}_已汇总.xlsx", stem))
    };

    // 创建纯 Rust 工作簿
    let mut out_wb = Workbook::new();

    // ----------------------------------------------------
    // Sheet 1: 保留原始数据 Sheet0
    // ----------------------------------------------------
    let ws_raw = out_wb.add_worksheet();
    ws_raw.set_name(&sheet0_name).map_err(|e| e.to_string())?;
    for (r_idx, row) in rows0.iter().enumerate() {
        for (c_idx, cell) in row.iter().enumerate() {
            let s = cell_as_string(cell);
            if let Ok(num) = s.parse::<f64>() {
                let _ = ws_raw.write_number(r_idx as u32, c_idx as u16, num);
            } else {
                let _ = ws_raw.write_string(r_idx as u32, c_idx as u16, &s);
            }
        }
    }

    // ----------------------------------------------------
    // Sheet 2: 汇总去重后的【销售明细】Sheet
    // ----------------------------------------------------
    let ws_summary = out_wb.add_worksheet();
    ws_summary.set_name(&target_sheet_name).map_err(|e| e.to_string())?;

    let fmt_title = Format::new()
        .set_bold()
        .set_font_size(18)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter);

    let fmt_header_left = Format::new()
        .set_bold()
        .set_font_size(11)
        .set_align(rust_xlsxwriter::FormatAlign::Left)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_header_center = Format::new()
        .set_bold()
        .set_font_size(11)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_left = Format::new()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Left)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_center = Format::new()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_money = Format::new()
        .set_font_size(10)
        .set_num_format("0.00_")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_qty = Format::new()
        .set_font_size(10)
        .set_num_format("#,##0")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    // 1. 标题行 (Row 0: 合并 0..8 列)
    let title_text = if sheet0_name.contains("销售表") {
        sheet0_name.replace("销售表", "销售明细表")
    } else {
        format!("{}销售明细表", sheet0_name)
    };

    ws_summary.set_row_height(0, 32).map_err(|e| e.to_string())?;
    ws_summary
        .merge_range(0, 0, 0, 8, &title_text, &fmt_title)
        .map_err(|e| e.to_string())?;

    // 2. 表头行 (Row 1)
    let target_headers = [
        "药品名称", "规格", "剂型", "制药厂", "单位", "数量", "进价金额", "零价金额", "库存",
    ];

    ws_summary.set_row_height(1, 24).map_err(|e| e.to_string())?;
    for (col_idx, h) in target_headers.iter().enumerate() {
        let fmt = if col_idx == 0 { &fmt_header_left } else { &fmt_header_center };
        ws_summary
            .write_string_with_format(1, col_idx as u16, *h, fmt)
            .map_err(|e| e.to_string())?;
    }

    // 3. 数据行 (Row 2 .. N+1)
    let mut current_row: u32 = 2;
    for item in &sorted_records {
        ws_summary.set_row_height(current_row, 20).map_err(|e| e.to_string())?;

        ws_summary.write_string_with_format(current_row, 0, &item.name, &fmt_cell_left).map_err(|e| e.to_string())?;
        ws_summary.write_string_with_format(current_row, 1, &item.spec, &fmt_cell_left).map_err(|e| e.to_string())?;
        ws_summary.write_string_with_format(current_row, 2, &item.dosage, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_summary.write_string_with_format(current_row, 3, &item.factory, &fmt_cell_left).map_err(|e| e.to_string())?;
        ws_summary.write_string_with_format(current_row, 4, &item.unit, &fmt_cell_center).map_err(|e| e.to_string())?;

        ws_summary.write_number_with_format(current_row, 5, item.qty, &fmt_cell_qty).map_err(|e| e.to_string())?;
        ws_summary.write_number_with_format(current_row, 6, item.cost_amt, &fmt_cell_money).map_err(|e| e.to_string())?;
        ws_summary.write_number_with_format(current_row, 7, item.retail_amt, &fmt_cell_money).map_err(|e| e.to_string())?;
        ws_summary.write_number_with_format(current_row, 8, item.stock, &fmt_cell_qty).map_err(|e| e.to_string())?;

        current_row += 1;
    }

    // 4. 合计行 (Row N+2)
    ws_summary.set_row_height(current_row, 22).map_err(|e| e.to_string())?;
    ws_summary.write_string_with_format(current_row, 0, "合计", &fmt_header_left).map_err(|e| e.to_string())?;
    ws_summary.write_blank(current_row, 1, &fmt_cell_left).map_err(|e| e.to_string())?;
    ws_summary.write_blank(current_row, 2, &fmt_cell_center).map_err(|e| e.to_string())?;
    ws_summary.write_number_with_format(current_row, 3, totals.original_count as f64, &fmt_cell_center).map_err(|e| e.to_string())?;
    ws_summary.write_string_with_format(current_row, 4, "条", &fmt_cell_center).map_err(|e| e.to_string())?;
    ws_summary.write_number_with_format(current_row, 5, totals.total_qty, &fmt_cell_qty).map_err(|e| e.to_string())?;
    ws_summary.write_number_with_format(current_row, 6, totals.total_in_amt, &fmt_cell_money).map_err(|e| e.to_string())?;
    ws_summary.write_number_with_format(current_row, 7, totals.total_retail_amt, &fmt_cell_money).map_err(|e| e.to_string())?;
    ws_summary.write_blank(current_row, 8, &fmt_cell_center).map_err(|e| e.to_string())?;

    // 列宽设置
    ws_summary.set_column_width(0, 26).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(1, 16).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(2, 10).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(3, 26).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(4, 8).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(5, 12).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(6, 16).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(7, 16).map_err(|e| e.to_string())?;
    ws_summary.set_column_width(8, 12).map_err(|e| e.to_string())?;

    // 保存文件
    out_wb
        .save(&out_path_buf)
        .map_err(|e| format!("保存 Excel 文件失败: {}", e))?;

    Ok(SalesProcessResponse {
        success: true,
        output_file: out_path_buf.to_string_lossy().to_string(),
        target_sheet: target_sheet_name,
        totals,
        error: None,
    })
}
