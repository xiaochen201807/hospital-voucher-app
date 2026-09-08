use crate::core::excel_utils::{
    cell_as_string, find_header_row, is_summary_row, optional_col, read_sheet_rows, required_col,
    row_as_f64, row_as_string,
};
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

    // 默认读取第 0 个原始 Sheet。
    let (sheet0_name, rows0) = read_sheet_rows(in_p, &[], "销售明细")?;
    if rows0.len() < 2 {
        return Err("源数据行数过少，无有效明细数据".into());
    }

    let header_row_idx =
        find_header_row(&rows0, &["药品名称", "品名", "商品名称"], 10, "销售明细")?;
    let header = &rows0[header_row_idx];

    let col_name = required_col(header, &["药品名称", "品名", "商品名称"], "药品名称")?;
    let col_spec = optional_col(header, &["规格"], col_name + 1);
    let col_dosage = optional_col(header, &["剂型"], col_spec + 1);
    let col_factory = optional_col(header, &["制药厂", "生产厂家", "厂家"], col_dosage + 1);
    let col_unit = optional_col(header, &["单位"], col_factory + 1);
    let col_qty = optional_col(
        header,
        &["数量", "实发数量", "发药数量", "销售数量"],
        col_unit + 1,
    );
    let col_cost_amt = optional_col(header, &["进价金额", "成本金额"], col_qty + 1);
    let col_retail_amt = optional_col(
        header,
        &["零价金额", "零售金额", "售价金额"],
        col_cost_amt + 1,
    );
    let col_stock = optional_col(header, &["库存"], col_retail_amt + 1);

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
        let raw_name = row_as_string(row, col_name);
        // 捕捉原表可能已有的合计行
        if !raw_name.is_empty() && is_summary_row(&raw_name) {
            source_total_qty = Some(row_as_f64(row, col_qty));
            source_total_cost = Some(row_as_f64(row, col_cost_amt));
            source_total_retail = Some(row_as_f64(row, col_retail_amt));
            continue;
        }
        if raw_name.is_empty() {
            continue;
        }

        original_count += 1;
        let name = raw_name;
        let spec = row_as_string(row, col_spec);
        let dosage = row_as_string(row, col_dosage);
        let factory = row_as_string(row, col_factory);
        let unit = row_as_string(row, col_unit);

        let qty = row_as_f64(row, col_qty);
        let cost_amt = row_as_f64(row, col_cost_amt);
        let retail_amt = row_as_f64(row, col_retail_amt);
        let stock = row_as_f64(row, col_stock);

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
            let stem = in_p
                .file_stem()
                .and_then(|s| s.to_str())
                .unwrap_or("销售表");
            parent.join(format!("{}_已汇总.xlsx", stem))
        }
    } else {
        let parent = in_p.parent().unwrap_or_else(|| Path::new("."));
        let stem = in_p
            .file_stem()
            .and_then(|s| s.to_str())
            .unwrap_or("销售表");
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
    ws_summary
        .set_name(&target_sheet_name)
        .map_err(|e| e.to_string())?;

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

    let fmt_cell_price = Format::new()
        .set_font_size(10)
        .set_num_format("0.0000")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_qty = Format::new()
        .set_font_size(10)
        .set_num_format("#,##0")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    // 1. 标题行 (Row 0: 合并 0..9 列)
    let title_text = if sheet0_name.contains("销售表") {
        sheet0_name.replace("销售表", "销售明细表")
    } else {
        format!("{}销售明细表", sheet0_name)
    };

    ws_summary
        .set_row_height(0, 32)
        .map_err(|e| e.to_string())?;
    ws_summary
        .merge_range(0, 0, 0, 9, &title_text, &fmt_title)
        .map_err(|e| e.to_string())?;

    // 2. 表头行 (Row 1)
    let target_headers = [
        "药品名称",
        "规格",
        "剂型",
        "制药厂",
        "单位",
        "数量",
        "进价",
        "进价金额",
        "零价金额",
        "库存",
    ];

    ws_summary
        .set_row_height(1, 24)
        .map_err(|e| e.to_string())?;
    for (col_idx, h) in target_headers.iter().enumerate() {
        let fmt = if col_idx == 0 {
            &fmt_header_left
        } else {
            &fmt_header_center
        };
        ws_summary
            .write_string_with_format(1, col_idx as u16, *h, fmt)
            .map_err(|e| e.to_string())?;
    }

    // 3. 数据行 (Row 2 .. N+1)
    let mut current_row: u32 = 2;
    for item in &sorted_records {
        ws_summary
            .set_row_height(current_row, 20)
            .map_err(|e| e.to_string())?;

        ws_summary
            .write_string_with_format(current_row, 0, &item.name, &fmt_cell_left)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_string_with_format(current_row, 1, &item.spec, &fmt_cell_left)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_string_with_format(current_row, 2, &item.dosage, &fmt_cell_center)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_string_with_format(current_row, 3, &item.factory, &fmt_cell_left)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_string_with_format(current_row, 4, &item.unit, &fmt_cell_center)
            .map_err(|e| e.to_string())?;

        ws_summary
            .write_number_with_format(current_row, 5, item.qty, &fmt_cell_qty)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_number_with_format(
                current_row,
                6,
                if item.qty.abs() > 1e-9 {
                    item.cost_amt / item.qty
                } else {
                    0.0
                },
                &fmt_cell_price,
            )
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_number_with_format(current_row, 7, item.cost_amt, &fmt_cell_money)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_number_with_format(current_row, 8, item.retail_amt, &fmt_cell_money)
            .map_err(|e| e.to_string())?;
        ws_summary
            .write_number_with_format(current_row, 9, item.stock, &fmt_cell_qty)
            .map_err(|e| e.to_string())?;

        current_row += 1;
    }

    // 4. 合计行 (Row N+2)
    ws_summary
        .set_row_height(current_row, 22)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_string_with_format(current_row, 0, "合计", &fmt_header_left)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_blank(current_row, 1, &fmt_cell_left)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_blank(current_row, 2, &fmt_cell_center)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_number_with_format(
            current_row,
            3,
            totals.original_count as f64,
            &fmt_cell_center,
        )
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_string_with_format(current_row, 4, "条", &fmt_cell_center)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_number_with_format(current_row, 5, totals.total_qty, &fmt_cell_qty)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_blank(current_row, 6, &fmt_cell_price)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_number_with_format(current_row, 7, totals.total_in_amt, &fmt_cell_money)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_number_with_format(current_row, 8, totals.total_retail_amt, &fmt_cell_money)
        .map_err(|e| e.to_string())?;
    ws_summary
        .write_blank(current_row, 9, &fmt_cell_center)
        .map_err(|e| e.to_string())?;

    // 列宽设置
    ws_summary
        .set_column_width(0, 26)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(1, 16)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(2, 10)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(3, 26)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(4, 8)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(5, 12)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(6, 12)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(7, 16)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(8, 16)
        .map_err(|e| e.to_string())?;
    ws_summary
        .set_column_width(9, 12)
        .map_err(|e| e.to_string())?;

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
