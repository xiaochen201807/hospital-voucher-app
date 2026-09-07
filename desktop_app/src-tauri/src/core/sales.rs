use crate::core::config::clean_text;
use crate::core::excel_utils::{cell_as_f64, cell_as_string, find_col_idx, open_excel};
use calamine::Reader;
use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::BTreeMap;
use std::path::{Path, PathBuf};

#[derive(Debug, Serialize, Deserialize)]
pub struct SalesSummaryResult {
    pub success: bool,
    pub output_file: String,
    pub total_raw_rows: usize,
    pub unique_drugs_count: usize,
    pub total_qty: f64,
    pub total_cost_amt: f64,
    pub total_retail_amt: f64,
    #[serde(default)]
    pub error: Option<String>,
}

#[derive(Debug, Clone)]
struct DrugAggregate {
    name: String,
    spec: String,
    dosage: String,
    factory: String,
    unit: String,
    qty: f64,
    cost_amt: f64,
    retail_amt: f64,
    stock: f64,
}

pub fn process_sales_file(
    input_path: &str,
    custom_output: Option<&str>,
    sheet_name_hint: Option<&str>,
) -> Result<SalesSummaryResult, String> {
    let in_p = Path::new(input_path);
    if !in_p.exists() {
        return Err(format!("销售明细文件 '{:?}' 不存在", in_p));
    }

    let mut workbook = open_excel(in_p)?;
    let sheet_names = workbook.sheet_names();
    if sheet_names.is_empty() {
        return Err("Excel 文件中未发现任何工作表".into());
    }

    // 优先匹配包含“销售”或指定名字的 sheet
    let target_sheet = if let Some(hint) = sheet_name_hint {
        sheet_names
            .iter()
            .find(|s| s.contains(hint))
            .cloned()
            .unwrap_or_else(|| sheet_names[0].clone())
    } else {
        sheet_names
            .iter()
            .find(|s| s.contains("销售") || s.contains("明细"))
            .cloned()
            .unwrap_or_else(|| sheet_names[0].clone())
    };

    let range = workbook
        .worksheet_range(&target_sheet)
        .map_err(|e| format!("读取工作表 '{}' 失败: {}", target_sheet, e))?;

    let rows: Vec<Vec<calamine::Data>> = range.rows().map(|r| r.to_vec()).collect();
    if rows.len() < 2 {
        return Err("表格行数过少，无有效明细数据".into());
    }

    // 寻找表头所在行（通常在前 5 行内包含“药品名称”或“品名”）
    let mut header_idx = None;
    for (i, row) in rows.iter().take(10).enumerate() {
        if find_col_idx(row, &["药品名称", "品名", "商品名称"]).is_some() {
            header_idx = Some(i);
            break;
        }
    }

    let header_row_idx = header_idx.ok_or_else(|| "未能在前10行中找到包含'药品名称'的有效表头".to_string())?;
    let header = &rows[header_row_idx];

    let col_name = find_col_idx(header, &["药品名称", "品名", "商品名称"]).unwrap();
    let col_spec = find_col_idx(header, &["规格"]).unwrap_or(col_name + 1);
    let col_dosage = find_col_idx(header, &["剂型"]).unwrap_or(col_spec + 1);
    let col_factory = find_col_idx(header, &["制药厂", "生产厂家", "厂家"]).unwrap_or(col_dosage + 1);
    let col_unit = find_col_idx(header, &["单位"]).unwrap_or(col_factory + 1);
    let col_qty = find_col_idx(header, &["数量", "实发数量", "发药数量", "销售数量"]).unwrap_or(col_unit + 1);
    let col_cost_amt = find_col_idx(header, &["进价金额", "成本金额"]).unwrap_or(col_qty + 1);
    let col_retail_amt = find_col_idx(header, &["零价金额", "零售金额", "售价金额"]).unwrap_or(col_cost_amt + 1);
    let col_stock = find_col_idx(header, &["库存"]).unwrap_or(col_retail_amt + 1);

    let mut map: BTreeMap<String, DrugAggregate> = BTreeMap::new();
    let mut raw_count = 0;

    for row in rows.iter().skip(header_row_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let name = cell_as_string(row.get(col_name).unwrap_or(&calamine::Data::Empty));
        if name.is_empty() || name.contains("合计") || name.contains("总计") {
            continue;
        }

        let spec = row.get(col_spec).map(cell_as_string).unwrap_or_default();
        let dosage = row.get(col_dosage).map(cell_as_string).unwrap_or_default();
        let factory = row.get(col_factory).map(cell_as_string).unwrap_or_default();
        let unit = row.get(col_unit).map(cell_as_string).unwrap_or_default();

        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let cost_amt = row.get(col_cost_amt).map(cell_as_f64).unwrap_or(0.0);
        let retail_amt = row.get(col_retail_amt).map(cell_as_f64).unwrap_or(0.0);
        let stock = row.get(col_stock).map(cell_as_f64).unwrap_or(0.0);

        let group_key = format!("{}|{}", clean_text(&name), clean_text(&spec));
        raw_count += 1;

        map.entry(group_key)
            .and_modify(|e| {
                e.qty += qty;
                e.cost_amt += cost_amt;
                e.retail_amt += retail_amt;
                if e.stock == 0.0 && stock > 0.0 {
                    e.stock = stock;
                }
            })
            .or_insert(DrugAggregate {
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

    // 确定输出文件路径
    let out_path_buf = if let Some(co) = custom_output {
        PathBuf::from(co)
    } else {
        let parent = in_p.parent().unwrap_or_else(|| Path::new("."));
        let stem = in_p.file_stem().and_then(|s| s.to_str()).unwrap_or("销售明细");
        parent.join(format!("{}_已汇总.xlsx", stem))
    };

    // 使用 rust_xlsxwriter 生成输出工作簿
    let mut out_wb = Workbook::new();
    let ws = out_wb.add_worksheet();
    ws.set_name("销售汇总")
        .map_err(|e| format!("设置工作表名称失败: {}", e))?;

    // 样式设计
    let header_fmt = Format::new()
        .set_bold()
        .set_font_size(11)
        .set_background_color(Color::RGB(0x2D3748))
        .set_font_color(Color::RGB(0xFFFFFF))
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let cell_str_fmt = Format::new()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Left)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let cell_num_fmt = Format::new()
        .set_font_size(10)
        .set_num_format("#,##0.00")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let cell_int_fmt = Format::new()
        .set_font_size(10)
        .set_num_format("#,##0")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let headers = [
        "药品名称", "规格", "剂型", "制药厂", "单位", "数量", "进价金额", "零价金额", "库存",
    ];

    ws.set_row_height(0, 26).map_err(|e| e.to_string())?;
    for (c, h) in headers.iter().enumerate() {
        ws.write_string_with_format(0, c as u16, *h, &header_fmt)
            .map_err(|e| e.to_string())?;
    }

    let mut row_idx: u32 = 1;
    let mut total_qty_sum = 0.0;
    let mut total_cost_sum = 0.0;
    let mut total_retail_sum = 0.0;

    for item in map.values() {
        ws.set_row_height(row_idx, 20).map_err(|e| e.to_string())?;

        ws.write_string_with_format(row_idx, 0, &item.name, &cell_str_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_string_with_format(row_idx, 1, &item.spec, &cell_str_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_string_with_format(row_idx, 2, &item.dosage, &cell_str_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_string_with_format(row_idx, 3, &item.factory, &cell_str_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_string_with_format(row_idx, 4, &item.unit, &cell_str_fmt)
            .map_err(|e| e.to_string())?;

        ws.write_number_with_format(row_idx, 5, item.qty, &cell_int_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_number_with_format(row_idx, 6, item.cost_amt, &cell_num_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_number_with_format(row_idx, 7, item.retail_amt, &cell_num_fmt)
            .map_err(|e| e.to_string())?;
        ws.write_number_with_format(row_idx, 8, item.stock, &cell_int_fmt)
            .map_err(|e| e.to_string())?;

        total_qty_sum += item.qty;
        total_cost_sum += item.cost_amt;
        total_retail_sum += item.retail_amt;

        row_idx += 1;
    }

    // 设置自适应列宽
    ws.set_column_width(0, 26).map_err(|e| e.to_string())?;
    ws.set_column_width(1, 16).map_err(|e| e.to_string())?;
    ws.set_column_width(2, 10).map_err(|e| e.to_string())?;
    ws.set_column_width(3, 26).map_err(|e| e.to_string())?;
    ws.set_column_width(4, 8).map_err(|e| e.to_string())?;
    ws.set_column_width(5, 12).map_err(|e| e.to_string())?;
    ws.set_column_width(6, 15).map_err(|e| e.to_string())?;
    ws.set_column_width(7, 15).map_err(|e| e.to_string())?;
    ws.set_column_width(8, 12).map_err(|e| e.to_string())?;

    out_wb
        .save(&out_path_buf)
        .map_err(|e| format!("保存汇总 Excel 失败: {}", e))?;

    Ok(SalesSummaryResult {
        success: true,
        output_file: out_path_buf.to_string_lossy().to_string(),
        total_raw_rows: raw_count,
        unique_drugs_count: map.len(),
        total_qty: (total_qty_sum * 100.0).round() / 100.0,
        total_cost_amt: (total_cost_sum * 100.0).round() / 100.0,
        total_retail_amt: (total_retail_sum * 100.0).round() / 100.0,
        error: None,
    })
}
