use crate::core::excel_utils::cell_as_string;
use crate::core::excel_utils::open_excel;
use calamine::{Data, Reader};
use rust_xlsxwriter::{Format, FormatAlign, FormatBorder, Workbook, Worksheet};
use std::path::Path;

/// 凭证导入模板的标准 26 列。
pub const VOUCHER_HEADERS: [&str; 26] = [
    "日期",
    "凭证字",
    "凭证号",
    "附件数",
    "分录序号",
    "摘要",
    "科目代码",
    "科目名称",
    "借方金额",
    "贷方金额",
    "客户",
    "供应商",
    "职员",
    "项目",
    "部门",
    "存货",
    "是否限定",
    "自定义类别",
    "自定义编码",
    "自定义类别1",
    "自定义编码1",
    "数量",
    "单价",
    "原币金额",
    "币别",
    "汇率",
];

/// 读取模板时只保留单元格数据，避免入库和出库各自实现一遍模板遍历。
#[derive(Debug, Clone)]
pub struct TemplateSheet {
    pub name: String,
    pub rows: Vec<Vec<Data>>,
}

pub fn load_template_sheets(path: &Path) -> Result<Vec<TemplateSheet>, String> {
    let mut workbook = open_excel(path)?;
    let mut sheets = Vec::new();

    for sheet_name in workbook.sheet_names() {
        let range = workbook
            .worksheet_range(&sheet_name)
            .map_err(|e| format!("读取凭证模板工作表 '{}' 失败: {}", sheet_name, e))?;
        sheets.push(TemplateSheet {
            name: sheet_name,
            rows: range.rows().map(|row| row.to_vec()).collect(),
        });
    }

    Ok(sheets)
}

/// 将模板中的非凭证附表复制到新工作簿。
///
/// 当前使用 calamine + rust_xlsxwriter，只能可靠复制单元格值；样式、合并和数据验证
/// 等元数据仍无法由这两个库完整保留。把这段逻辑集中到这里，至少保证入库和出库行为一致。
pub fn copy_template_sheets(
    workbook: &mut Workbook,
    sheets: &[TemplateSheet],
) -> Result<(), String> {
    for sheet in sheets {
        if sheet.name.contains("凭证") || sheet.name.contains("模版") {
            continue;
        }

        let worksheet = workbook.add_worksheet();
        worksheet.set_name(&sheet.name).map_err(|e| e.to_string())?;

        for (row_idx, row) in sheet.rows.iter().enumerate() {
            for (col_idx, cell) in row.iter().enumerate() {
                let row_idx = row_idx as u32;
                let col_idx = col_idx as u16;
                match cell {
                    Data::Float(value) => {
                        worksheet
                            .write_number(row_idx, col_idx, *value)
                            .map_err(|e| e.to_string())?;
                    }
                    Data::Int(value) => {
                        worksheet
                            .write_number(row_idx, col_idx, *value as f64)
                            .map_err(|e| e.to_string())?;
                    }
                    Data::String(value) => {
                        worksheet
                            .write_string(row_idx, col_idx, value)
                            .map_err(|e| e.to_string())?;
                    }
                    Data::Bool(value) => {
                        worksheet
                            .write_boolean(row_idx, col_idx, *value)
                            .map_err(|e| e.to_string())?;
                    }
                    Data::DateTime(_)
                    | Data::DateTimeIso(_)
                    | Data::DurationIso(_)
                    | Data::Error(_) => {
                        let value = cell_as_string(cell);
                        if !value.is_empty() {
                            worksheet
                                .write_string(row_idx, col_idx, &value)
                                .map_err(|e| e.to_string())?;
                        }
                    }
                    Data::Empty => {}
                }
            }
        }
    }

    Ok(())
}

/// 凭证表格的公共样式。数量格式由业务方传入：出库通常为整数，入库允许两位小数。
pub struct VoucherFormats {
    pub header: Format,
    pub text: Format,
    pub left: Format,
    pub money: Format,
    pub qty: Format,
}

impl VoucherFormats {
    pub fn new(quantity_format: &str) -> Self {
        Self {
            header: Format::new()
                .set_bold()
                .set_font_size(10)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_border(FormatBorder::Thin),
            text: Format::new()
                .set_font_size(10)
                .set_align(FormatAlign::Center)
                .set_align(FormatAlign::VerticalCenter)
                .set_border(FormatBorder::Thin),
            left: Format::new()
                .set_font_size(10)
                .set_align(FormatAlign::Left)
                .set_align(FormatAlign::VerticalCenter)
                .set_border(FormatBorder::Thin),
            money: Format::new()
                .set_font_size(10)
                .set_num_format("#,##0.00")
                .set_align(FormatAlign::Right)
                .set_align(FormatAlign::VerticalCenter)
                .set_border(FormatBorder::Thin),
            qty: Format::new()
                .set_font_size(10)
                .set_num_format(quantity_format)
                .set_align(FormatAlign::Right)
                .set_align(FormatAlign::VerticalCenter)
                .set_border(FormatBorder::Thin),
        }
    }
}

pub fn write_voucher_headers(
    worksheet: &mut Worksheet,
    formats: &VoucherFormats,
) -> Result<(), String> {
    worksheet.set_row_height(0, 26).map_err(|e| e.to_string())?;
    for (col_idx, header) in VOUCHER_HEADERS.iter().enumerate() {
        worksheet
            .write_string_with_format(0, col_idx as u16, *header, &formats.header)
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

pub fn write_blank_range(
    worksheet: &mut Worksheet,
    row: u32,
    start_col: u16,
    end_col: u16,
    format: &Format,
) -> Result<(), String> {
    for col in start_col..end_col {
        worksheet
            .write_blank(row, col, format)
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

fn write_optional_string(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: Option<&str>,
    format: &Format,
) -> Result<(), String> {
    if let Some(value) = value.filter(|value| !value.is_empty()) {
        worksheet
            .write_string_with_format(row, col, value, format)
            .map_err(|e| e.to_string())?;
    } else {
        worksheet
            .write_blank(row, col, format)
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

fn write_optional_number(
    worksheet: &mut Worksheet,
    row: u32,
    col: u16,
    value: Option<f64>,
    format: &Format,
) -> Result<(), String> {
    if let Some(value) = value {
        worksheet
            .write_number_with_format(row, col, value, format)
            .map_err(|e| e.to_string())?;
    } else {
        worksheet
            .write_blank(row, col, format)
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

/// 一条凭证分录的通用表示。
pub struct VoucherLine<'a> {
    pub date: &'a str,
    pub voucher_no: Option<&'a str>,
    pub seq: u32,
    pub summary: Option<&'a str>,
    pub subject_code: Option<&'a str>,
    pub debit_amount: Option<f64>,
    pub credit_amount: Option<f64>,
    pub supplier_code: Option<&'a str>,
    pub inventory_code: Option<&'a str>,
    pub qty: Option<f64>,
    pub price: Option<f64>,
    pub original_amount: Option<f64>,
}

pub fn write_voucher_line(
    worksheet: &mut Worksheet,
    row: u32,
    line: &VoucherLine<'_>,
    formats: &VoucherFormats,
) -> Result<(), String> {
    worksheet
        .set_row_height(row, 20)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 0, line.date, &formats.text)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 1, "记", &formats.text)
        .map_err(|e| e.to_string())?;
    write_optional_string(worksheet, row, 2, line.voucher_no, &formats.text)?;
    worksheet
        .write_blank(row, 3, &formats.text)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 4, line.seq as f64, &formats.text)
        .map_err(|e| e.to_string())?;
    write_optional_string(worksheet, row, 5, line.summary, &formats.left)?;
    write_optional_string(worksheet, row, 6, line.subject_code, &formats.text)?;
    worksheet
        .write_blank(row, 7, &formats.left)
        .map_err(|e| e.to_string())?;
    write_optional_number(worksheet, row, 8, line.debit_amount, &formats.money)?;
    write_optional_number(worksheet, row, 9, line.credit_amount, &formats.money)?;

    write_blank_range(worksheet, row, 10, 15, &formats.text)?;
    write_optional_string(worksheet, row, 11, line.supplier_code, &formats.text)?;
    write_optional_string(worksheet, row, 15, line.inventory_code, &formats.text)?;
    write_blank_range(worksheet, row, 16, 21, &formats.text)?;
    write_optional_number(worksheet, row, 21, line.qty, &formats.qty)?;
    write_optional_number(worksheet, row, 22, line.price, &formats.money)?;
    write_optional_number(worksheet, row, 23, line.original_amount, &formats.money)?;
    worksheet
        .write_string_with_format(row, 24, "RMB", &formats.text)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 25, 1.0, &formats.text)
        .map_err(|e| e.to_string())?;

    Ok(())
}

pub fn set_default_voucher_column_widths(
    worksheet: &mut Worksheet,
    overrides: &[(u16, f64)],
) -> Result<(), String> {
    for col in 0..26 {
        worksheet
            .set_column_width(col, 13)
            .map_err(|e| e.to_string())?;
    }
    for (col, width) in overrides {
        worksheet
            .set_column_width(*col, *width)
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}
