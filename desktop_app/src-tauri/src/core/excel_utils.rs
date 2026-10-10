use calamine::{open_workbook_auto, Data, Reader, Sheets};
use quick_xml::{events::Event, Reader as XmlReader};
use std::path::Path;

/// 读取后的工作表数据。
pub type SheetRows = Vec<Vec<Data>>;

/// 从单元格数据转换为纯字符串（去除首尾空白）
pub fn cell_as_string(cell: &Data) -> String {
    match cell {
        Data::String(s) => s.trim().to_string(),
        Data::Float(f) => {
            if f.fract() == 0.0 {
                format!("{:.0}", f)
            } else {
                format!("{}", f)
            }
        }
        Data::Int(i) => format!("{}", i),
        Data::Bool(b) => format!("{}", b),
        Data::DateTime(d) => format!("{}", d),
        Data::DateTimeIso(s) | Data::DurationIso(s) => s.trim().to_string(),
        Data::Error(e) => format!("{:?}", e),
        Data::Empty => String::new(),
    }
}

/// 从单元格转换为 f64 浮点数值
pub fn cell_as_f64(cell: &Data) -> f64 {
    match cell {
        Data::Float(f) => *f,
        Data::Int(i) => *i as f64,
        Data::String(s) => {
            let cleaned = s.replace(',', "").trim().to_string();
            cleaned.parse::<f64>().unwrap_or(0.0)
        }
        _ => 0.0,
    }
}

/// 在表头行中根据候选列名列表查找匹配的列索引
pub fn find_col_idx(header: &[Data], candidates: &[&str]) -> Option<usize> {
    let normalized_candidates: Vec<String> = candidates
        .iter()
        .map(|candidate| normalize_header(candidate))
        .collect();

    // 先做完整匹配，避免“金额”抢先匹配到“零售金额”等更具体的列。
    for (idx, cell) in header.iter().enumerate() {
        let name = normalize_header(&cell_as_string(cell));
        if normalized_candidates.contains(&name) {
            return Some(idx);
        }
    }

    // 兼容医院系统导出的长表头，再使用模糊匹配；候选词按长度降序，优先具体别名。
    let mut candidate_order: Vec<usize> = (0..normalized_candidates.len()).collect();
    candidate_order.sort_by_key(|idx| std::cmp::Reverse(normalized_candidates[*idx].len()));
    for candidate_idx in &candidate_order {
        for (idx, cell) in header.iter().enumerate() {
            let name = normalize_header(&cell_as_string(cell));
            if name.contains(&normalized_candidates[*candidate_idx]) {
                return Some(idx);
            }
        }
    }
    None
}

fn normalize_header(value: &str) -> String {
    value
        .chars()
        .filter(|c| !matches!(*c, ' ' | '\t' | '\r' | '\n' | '　'))
        .collect::<String>()
        .to_lowercase()
}

fn select_sheet_name(
    sheet_names: &[String],
    exact_hints: &[&str],
    contains_hints: &[&str],
    prefer_last_contains: bool,
) -> Option<String> {
    for hint in exact_hints {
        if let Some(name) = sheet_names.iter().find(|name| name.as_str() == *hint) {
            return Some(name.clone());
        }
    }

    let contains_match = |name: &&String| contains_hints.iter().any(|hint| name.contains(hint));
    if prefer_last_contains {
        sheet_names.iter().rev().find(contains_match).cloned()
    } else {
        sheet_names.iter().find(contains_match).cloned()
    }
}

/// 按工作表名称提示读取数据；未命中提示时使用第一个工作表。
pub fn read_sheet_rows(
    path: &Path,
    sheet_hints: &[&str],
    context: &str,
) -> Result<(String, SheetRows), String> {
    if is_excel_xml_file(path) {
        return read_excel_xml_sheet_rows(path, sheet_hints, context);
    }

    let mut workbook = open_excel(path)?;
    let sheet_names = workbook.sheet_names();
    let sheet_name = select_sheet_name(&sheet_names, &[], sheet_hints, false)
        .or_else(|| sheet_names.first().cloned())
        .ok_or_else(|| format!("{}中无有效工作表", context))?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取{}失败: {}", context, e))?;
    let rows = range.rows().map(|row| row.to_vec()).collect();
    Ok((sheet_name, rows))
}

/// 读取需要优先使用汇总 Sheet 的工作簿。
///
/// 先按名称精确命中，再从包含提示词的 Sheet 中选择最后一个。这样第一步生成的
/// “原始销售表 + 销售明细”双 Sheet 文件，会优先读取去重后的“销售明细”。
pub fn read_sheet_rows_prefer(
    path: &Path,
    exact_sheet_hints: &[&str],
    sheet_hints: &[&str],
    context: &str,
) -> Result<(String, SheetRows), String> {
    if is_excel_xml_file(path) {
        let sheets = read_excel_xml_sheets(path, context)?;
        let sheet_names = sheets
            .iter()
            .map(|(name, _)| name.clone())
            .collect::<Vec<_>>();
        let sheet_name = select_sheet_name(&sheet_names, exact_sheet_hints, sheet_hints, true)
            .or_else(|| sheet_names.last().cloned())
            .ok_or_else(|| format!("{}中无有效工作表", context))?;
        let rows = sheets
            .into_iter()
            .find(|(name, _)| name == &sheet_name)
            .map(|(_, rows)| rows)
            .ok_or_else(|| format!("{}中无有效工作表", context))?;
        return Ok((sheet_name, rows));
    }

    let mut workbook = open_excel(path)?;
    let sheet_names = workbook.sheet_names();
    let sheet_name = select_sheet_name(&sheet_names, exact_sheet_hints, sheet_hints, true)
        .or_else(|| sheet_names.last().cloned())
        .ok_or_else(|| format!("{}中无有效工作表", context))?;

    let range = workbook
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取{}失败: {}", context, e))?;
    let rows = range.rows().map(|row| row.to_vec()).collect();
    Ok((sheet_name, rows))
}

/// 在前若干行中定位表头，统一处理找不到表头的错误。
pub fn find_header_row(
    rows: &[Vec<Data>],
    candidates: &[&str],
    scan_limit: usize,
    context: &str,
) -> Result<usize, String> {
    rows.iter()
        .take(scan_limit)
        .position(|row| find_col_idx(row, candidates).is_some())
        .ok_or_else(|| format!("{}中未能在前{}行找到有效表头", context, scan_limit))
}

pub fn required_col(header: &[Data], candidates: &[&str], label: &str) -> Result<usize, String> {
    find_col_idx(header, candidates).ok_or_else(|| format!("表头中缺少必需列：{}", label))
}

/// 只查找表头完全一致的列，不执行包含匹配。
pub fn find_exact_col_idx(header: &[Data], candidates: &[&str]) -> Option<usize> {
    let normalized_candidates: Vec<String> = candidates
        .iter()
        .map(|candidate| normalize_header(candidate))
        .collect();

    header.iter().enumerate().find_map(|(idx, cell)| {
        let name = normalize_header(&cell_as_string(cell));
        normalized_candidates.contains(&name).then_some(idx)
    })
}

pub fn optional_col(header: &[Data], candidates: &[&str], fallback: usize) -> usize {
    find_col_idx(header, candidates).unwrap_or(fallback)
}

pub fn row_as_string(row: &[Data], col: usize) -> String {
    row.get(col).map(cell_as_string).unwrap_or_default()
}

pub fn row_as_f64(row: &[Data], col: usize) -> f64 {
    row.get(col).map(cell_as_f64).unwrap_or(0.0)
}

/// 将金额统一按人民币分（两位小数）取整。
pub fn round_currency(value: f64) -> f64 {
    if !value.is_finite() {
        return 0.0;
    }
    (value * 100.0).round() / 100.0
}

/// 将金额转换为分，供凭证合计使用，避免逐笔金额相加时累积浮点误差。
pub fn currency_cents(value: f64) -> i64 {
    if !value.is_finite() {
        return 0;
    }
    (value * 100.0).round() as i64
}

pub fn cents_to_currency(value: i64) -> f64 {
    value as f64 / 100.0
}

/// 判断数量是否可以视为 0，兼容 Excel 浮点读取产生的极小残差。
pub fn is_effectively_zero(value: f64) -> bool {
    value.abs() < 1e-9
}

pub fn is_summary_row(name: &str) -> bool {
    name.is_empty() || name.contains("合计") || name.contains("总计")
}

pub fn amount_or_product(amount: f64, qty: f64, price: f64) -> f64 {
    if amount == 0.0 && qty > 0.0 && price > 0.0 {
        round_currency(qty * price)
    } else {
        amount
    }
}

pub fn collect_source_dates(
    rows: &[Vec<Data>],
    start_row: usize,
    name_col: usize,
    date_col: Option<usize>,
) -> Vec<String> {
    let Some(date_col) = date_col else {
        return Vec::new();
    };

    rows.iter()
        .skip(start_row)
        .filter_map(|row| {
            if is_summary_row(&row_as_string(row, name_col)) {
                return None;
            }
            let value = row_as_string(row, date_col);
            (!value.is_empty()).then_some(value)
        })
        .collect()
}

const OLE_COMPOUND_FILE_MAGIC: &[u8] = &[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];

fn file_prefix(path: &Path) -> Result<Vec<u8>, std::io::Error> {
    use std::io::Read;

    let mut file = std::fs::File::open(path)?;
    let mut buf = vec![0u8; 512];
    let n = file.read(&mut buf)?;
    buf.truncate(n);
    Ok(buf)
}

fn is_zip_file(prefix: &[u8]) -> bool {
    prefix.starts_with(b"PK\x03\x04")
        || prefix.starts_with(b"PK\x05\x06")
        || prefix.starts_with(b"PK\x07\x08")
}

fn is_excel_xml_prefix(prefix: &[u8]) -> bool {
    let sample = String::from_utf8_lossy(prefix)
        .trim_start_matches('\u{feff}')
        .trim_start()
        .to_lowercase();
    (sample.starts_with("<?xml") || sample.starts_with("<workbook"))
        && sample.contains("<workbook")
        && sample.contains("schemas-microsoft-com:office:spreadsheet")
}

fn is_html_or_generic_xml_prefix(prefix: &[u8]) -> bool {
    let sample = String::from_utf8_lossy(prefix).to_lowercase();
    sample.contains("<html")
        || sample.contains("<table")
        || sample.contains("<?xml")
        || sample.contains("xmlns:")
}

fn is_excel_xml_file(path: &Path) -> bool {
    file_prefix(path)
        .map(|prefix| is_excel_xml_prefix(&prefix))
        .unwrap_or(false)
}

fn xml_import_error(context: &str, error: impl std::fmt::Display) -> String {
    format!("解析{} Excel 2003 XML 表格失败: {}", context, error)
}

/// 读取 Excel 2003 XML（许多医院系统会将它导出为 `.xls`）中的全部工作表。
///
/// 这类文件使用 `ss:Index` 表示稀疏单元格列号，单元格值统一读取为字符串，后续
/// 的表头匹配和数值转换会沿用标准 Excel 读取路径。
fn read_excel_xml_sheets(path: &Path, context: &str) -> Result<Vec<(String, SheetRows)>, String> {
    let xml = std::fs::read_to_string(path).map_err(|e| xml_import_error(context, e))?;
    let mut reader = XmlReader::from_str(&xml);
    reader.config_mut().trim_text(true);

    let mut sheets: Vec<(String, SheetRows)> = Vec::new();
    let mut current_sheet: Option<usize> = None;
    let mut current_row: Option<Vec<Data>> = None;
    let mut current_column = 0usize;
    let mut current_cell_column: Option<usize> = None;
    let mut in_cell_data = false;
    let mut cell_value = String::new();

    loop {
        match reader.read_event() {
            Ok(Event::Start(event)) if event.local_name().as_ref() == "Worksheet" => {
                let name = event
                    .attributes()
                    .flatten()
                    .find(|attribute| attribute.key.local_name().as_ref() == "Name")
                    .map(|attribute| attribute.value.as_ref().trim().to_string())
                    .filter(|name| !name.is_empty())
                    .unwrap_or_else(|| format!("Sheet{}", sheets.len() + 1));
                sheets.push((name, Vec::new()));
                current_sheet = Some(sheets.len() - 1);
            }
            Ok(Event::Start(event))
                if current_sheet.is_some() && event.local_name().as_ref() == "Row" =>
            {
                current_row = Some(Vec::new());
                current_column = 0;
            }
            Ok(Event::Start(event))
                if current_row.is_some() && event.local_name().as_ref() == "Cell" =>
            {
                current_column += 1;
                for attribute in event.attributes().flatten() {
                    if attribute.key.local_name().as_ref() == "Index" {
                        current_column = attribute
                            .value
                            .as_ref()
                            .parse::<usize>()
                            .map_err(|_| xml_import_error(context, "存在无效的单元格列号"))?;
                    }
                }
                if current_column == 0 || current_column > 10000 {
                    return Err(xml_import_error(context, "存在无效的单元格列号"));
                }
                if let Some(row) = current_row.as_mut() {
                    row.resize(row.len().max(current_column), Data::Empty);
                }
                current_cell_column = Some(current_column);
            }
            Ok(Event::Start(event))
                if current_cell_column.is_some() && event.local_name().as_ref() == "Data" =>
            {
                in_cell_data = true;
                cell_value.clear();
            }
            Ok(Event::Empty(event))
                if current_row.is_some() && event.local_name().as_ref() == "Cell" =>
            {
                current_column += 1;
                for attribute in event.attributes().flatten() {
                    if attribute.key.local_name().as_ref() == "Index" {
                        current_column = attribute
                            .value
                            .as_ref()
                            .parse::<usize>()
                            .map_err(|_| xml_import_error(context, "存在无效的单元格列号"))?;
                    }
                }
                if current_column == 0 || current_column > 10000 {
                    return Err(xml_import_error(context, "存在无效的单元格列号"));
                }
                if let Some(row) = current_row.as_mut() {
                    row.resize(row.len().max(current_column), Data::Empty);
                }
            }
            Ok(Event::Empty(event))
                if current_cell_column.is_some() && event.local_name().as_ref() == "Data" =>
            {
                if let (Some(row), Some(column)) = (current_row.as_mut(), current_cell_column) {
                    row[column - 1] = Data::String(String::new());
                }
            }
            Ok(Event::Text(event)) if in_cell_data => {
                let value = quick_xml::escape::unescape(event.as_ref())
                    .map_err(|e| xml_import_error(context, e))?;
                cell_value.push_str(&value);
            }
            Ok(Event::CData(event)) if in_cell_data => {
                cell_value.push_str(event.as_ref());
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == "Data" => {
                if in_cell_data {
                    if let (Some(row), Some(column)) = (current_row.as_mut(), current_cell_column) {
                        row[column - 1] = Data::String(cell_value.trim().to_string());
                    }
                    in_cell_data = false;
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == "Cell" => {
                current_cell_column = None;
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == "Row" => {
                if let (Some(sheet), Some(row)) = (current_sheet, current_row.take()) {
                    sheets[sheet].1.push(row);
                }
            }
            Ok(Event::End(event)) if event.local_name().as_ref() == "Worksheet" => {
                current_sheet = None;
                current_row = None;
                current_cell_column = None;
            }
            Ok(Event::Eof) => break,
            Err(error) => return Err(xml_import_error(context, error)),
            _ => {}
        }
    }

    Ok(sheets)
}

fn read_excel_xml_sheet_rows(
    path: &Path,
    sheet_hints: &[&str],
    context: &str,
) -> Result<(String, SheetRows), String> {
    let sheets = read_excel_xml_sheets(path, context)?;
    let sheet_names = sheets
        .iter()
        .map(|(name, _)| name.clone())
        .collect::<Vec<_>>();
    let sheet_name = select_sheet_name(&sheet_names, &[], sheet_hints, false)
        .or_else(|| sheet_names.first().cloned())
        .ok_or_else(|| format!("{}中无有效工作表", context))?;
    let rows = sheets
        .into_iter()
        .find(|(name, _)| name == &sheet_name)
        .map(|(_, rows)| rows)
        .ok_or_else(|| format!("{}中无有效工作表", context))?;
    Ok((sheet_name, rows))
}

/// 打开任意 Excel 文件（`.xlsx` 或 `.xls`），优先依据文件内容识别格式。
pub fn open_excel(path: &Path) -> Result<Sheets<std::io::BufReader<std::fs::File>>, String> {
    let prefix = file_prefix(path).ok();

    // calamine 的 open_workbook_auto 只依据扩展名。先按标准文件签名处理，兼容
    // “真实格式与扩展名不一致”的 WPS/医院系统导出文件。
    if let Some(prefix) = prefix.as_deref() {
        if is_zip_file(prefix) {
            return calamine::open_workbook::<calamine::Xlsx<_>, _>(path)
                .map(Sheets::Xlsx)
                .map_err(|error| format!("XLSX 文件解析失败: {}", error));
        }
        if prefix.starts_with(OLE_COMPOUND_FILE_MAGIC) {
            return calamine::open_workbook::<calamine::Xls<_>, _>(path)
                .map(Sheets::Xls)
                .map_err(|error| format!("XLS 文件解析失败: {}", error));
        }
        if is_excel_xml_prefix(prefix) {
            return Err(
                "该文件是 Excel 2003 XML 表格（扩展名通常为 .xls），请通过表格导入功能读取，或用 WPS/Excel 另存为标准 .xlsx 后再打开。"
                    .to_string(),
            );
        }
        if prefix.is_empty() {
            return Err("文件大小为 0 字节（空文件），请检查导出是否完整。".to_string());
        }
    }

    match open_workbook_auto(path) {
        Ok(sheets) => Ok(sheets),
        Err(error) => {
            if let Some(prefix) = prefix.as_deref() {
                if is_html_or_generic_xml_prefix(prefix) {
                    return Err(
                        "该文件疑似为医院/库管系统导出的 HTML/XML 网页伪表格（非标准 Excel 文件）。\n\n【解决方法】：请先用 WPS 或 Microsoft Excel 打开该报表，点击【文件】->【另存为】，格式选择【Excel 工作簿 (*.xlsx)】保存后再重新导入！"
                            .to_string(),
                    );
                }
            }
            Err(format!(
                "打开 Excel 文件失败: {}\n\n【排查建议】：请确认文件是标准的 .xls 或 .xlsx；如果报表由医院系统导出，请先用 WPS 或 Excel 打开并另存为标准 Excel 工作簿后再导入。",
                error
            ))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use rust_xlsxwriter::Workbook;

    #[test]
    fn exact_header_match_beats_broad_contains_match() {
        let header = vec![
            Data::String("零售金额".to_string()),
            Data::String("进价金额".to_string()),
        ];

        assert_eq!(find_col_idx(&header, &["金额", "进价金额"]), Some(1));
    }

    #[test]
    fn amount_falls_back_to_quantity_times_price() {
        assert_eq!(amount_or_product(0.0, 3.0, 1.234), 3.70);
        assert_eq!(amount_or_product(9.99, 3.0, 1.234), 9.99);
    }

    #[test]
    fn currency_totals_are_accumulated_in_cents() {
        let cents = currency_cents(3.0 * 1.234) + currency_cents(2.0 * 0.567);
        assert_eq!(cents_to_currency(cents), 4.83);
        assert_eq!(round_currency(4.236), 4.24);
    }

    #[test]
    fn tiny_quantity_residual_is_treated_as_zero() {
        assert!(is_effectively_zero(0.0));
        assert!(is_effectively_zero(-1e-10));
        assert!(!is_effectively_zero(1e-6));
    }

    #[test]
    fn summary_rows_are_shared_by_all_importers() {
        assert!(is_summary_row("合计"));
        assert!(is_summary_row("总计金额"));
        assert!(is_summary_row(""));
        assert!(!is_summary_row("阿莫西林"));
    }

    #[test]
    fn summary_sheet_selection_prefers_exact_then_last_matching_sheet() {
        let sheet_names = vec![
            "原始销售表".to_string(),
            "销售明细".to_string(),
            "备注销售".to_string(),
        ];

        assert_eq!(
            select_sheet_name(&sheet_names, &["销售明细"], &["销售"], true),
            Some("销售明细".to_string())
        );
        assert_eq!(
            select_sheet_name(&sheet_names, &["不存在"], &["销售"], true),
            Some("备注销售".to_string())
        );
    }

    #[test]
    fn exact_column_lookup_does_not_treat_amount_as_unit_price() {
        let header = vec![
            Data::String("数量".to_string()),
            Data::String("进价金额".to_string()),
        ];

        assert_eq!(find_exact_col_idx(&header, &["进价"]), None);
        assert_eq!(find_exact_col_idx(&header, &["进价金额"]), Some(1));
    }

    #[test]
    fn xlsx_content_is_opened_even_when_extension_is_xls() {
        let path = std::env::temp_dir().join(format!(
            "excel_signature_{}_{}.xls",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        let mut workbook = Workbook::new();
        let worksheet = workbook.add_worksheet();
        worksheet.set_name("销售明细").unwrap();
        worksheet.write_string(0, 0, "药品名称").unwrap();
        worksheet.write_string(1, 0, "测试药品").unwrap();
        workbook.save(&path).unwrap();

        let mut opened = open_excel(&path).unwrap();
        assert_eq!(opened.sheet_names(), vec!["销售明细".to_string()]);
        let range = opened.worksheet_range("销售明细").unwrap();
        assert_eq!(cell_as_string(&range.get((1, 0)).unwrap()), "测试药品");

        std::fs::remove_file(path).unwrap();
    }

    #[test]
    fn excel_2003_xml_xls_is_read_by_common_sheet_loader() {
        let path = std::env::temp_dir().join(format!(
            "excel_xml_{}_{}.xls",
            std::process::id(),
            std::time::SystemTime::now()
                .duration_since(std::time::UNIX_EPOCH)
                .unwrap()
                .as_nanos()
        ));
        let xml = r#"<?xml version="1.0" encoding="UTF-8"?>
<Workbook xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
  <Worksheet ss:Name="库存明细"><Table>
    <Row><Cell><Data ss:Type="String">名称</Data></Cell><Cell><Data ss:Type="String">数量</Data></Cell></Row>
    <Row><Cell><Data ss:Type="String">测试药品</Data></Cell><Cell ss:Index="2"><Data ss:Type="Number">3</Data></Cell></Row>
  </Table></Worksheet>
</Workbook>"#;
        std::fs::write(&path, xml).unwrap();

        let (sheet, rows) = read_sheet_rows(&path, &["库存"], "库存报表").unwrap();
        assert_eq!(sheet, "库存明细");
        assert_eq!(cell_as_string(&rows[0][0]), "名称");
        assert_eq!(cell_as_string(&rows[1][1]), "3");
        assert_eq!(cell_as_f64(&rows[1][1]), 3.0);

        std::fs::remove_file(path).unwrap();
    }
}
