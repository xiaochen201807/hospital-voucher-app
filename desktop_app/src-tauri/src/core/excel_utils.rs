use calamine::{open_workbook_auto, Data, Sheets};
use std::path::Path;

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
    for (idx, cell) in header.iter().enumerate() {
        let name = cell_as_string(cell).replace([' ', '\t', '\r', '\n'], "");
        for cand in candidates {
            if name.contains(cand) {
                return Some(idx);
            }
        }
    }
    None
}

/// 打开任意 Excel 文件 (.xlsx 或 .xls)
pub fn open_excel(path: &Path) -> Result<Sheets<std::io::BufReader<std::fs::File>>, String> {
    open_workbook_auto(path).map_err(|e| format!("打开 Excel 文件 '{:?}' 失败: {}", path, e))
}
