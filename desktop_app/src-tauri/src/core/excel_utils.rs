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
    match open_workbook_auto(path) {
        Ok(sheets) => Ok(sheets),
        Err(e) => {
            use std::io::Read;
            // 尝试读取前 512 字节探测真实文件类型
            if let Ok(mut file) = std::fs::File::open(path) {
                let mut buf = vec![0u8; 512];
                let n = file.read(&mut buf).unwrap_or(0);
                buf.truncate(n);

                // 1. 如果是 PK.. (ZIP / XLSX)，尝试以 Xlsx 打开
                if buf.starts_with(b"PK\x03\x04") {
                    if let Ok(xlsx) = calamine::open_workbook::<calamine::Xlsx<_>, _>(path) {
                        return Ok(Sheets::Xlsx(xlsx));
                    }
                }

                // 2. 如果是 OLE 复合文档头，尝试以 Xls 打开
                if buf.starts_with(&[0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]) {
                    if let Ok(xls) = calamine::open_workbook::<calamine::Xls<_>, _>(path) {
                        return Ok(Sheets::Xls(xls));
                    }
                }

                // 3. 如果是 HTML 或 XML 伪 Excel 表格
                let sample_str = String::from_utf8_lossy(&buf).to_lowercase();
                if sample_str.contains("<html")
                    || sample_str.contains("<table")
                    || sample_str.contains("<?xml")
                    || sample_str.contains("xmlns:")
                {
                    return Err(format!(
                        "该文件疑似为医院/库管系统导出的 HTML/XML 网页伪表格（非标准 Excel 二进制文件）。\n\n【解决方法】：请先用 WPS 或 Microsoft Excel 打开该报表，点击【文件】->【另存为】，格式选择【Excel 工作簿 (*.xlsx)】保存后再重新导入！"
                    ));
                }

                if n == 0 {
                    return Err("文件大小为 0 字节（空文件），请检查导出是否完整。".to_string());
                }
            }

            Err(format!(
                "打开 Excel 文件失败: {}\n\n【排查建议】：如果该报表由医院系统导出，请先用 WPS 或 Excel 打开并【另存为】标准的【.xlsx】格式后再导入使用。",
                e
            ))
        }
    }
}

