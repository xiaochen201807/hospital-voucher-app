use crate::core::excel_utils::{cell_as_f64, cell_as_string, find_exact_col_idx, read_sheet_rows};
use crate::core::models::LedgerEntry;
use calamine::Data;
use std::path::Path;

pub type CategorizedLedger = (Vec<LedgerEntry>, Vec<LedgerEntry>, Vec<LedgerEntry>);

#[derive(Debug, Clone, Copy)]
struct LedgerColumns {
    code: usize,
    init_price: usize,
    end_qty: usize,
    end_price: usize,
    end_amount: usize,
}

fn grouped_column_start(row: &[Data], keyword: &str) -> Option<usize> {
    row.iter()
        .position(|cell| cell_as_string(cell).contains(keyword))
}

fn subheader_column(header: &[Data], start: usize, label: &str) -> Option<usize> {
    header
        .iter()
        .enumerate()
        .skip(start)
        .find_map(|(index, cell)| (cell_as_string(cell) == label).then_some(index))
}

/// 从总账的分组表头与明细表头中定位数量、单价、金额列。
///
/// 医院导出的总账列顺序虽然目前稳定，但“期初/期末”是两行表头，直接写死
/// 17/18/19 列容易在系统新增列后静默读错。保留旧列号作为最后兜底，兼容没有
/// 标准表头的历史文件。
fn discover_ledger_columns(rows: &[Vec<Data>]) -> LedgerColumns {
    let fallback = LedgerColumns {
        code: 0,
        init_price: 5,
        end_qty: 16,
        end_price: 17,
        end_amount: 18,
    };

    let Some(header_idx) = rows.iter().take(12).position(|row| {
        find_exact_col_idx(row, &["科目编码"]).is_some()
            && find_exact_col_idx(row, &["方向"]).is_some()
    }) else {
        return fallback;
    };
    let header = &rows[header_idx];
    let code = find_exact_col_idx(header, &["科目编码"]).unwrap_or(fallback.code);
    let Some(group_row) = header_idx.checked_sub(1).and_then(|idx| rows.get(idx)) else {
        return fallback;
    };
    let Some(init_start) = grouped_column_start(group_row, "期初") else {
        return fallback;
    };
    let Some(end_start) = grouped_column_start(group_row, "期末") else {
        return fallback;
    };

    let Some(init_price) = subheader_column(header, init_start, "单价") else {
        return fallback;
    };
    let Some(end_qty) = subheader_column(header, end_start, "数量") else {
        return fallback;
    };
    let Some(end_price) = subheader_column(header, end_start, "单价") else {
        return fallback;
    };
    let Some(end_amount) = subheader_column(header, end_start, "金额") else {
        return fallback;
    };

    LedgerColumns {
        code,
        init_price,
        end_qty,
        end_price,
        end_amount,
    }
}

/// 读取财务总账中的存货科目。
pub fn load_ledger_entries(ledger_path: &Path) -> Result<Vec<LedgerEntry>, String> {
    let (_, rows) = read_sheet_rows(ledger_path, &["总账", "数量金额"], "总账")?;
    let columns = discover_ledger_columns(&rows);

    let mut entries = Vec::new();
    for row in &rows {
        if row.len() < 4 {
            continue;
        }

        let code = row
            .get(columns.code)
            .map(cell_as_string)
            .unwrap_or_default();
        if !code.starts_with("1201_") {
            continue;
        }

        let name_full = cell_as_string(&row[1]);
        let aux_code = code.trim_start_matches("1201_").to_string();
        let raw_name = name_full.trim_start_matches("存货_").trim();
        let mut name_parts = raw_name.splitn(2, ' ');
        let drug_name = name_parts.next().unwrap_or_default().to_string();
        let spec = name_parts.next().unwrap_or_default().to_string();

        // 结存数量、期末单价、期末金额优先按总账的两行表头定位；没有标准表头时，
        // `discover_ledger_columns` 才会退回历史文件的固定列号。
        let end_qty = row.get(columns.end_qty).map(cell_as_f64).unwrap_or(0.0);
        let end_price = row.get(columns.end_price).map(cell_as_f64).unwrap_or(0.0);
        let end_amount = row.get(columns.end_amount).map(cell_as_f64).unwrap_or(0.0);
        let init_price = row.get(columns.init_price).map(cell_as_f64).unwrap_or(0.0);
        let price = if end_price > 0.0 {
            end_price
        } else if init_price > 0.0 {
            init_price
        } else {
            0.0
        };

        entries.push(LedgerEntry {
            code,
            aux_code,
            name_full,
            drug_name,
            spec,
            price,
            end_qty,
            end_amount,
        });
    }

    Ok(entries)
}

/// 按存货科目编码将总账分成西药房、中药房和耗材库。
pub fn categorize_ledger(entries: &[LedgerEntry]) -> CategorizedLedger {
    let mut ledger_xy = Vec::new();
    let mut ledger_zy = Vec::new();
    let mut ledger_hc = Vec::new();

    for entry in entries {
        if entry.code.starts_with("1201_XY") {
            ledger_xy.push(entry.clone());
        } else if entry.code.starts_with("1201_ZY") || entry.code.starts_with("1201_KL") {
            ledger_zy.push(entry.clone());
        } else if entry.code.starts_with("1201_HC") {
            ledger_hc.push(entry.clone());
        }
    }

    (ledger_xy, ledger_zy, ledger_hc)
}

/// 一次读取总账并完成分类，供账实核对使用。
pub fn load_categorized_ledger(ledger_path: &Path) -> Result<CategorizedLedger, String> {
    let entries = load_ledger_entries(ledger_path)?;
    Ok(categorize_ledger(&entries))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn set_cell(row: &mut [Data], index: usize, value: &str) {
        row[index] = Data::String(value.to_string());
    }

    #[test]
    fn discovers_current_two_row_ledger_header() {
        let mut group_row = vec![Data::Empty; 19];
        set_cell(&mut group_row, 3, "期初余额");
        set_cell(&mut group_row, 15, "期末余额");

        let mut header_row = vec![Data::Empty; 19];
        set_cell(&mut header_row, 0, "科目编码");
        set_cell(&mut header_row, 3, "方向");
        set_cell(&mut header_row, 4, "数量");
        set_cell(&mut header_row, 5, "单价");
        set_cell(&mut header_row, 6, "金额");
        set_cell(&mut header_row, 15, "方向");
        set_cell(&mut header_row, 16, "数量");
        set_cell(&mut header_row, 17, "单价");
        set_cell(&mut header_row, 18, "金额");

        let columns = discover_ledger_columns(&[group_row, header_row]);

        assert_eq!(columns.code, 0);
        assert_eq!(columns.init_price, 5);
        assert_eq!(columns.end_qty, 16);
        assert_eq!(columns.end_price, 17);
        assert_eq!(columns.end_amount, 18);
    }
}
