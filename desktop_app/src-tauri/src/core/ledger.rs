use crate::core::excel_utils::{cell_as_f64, cell_as_string, open_excel};
use crate::core::models::LedgerEntry;
use calamine::Reader;
use std::path::Path;

pub type CategorizedLedger = (Vec<LedgerEntry>, Vec<LedgerEntry>, Vec<LedgerEntry>);

/// 读取财务总账中的存货科目。
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
        let aux_code = code.trim_start_matches("1201_").to_string();
        let raw_name = name_full.trim_start_matches("存货_").trim();
        let mut name_parts = raw_name.splitn(2, ' ');
        let drug_name = name_parts.next().unwrap_or_default().to_string();
        let spec = name_parts.next().unwrap_or_default().to_string();

        // 结存数量为第 17 列，期末单价为第 18 列，期初单价为第 6 列。
        let end_qty = row.get(16).map(cell_as_f64).unwrap_or(0.0);
        let end_price = row.get(17).map(cell_as_f64).unwrap_or(0.0);
        let init_price = row.get(5).map(cell_as_f64).unwrap_or(0.0);
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
