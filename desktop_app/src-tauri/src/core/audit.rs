use crate::core::config::ConfigData;
use crate::core::excel_utils::{cell_as_f64, cell_as_string, find_col_idx, open_excel};
use crate::core::outbound::{match_drug, LedgerEntry};
use calamine::Reader;
use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::path::{Path, PathBuf};

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct AuditRecord {
    pub category: String,
    pub status: String,      // EQUAL, DIFF_QTY, WH_ONLY, LEDGER_ONLY
    pub status_desc: String, // 数量完全吻合, 存在数量差异, 仅财务有结存, 仅库管有在库
    pub name: String,
    pub spec: String,
    pub factory: String,
    pub unit: String,
    pub ledger_code: String,
    pub ledger_name: String,
    pub ledger_qty: f64,
    pub wh_qty: f64,
    pub diff_qty: f64,
    pub ledger_price: f64,
    pub wh_price: f64,
    pub ledger_amt: f64,
    pub wh_amt: f64,
    pub diff_amt: f64,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct CategorySummary {
    pub total_items: usize,
    pub equal_count: usize,
    pub diff_count: usize,
    pub wh_only_count: usize,
    pub ledger_only_count: usize,
    pub match_rate: f64,
    pub total_diff_amt: f64,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct CategoryAuditResult {
    pub category: String,
    pub summary: CategorySummary,
    pub records: Vec<AuditRecord>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct OverallAuditResult {
    pub success: bool,
    pub output_file: String,
    pub overall: CategorySummary,
    pub categories: Vec<CategoryAuditResult>,
    #[serde(default)]
    pub error: Option<String>,
}

#[derive(Debug, Clone)]
struct WarehouseItem {
    name: String,
    spec: String,
    factory: String,
    unit: String,
    qty: f64,
    price: f64,
    amount: f64,
}

/// 读取库管系统报表（西药、中药、耗材）
fn load_warehouse_items(path: &Path) -> Result<Vec<WarehouseItem>, String> {
    let mut wb = open_excel(path)?;
    let sheet_name = wb
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "库管报表中无工作表".to_string())?;

    let range = wb
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取库管报表失败: {}", e))?;

    let rows: Vec<Vec<calamine::Data>> = range.rows().map(|r| r.to_vec()).collect();
    if rows.len() < 2 {
        return Ok(Vec::new());
    }

    let mut h_idx = 0;
    for (i, r) in rows.iter().take(10).enumerate() {
        if find_col_idx(r, &["药品名称", "品名", "材料名称", "名称", "耗材名称"]).is_some() {
            h_idx = i;
            break;
        }
    }

    let header = &rows[h_idx];
    let col_name = find_col_idx(header, &["药品名称", "品名", "材料名称", "名称", "耗材名称"]).unwrap();
    let col_spec = find_col_idx(header, &["规格"]).unwrap_or(col_name + 1);
    let col_factory = find_col_idx(header, &["制药厂", "生产厂家", "厂家", "生产商"]).unwrap_or(col_spec + 1);
    let col_unit = find_col_idx(header, &["单位"]).unwrap_or(col_factory + 1);
    let col_qty = find_col_idx(header, &["数量", "结存数量", "在库数量", "库存数量"]).unwrap_or(col_unit + 1);
    let col_price = find_col_idx(header, &["单价", "成本价", "结存单价"]).unwrap_or(col_qty + 1);
    let col_amt = find_col_idx(header, &["金额", "结存金额", "成本金额"]).unwrap_or(col_price + 1);

    let mut items = Vec::new();
    for row in rows.iter().skip(h_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let name = cell_as_string(row.get(col_name).unwrap_or(&calamine::Data::Empty));
        if name.is_empty() || name.contains("合计") || name.contains("总计") {
            continue;
        }

        let spec = row.get(col_spec).map(cell_as_string).unwrap_or_default();
        let factory = row.get(col_factory).map(cell_as_string).unwrap_or_default();
        let unit = row.get(col_unit).map(cell_as_string).unwrap_or_default();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let price = row.get(col_price).map(cell_as_f64).unwrap_or(0.0);
        let mut amount = row.get(col_amt).map(cell_as_f64).unwrap_or(0.0);
        if amount == 0.0 && qty > 0.0 && price > 0.0 {
            amount = (qty * price * 100.0).round() / 100.0;
        }

        items.push(WarehouseItem {
            name,
            spec,
            factory,
            unit,
            qty,
            price,
            amount,
        });
    }

    Ok(items)
}

/// 执行单库比对
fn audit_single_category(
    category_name: &str,
    ledger_entries: &[LedgerEntry],
    wh_items: &[WarehouseItem],
    config: &ConfigData,
) -> CategoryAuditResult {
    let mut records = Vec::new();
    let mut matched_ledger_codes = HashSet::new();

    // 1. 遍历库管在库品规，匹配财务总账
    for w in wh_items {
        let matched = match_drug(&w.name, &w.spec, &w.factory, ledger_entries, config);

        if let Some(m) = matched {
            matched_ledger_codes.insert(m.code.clone());
            let l_qty = m.end_qty;
            let l_price = m.price;
            let l_amt = (l_qty * l_price * 100.0).round() / 100.0;
            let diff_qty = ((l_qty - w.qty) * 100.0).round() / 100.0;
            let diff_amt = ((l_amt - w.amount) * 100.0).round() / 100.0;

            let (status, status_desc) = if diff_qty.abs() < 0.001 {
                ("EQUAL".to_string(), "数量完全吻合".to_string())
            } else {
                ("DIFF_QTY".to_string(), "存在数量差异".to_string())
            };

            records.push(AuditRecord {
                category: category_name.to_string(),
                status,
                status_desc,
                name: w.name.clone(),
                spec: w.spec.clone(),
                factory: w.factory.clone(),
                unit: w.unit.clone(),
                ledger_code: m.code.clone(),
                ledger_name: m.name_full.clone(),
                ledger_qty: l_qty,
                wh_qty: w.qty,
                diff_qty,
                ledger_price: l_price,
                wh_price: w.price,
                ledger_amt: l_amt,
                wh_amt: w.amount,
                diff_amt,
            });
        } else {
            // 仅库管有在库，财务未建账或无此编码
            records.push(AuditRecord {
                category: category_name.to_string(),
                status: "WH_ONLY".to_string(),
                status_desc: "仅库管有在库".to_string(),
                name: w.name.clone(),
                spec: w.spec.clone(),
                factory: w.factory.clone(),
                unit: w.unit.clone(),
                ledger_code: "-".to_string(),
                ledger_name: "-".to_string(),
                ledger_qty: 0.0,
                wh_qty: w.qty,
                diff_qty: -w.qty,
                ledger_price: 0.0,
                wh_price: w.price,
                ledger_amt: 0.0,
                wh_amt: w.amount,
                diff_amt: -w.amount,
            });
        }
    }

    // 2. 检查财务总账中存在但库管未列出的品规 (仅财务有账)
    for l in ledger_entries {
        if !matched_ledger_codes.contains(&l.code) {
            let l_amt = (l.end_qty * l.price * 100.0).round() / 100.0;
            records.push(AuditRecord {
                category: category_name.to_string(),
                status: "LEDGER_ONLY".to_string(),
                status_desc: "仅财务有结存".to_string(),
                name: l.drug_name.clone(),
                spec: l.spec.clone(),
                factory: "-".to_string(),
                unit: "-".to_string(),
                ledger_code: l.code.clone(),
                ledger_name: l.name_full.clone(),
                ledger_qty: l.end_qty,
                wh_qty: 0.0,
                diff_qty: l.end_qty,
                ledger_price: l.price,
                wh_price: 0.0,
                ledger_amt: l_amt,
                wh_amt: 0.0,
                diff_amt: l_amt,
            });
        }
    }

    let total_items = records.len();
    let equal_count = records.iter().filter(|r| r.status == "EQUAL").count();
    let diff_count = records.iter().filter(|r| r.status == "DIFF_QTY").count();
    let wh_only_count = records.iter().filter(|r| r.status == "WH_ONLY").count();
    let ledger_only_count = records.iter().filter(|r| r.status == "LEDGER_ONLY").count();
    let total_diff_amt = records.iter().map(|r| r.diff_amt).sum::<f64>();
    let match_rate = if total_items > 0 {
        ((equal_count as f64 / total_items as f64) * 1000.0).round() / 10.0
    } else {
        0.0
    };

    CategoryAuditResult {
        category: category_name.to_string(),
        summary: CategorySummary {
            total_items,
            equal_count,
            diff_count,
            wh_only_count,
            ledger_only_count,
            match_rate,
            total_diff_amt: (total_diff_amt * 100.0).round() / 100.0,
        },
        records,
    }
}

/// 执行多库账实核对并输出 Excel 审计分析报告
pub fn run_inventory_audit(
    ledger_path: &str,
    west_path: Option<&str>,
    tcm_path: Option<&str>,
    hc_path: Option<&str>,
    custom_output: Option<&str>,
    config: &ConfigData,
) -> Result<OverallAuditResult, String> {
    let ledger_p = Path::new(ledger_path);
    if !ledger_p.exists() {
        return Err(format!("财务总账文件 '{:?}' 不存在", ledger_p));
    }

    // 1. 读取总账并分类划分
    let mut wb_ledger = open_excel(ledger_p)?;
    let sheet_name = wb_ledger
        .sheet_names()
        .first()
        .cloned()
        .ok_or_else(|| "总账中无有效工作表".to_string())?;

    let range = wb_ledger
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取总账失败: {}", e))?;

    let mut ledger_xy = Vec::new();
    let mut ledger_zy = Vec::new();
    let mut ledger_hc = Vec::new();

    for row in range.rows() {
        if row.len() < 4 {
            continue;
        }
        let code = cell_as_string(&row[0]);
        if !code.starts_with("1201_") {
            continue;
        }

        let name_full = cell_as_string(&row[1]);
        let aux_code = code.replace("1201_", "");
        let raw_name = name_full.trim_start_matches("存货_").trim();
        let parts: Vec<&str> = raw_name.splitn(2, ' ').collect();
        let drug_name = parts[0].to_string();
        let spec = if parts.len() > 1 { parts[1].to_string() } else { String::new() };

        let end_qty = if row.len() > 16 { cell_as_f64(&row[16]) } else { 0.0 };
        let price = if row.len() > 17 { cell_as_f64(&row[17]) } else { 0.0 };

        let entry = LedgerEntry {
            code: code.clone(),
            aux_code,
            name_full,
            drug_name,
            spec,
            price,
            end_qty,
        };

        if code.starts_with("1201_XY") {
            ledger_xy.push(entry);
        } else if code.starts_with("1201_ZY") {
            ledger_zy.push(entry);
        } else if code.starts_with("1201_HC") {
            ledger_hc.push(entry);
        }
    }

    let mut categories = Vec::new();

    // 2. 比对西药房
    if let Some(wp) = west_path {
        let p = Path::new(wp);
        if p.exists() {
            let items = load_warehouse_items(p)?;
            let res = audit_single_category("西药房", &ledger_xy, &items, config);
            categories.push(res);
        }
    }

    // 3. 比对中药房
    if let Some(tp) = tcm_path {
        let p = Path::new(tp);
        if p.exists() {
            let items = load_warehouse_items(p)?;
            let res = audit_single_category("中药房", &ledger_zy, &items, config);
            categories.push(res);
        }
    }

    // 4. 比对耗材库
    if let Some(hp) = hc_path {
        let p = Path::new(hp);
        if p.exists() {
            let items = load_warehouse_items(p)?;
            let res = audit_single_category("耗材库", &ledger_hc, &items, config);
            categories.push(res);
        }
    }

    if categories.is_empty() {
        return Err("未指定有效的库管库存报表（西药房/中药房/耗材库）".into());
    }

    // 5. 汇总全局 KPI
    let grand_total = categories.iter().map(|c| c.summary.total_items).sum();
    let grand_equal = categories.iter().map(|c| c.summary.equal_count).sum();
    let grand_diff = categories.iter().map(|c| c.summary.diff_count).sum();
    let grand_wh_only = categories.iter().map(|c| c.summary.wh_only_count).sum();
    let grand_ledger_only = categories.iter().map(|c| c.summary.ledger_only_count).sum();
    let grand_diff_amt: f64 = categories.iter().map(|c| c.summary.total_diff_amt).sum();
    let overall_rate = if grand_total > 0 {
        ((grand_equal as f64 / grand_total as f64) * 1000.0).round() / 10.0
    } else {
        0.0
    };

    let overall_summary = CategorySummary {
        total_items: grand_total,
        equal_count: grand_equal,
        diff_count: grand_diff,
        wh_only_count: grand_wh_only,
        ledger_only_count: grand_ledger_only,
        match_rate: overall_rate,
        total_diff_amt: (grand_diff_amt * 100.0).round() / 100.0,
    };

    // 6. 使用 rust_xlsxwriter 生成 3 个 Sheet 审计底稿
    let mut out_wb = Workbook::new();

    // 样式
    let fmt_title = Format::new()
        .set_bold()
        .set_font_size(14)
        .set_align(rust_xlsxwriter::FormatAlign::Left)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter);

    let fmt_header = Format::new()
        .set_bold()
        .set_font_size(10)
        .set_background_color(Color::RGB(0x1E293B))
        .set_font_color(Color::RGB(0xFFFFFF))
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_center = Format::new()
        .set_font_size(10)
        .set_align(rust_xlsxwriter::FormatAlign::Center)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    let fmt_cell_left = Format::new()
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

    let fmt_rate = Format::new()
        .set_font_size(10)
        .set_num_format("0.0%")
        .set_align(rust_xlsxwriter::FormatAlign::Right)
        .set_align(rust_xlsxwriter::FormatAlign::VerticalCenter)
        .set_border(FormatBorder::Thin);

    // Sheet 1: 看板总览
    let ws_dash = out_wb.add_worksheet();
    ws_dash.set_name("各库房对账汇总看板").map_err(|e| e.to_string())?;

    ws_dash.write_string_with_format(0, 0, "石家庄心理医院 - 账实库存核对看板", &fmt_title).map_err(|e| e.to_string())?;

    let dash_headers = [
        "库别", "核对品规总数", "完全吻合项", "数量差异项", "仅财务有结存", "仅库管有在库", "账实吻合率", "涉及净差异金额(元)"
    ];
    for (c, h) in dash_headers.iter().enumerate() {
        ws_dash.write_string_with_format(2, c as u16, *h, &fmt_header).map_err(|e| e.to_string())?;
    }

    for (idx, cat) in categories.iter().enumerate() {
        let row_d = (3 + idx) as u32;
        ws_dash.write_string_with_format(row_d, 0, &cat.category, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 1, cat.summary.total_items as f64, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 2, cat.summary.equal_count as f64, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 3, cat.summary.diff_count as f64, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 4, cat.summary.ledger_only_count as f64, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 5, cat.summary.wh_only_count as f64, &fmt_cell_center).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 6, cat.summary.match_rate / 100.0, &fmt_rate).map_err(|e| e.to_string())?;
        ws_dash.write_number_with_format(row_d, 7, cat.summary.total_diff_amt, &fmt_money).map_err(|e| e.to_string())?;
    }

    for c in 0..8 {
        ws_dash.set_column_width(c, 16).map_err(|e| e.to_string())?;
    }

    // Sheet 2: 差异重点排查清单
    let ws_diff = out_wb.add_worksheet();
    ws_diff.set_name("差异重点排查清单").map_err(|e| e.to_string())?;

    let item_headers = [
        "库房", "核对状态", "药品/材料名称", "规格", "生产厂家", "单位", "财务存货编码",
        "财务结存数量", "库管在库数量", "数量差异", "财务金额", "库管金额", "差异金额"
    ];

    for (c, h) in item_headers.iter().enumerate() {
        ws_diff.write_string_with_format(0, c as u16, *h, &fmt_header).map_err(|e| e.to_string())?;
    }

    let mut row_diff = 1;
    for cat in &categories {
        for r in &cat.records {
            if r.status == "EQUAL" {
                continue;
            }
            ws_diff.write_string_with_format(row_diff, 0, &r.category, &fmt_cell_center).map_err(|e| e.to_string())?;
            ws_diff.write_string_with_format(row_diff, 1, &r.status_desc, &fmt_cell_center).map_err(|e| e.to_string())?;
            ws_diff.write_string_with_format(row_diff, 2, &r.name, &fmt_cell_left).map_err(|e| e.to_string())?;
            ws_diff.write_string_with_format(row_diff, 3, &r.spec, &fmt_cell_left).map_err(|e| e.to_string())?;
            ws_diff.write_string_with_format(row_diff, 4, &r.factory, &fmt_cell_left).map_err(|e| e.to_string())?;
            ws_diff.write_string_with_format(row_diff, 5, &r.unit, &fmt_cell_center).map_err(|e| e.to_string())?;
            ws_diff.write_string_with_format(row_diff, 6, &r.ledger_code, &fmt_cell_center).map_err(|e| e.to_string())?;

            ws_diff.write_number_with_format(row_diff, 7, r.ledger_qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws_diff.write_number_with_format(row_diff, 8, r.wh_qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws_diff.write_number_with_format(row_diff, 9, r.diff_qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws_diff.write_number_with_format(row_diff, 10, r.ledger_amt, &fmt_money).map_err(|e| e.to_string())?;
            ws_diff.write_number_with_format(row_diff, 11, r.wh_amt, &fmt_money).map_err(|e| e.to_string())?;
            ws_diff.write_number_with_format(row_diff, 12, r.diff_amt, &fmt_money).map_err(|e| e.to_string())?;

            row_diff += 1;
        }
    }

    for c in 0..13 {
        ws_diff.set_column_width(c, 14).map_err(|e| e.to_string())?;
    }
    ws_diff.set_column_width(2, 24).map_err(|e| e.to_string())?;
    ws_diff.set_column_width(3, 16).map_err(|e| e.to_string())?;
    ws_diff.set_column_width(4, 20).map_err(|e| e.to_string())?;

    // Sheet 3: 全部品规对照底稿
    let ws_all = out_wb.add_worksheet();
    ws_all.set_name("全部品规对照底稿").map_err(|e| e.to_string())?;

    for (c, h) in item_headers.iter().enumerate() {
        ws_all.write_string_with_format(0, c as u16, *h, &fmt_header).map_err(|e| e.to_string())?;
    }

    let mut row_all = 1;
    for cat in &categories {
        for r in &cat.records {
            ws_all.write_string_with_format(row_all, 0, &r.category, &fmt_cell_center).map_err(|e| e.to_string())?;
            ws_all.write_string_with_format(row_all, 1, &r.status_desc, &fmt_cell_center).map_err(|e| e.to_string())?;
            ws_all.write_string_with_format(row_all, 2, &r.name, &fmt_cell_left).map_err(|e| e.to_string())?;
            ws_all.write_string_with_format(row_all, 3, &r.spec, &fmt_cell_left).map_err(|e| e.to_string())?;
            ws_all.write_string_with_format(row_all, 4, &r.factory, &fmt_cell_left).map_err(|e| e.to_string())?;
            ws_all.write_string_with_format(row_all, 5, &r.unit, &fmt_cell_center).map_err(|e| e.to_string())?;
            ws_all.write_string_with_format(row_all, 6, &r.ledger_code, &fmt_cell_center).map_err(|e| e.to_string())?;

            ws_all.write_number_with_format(row_all, 7, r.ledger_qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws_all.write_number_with_format(row_all, 8, r.wh_qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws_all.write_number_with_format(row_all, 9, r.diff_qty, &fmt_qty).map_err(|e| e.to_string())?;
            ws_all.write_number_with_format(row_all, 10, r.ledger_amt, &fmt_money).map_err(|e| e.to_string())?;
            ws_all.write_number_with_format(row_all, 11, r.wh_amt, &fmt_money).map_err(|e| e.to_string())?;
            ws_all.write_number_with_format(row_all, 12, r.diff_amt, &fmt_money).map_err(|e| e.to_string())?;

            row_all += 1;
        }
    }

    for c in 0..13 {
        ws_all.set_column_width(c, 14).map_err(|e| e.to_string())?;
    }
    ws_all.set_column_width(2, 24).map_err(|e| e.to_string())?;
    ws_all.set_column_width(3, 16).map_err(|e| e.to_string())?;
    ws_all.set_column_width(4, 20).map_err(|e| e.to_string())?;

    let out_path_buf = if let Some(co) = custom_output {
        PathBuf::from(co)
    } else {
        let parent = ledger_p.parent().unwrap_or_else(|| Path::new("."));
        parent.join("账实库存核对分析报告_已生成.xlsx")
    };

    out_wb
        .save(&out_path_buf)
        .map_err(|e| format!("保存审计报告 Excel 失败: {}", e))?;

    Ok(OverallAuditResult {
        success: true,
        output_file: out_path_buf.to_string_lossy().to_string(),
        overall: overall_summary,
        categories,
        error: None,
    })
}
