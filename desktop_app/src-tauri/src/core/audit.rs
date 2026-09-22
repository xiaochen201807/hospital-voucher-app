use crate::core::config::ConfigData;
use crate::core::excel_utils::{
    amount_or_product, cents_to_currency, currency_cents, find_col_idx, find_header_row,
    is_summary_row, optional_col, read_sheet_rows, required_col, row_as_f64, row_as_string,
};
pub use crate::core::ledger::load_categorized_ledger;
use crate::core::matching::{build_ledger_candidates, match_drug_with_method};
use crate::core::models::LedgerEntry;
use rust_xlsxwriter::{Color, Format, FormatBorder, Workbook, Worksheet};
use serde::{Deserialize, Serialize};
use std::collections::HashSet;
use std::path::{Path, PathBuf};

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct AuditRecord {
    pub category: String,
    pub status: String,      // EQUAL, DIFF, WH_ONLY, LEDGER_ONLY
    pub status_desc: String, // 数量和金额吻合/存在差异, 仅财务有结存, 仅库管有在库
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
pub struct WarehouseItem {
    pub name: String,
    pub spec: String,
    pub dosage_form: String,
    pub factory: String,
    pub unit: String,
    pub qty: f64,
    pub price: f64,
    pub amount: f64,
}

// 保留旧的模块路径，避免外部调用方因公共类型迁移而立即失效。
pub use crate::core::models::LedgerCandidateOption;

/// 第一步：智能识别映射项（前端预览与确认）
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct AuditMappingItem {
    pub id: usize,
    pub category: String, // "西药房", "中药房", "耗材库"
    pub wh_name: String,
    pub wh_spec: String,
    pub wh_factory: String,
    pub wh_unit: String,
    pub wh_qty: f64,
    pub wh_price: f64,
    pub wh_amount: f64,
    pub matched: bool,
    pub checked: bool, // 自动识别后默认自动勾选上！
    pub match_method: String,
    pub ledger_code: String,
    pub ledger_name: String,
    pub ledger_qty: f64,
    pub ledger_price: f64,
    pub ledger_amount: f64,
    pub candidates: Vec<LedgerCandidateOption>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct AuditMappingPreviewResult {
    pub success: bool,
    pub total_items: usize,
    pub matched_count: usize,
    pub unmatched_count: usize,
    pub match_rate: f64,
    pub items: Vec<AuditMappingItem>,
    #[serde(default)]
    pub error: Option<String>,
}

/// 前端用户确认后的映射关系项
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ConfirmedAuditMappingItem {
    pub id: usize,
    pub category: String,
    pub wh_name: String,
    pub wh_spec: String,
    pub wh_factory: String,
    pub wh_unit: String,
    pub wh_qty: f64,
    pub wh_price: f64,
    pub wh_amount: f64,
    pub checked: bool,
    pub ledger_code: String,
}

/// 读取库管系统报表（西药、中药、耗材）
pub(crate) fn load_warehouse_items(path: &Path) -> Result<Vec<WarehouseItem>, String> {
    let (_, rows) = read_sheet_rows(path, &[], "库管报表")?;
    if rows.len() < 2 {
        return Ok(Vec::new());
    }

    let h_idx = find_header_row(
        &rows,
        &["药品名称", "品名", "材料名称", "名称", "耗材名称"],
        10,
        "库管报表",
    )?;
    let header = &rows[h_idx];
    let col_name = required_col(
        header,
        &["药品名称", "品名", "材料名称", "名称", "耗材名称"],
        "药品/材料名称",
    )?;
    let col_spec = optional_col(header, &["规格"], col_name + 1);
    let col_dosage_form = find_col_idx(header, &["剂型", "剂型名称", "剂型类别"]);
    let col_factory = optional_col(
        header,
        &["制药厂", "生产厂家", "厂家", "生产商"],
        col_dosage_form.unwrap_or(col_spec) + 1,
    );
    let col_unit = optional_col(header, &["单位"], col_factory + 1);
    let col_qty = optional_col(
        header,
        &["数量", "结存数量", "在库数量", "库存数量"],
        col_unit + 1,
    );
    let col_price = optional_col(header, &["单价", "成本价", "结存单价"], col_qty + 1);
    let col_amt = optional_col(header, &["金额", "结存金额", "成本金额"], col_price + 1);

    let mut items = Vec::new();
    for row in rows.iter().skip(h_idx + 1) {
        if row.is_empty() {
            continue;
        }
        let name = row_as_string(row, col_name);
        if is_summary_row(&name) {
            continue;
        }

        let spec = row_as_string(row, col_spec);
        let dosage_form = col_dosage_form
            .map(|col| row_as_string(row, col))
            .unwrap_or_default();
        let factory = row_as_string(row, col_factory);
        let unit = row_as_string(row, col_unit);
        let qty = row_as_f64(row, col_qty);
        let price = row_as_f64(row, col_price);
        let amount = cents_to_currency(currency_cents(amount_or_product(
            row_as_f64(row, col_amt),
            qty,
            price,
        )));

        items.push(WarehouseItem {
            name,
            spec,
            dosage_form,
            factory,
            unit,
            qty,
            price,
            amount,
        });
    }

    Ok(items)
}

fn build_mapping_item(
    id: usize,
    category: &str,
    w: WarehouseItem,
    ledger_entries: &[LedgerEntry],
    config: &ConfigData,
) -> AuditMappingItem {
    let matched_res = match_drug_with_method(&w.name, &w.spec, &w.factory, ledger_entries, config);

    let relevant_cands = build_ledger_candidates(&w.name, ledger_entries);

    let (
        matched,
        checked,
        match_method,
        ledger_code,
        ledger_name,
        ledger_qty,
        ledger_price,
        ledger_amount,
    ) = if let Some((entry, method)) = matched_res {
        let ledger_amount = cents_to_currency(currency_cents(entry.end_amount));
        (
            true,
            true,
            method,
            entry.code,
            entry.name_full,
            entry.end_qty,
            entry.price,
            ledger_amount,
        )
    } else {
        (
            false,
            false,
            "未自动匹配".to_string(),
            String::new(),
            String::new(),
            0.0,
            0.0,
            0.0,
        )
    };

    AuditMappingItem {
        id,
        category: category.to_string(),
        wh_name: w.name,
        wh_spec: w.spec,
        wh_factory: w.factory,
        wh_unit: w.unit,
        wh_qty: w.qty,
        wh_price: w.price,
        wh_amount: w.amount,
        matched,
        checked,
        match_method,
        ledger_code,
        ledger_name,
        ledger_qty,
        ledger_price,
        ledger_amount,
        candidates: relevant_cands,
    }
}

/// 第一步：智能识别多库与总账映射关系（供前端展示与确认）
pub fn preview_inventory_audit_mapping(
    ledger_path: &str,
    west_path: Option<&str>,
    tcm_path: Option<&str>,
    hc_path: Option<&str>,
    config: &ConfigData,
) -> Result<AuditMappingPreviewResult, String> {
    let ledger_p = Path::new(ledger_path);
    if !ledger_p.exists() {
        return Err(format!("财务总账文件 '{:?}' 不存在", ledger_p));
    }

    let (ledger_xy, ledger_zy, ledger_hc) = load_categorized_ledger(ledger_p)?;

    let mut all_items = Vec::new();
    let mut next_id = 1;

    let warehouse_defs: [(&str, Option<&str>, &[LedgerEntry]); 3] = [
        ("西药房", west_path, &ledger_xy),
        ("中药房", tcm_path, &ledger_zy),
        ("耗材库", hc_path, &ledger_hc),
    ];
    for (category, warehouse_path, ledger_entries) in warehouse_defs {
        let Some(warehouse_path) = warehouse_path else {
            continue;
        };
        let path = Path::new(warehouse_path);
        if !path.exists() {
            continue;
        }

        for warehouse_item in load_warehouse_items(path)? {
            all_items.push(build_mapping_item(
                next_id,
                category,
                warehouse_item,
                ledger_entries,
                config,
            ));
            next_id += 1;
        }
    }

    if all_items.is_empty() {
        return Err("未指定有效的库管库存报表（西药房/中药房/耗材库）".into());
    }

    let total_items = all_items.len();
    let matched_count = all_items.iter().filter(|it| it.matched).count();
    let unmatched_count = total_items - matched_count;
    let match_rate = if total_items > 0 {
        ((matched_count as f64 / total_items as f64) * 1000.0).round() / 10.0
    } else {
        0.0
    };

    Ok(AuditMappingPreviewResult {
        success: true,
        total_items,
        matched_count,
        unmatched_count,
        match_rate,
        items: all_items,
        error: None,
    })
}

/// 第二步：根据用户确认的勾选与映射关系，执行比对并生成审计分析报告 Excel
pub fn execute_inventory_audit_with_mapping(
    ledger_path: &str,
    confirmed_items: Vec<ConfirmedAuditMappingItem>,
    custom_output: Option<&str>,
    _config: &ConfigData,
) -> Result<OverallAuditResult, String> {
    let ledger_p = Path::new(ledger_path);
    if !ledger_p.exists() {
        return Err(format!("财务总账文件 '{:?}' 不存在", ledger_p));
    }

    let (ledger_xy, ledger_zy, ledger_hc) = load_categorized_ledger(ledger_p)?;

    let mut categories = Vec::new();
    let category_defs = [
        ("西药房", &ledger_xy),
        ("中药房", &ledger_zy),
        ("耗材库", &ledger_hc),
    ];

    for (cat_name, ledger_list) in &category_defs {
        let cat_items: Vec<&ConfirmedAuditMappingItem> = confirmed_items
            .iter()
            .filter(|it| it.category == *cat_name)
            .collect();

        if cat_items.is_empty() {
            continue;
        }

        let mut records = Vec::new();
        let mut matched_ledger_codes = HashSet::new();

        // 1. 处理库管项目
        for it in cat_items {
            let mut matched_entry: Option<&LedgerEntry> = None;
            if it.checked && !it.ledger_code.is_empty() {
                matched_entry = ledger_list.iter().find(|l| l.code == it.ledger_code);
            }

            if let Some(m) = matched_entry {
                matched_ledger_codes.insert(m.code.clone());
                let l_qty = m.end_qty;
                let l_price = m.price;
                // 总账期末金额可能包含历史尾差，不能重新用“数量 × 单价”覆盖。
                let l_amt = cents_to_currency(currency_cents(m.end_amount));
                let diff_qty = ((l_qty - it.wh_qty) * 100.0).round() / 100.0;
                let diff_amt = cents_to_currency(currency_cents(l_amt - it.wh_amount));

                let qty_equal = diff_qty.abs() < 0.001;
                let amount_equal = diff_amt.abs() < 0.01;
                let (status, status_desc) = if qty_equal && amount_equal {
                    ("EQUAL".to_string(), "数量和金额均吻合".to_string())
                } else if !qty_equal && !amount_equal {
                    ("DIFF".to_string(), "数量和金额均有差异".to_string())
                } else if !qty_equal {
                    ("DIFF".to_string(), "存在数量差异".to_string())
                } else {
                    ("DIFF".to_string(), "存在金额差异".to_string())
                };

                records.push(AuditRecord {
                    category: cat_name.to_string(),
                    status,
                    status_desc,
                    name: it.wh_name.clone(),
                    spec: it.wh_spec.clone(),
                    factory: it.wh_factory.clone(),
                    unit: it.wh_unit.clone(),
                    ledger_code: m.code.clone(),
                    ledger_name: m.name_full.clone(),
                    ledger_qty: l_qty,
                    wh_qty: it.wh_qty,
                    diff_qty,
                    ledger_price: l_price,
                    wh_price: it.wh_price,
                    ledger_amt: l_amt,
                    wh_amt: it.wh_amount,
                    diff_amt,
                });
            } else {
                records.push(AuditRecord {
                    category: cat_name.to_string(),
                    status: "WH_ONLY".to_string(),
                    status_desc: "仅库管有在库".to_string(),
                    name: it.wh_name.clone(),
                    spec: it.wh_spec.clone(),
                    factory: it.wh_factory.clone(),
                    unit: it.wh_unit.clone(),
                    ledger_code: "-".to_string(),
                    ledger_name: "-".to_string(),
                    ledger_qty: 0.0,
                    wh_qty: it.wh_qty,
                    diff_qty: -it.wh_qty,
                    ledger_price: 0.0,
                    wh_price: it.wh_price,
                    ledger_amt: 0.0,
                    wh_amt: it.wh_amount,
                    diff_amt: -it.wh_amount,
                });
            }
        }

        // 2. 处理仅财务有结存的项目
        for l in *ledger_list {
            if !matched_ledger_codes.contains(&l.code) {
                let l_amt = cents_to_currency(currency_cents(l.end_amount));
                records.push(AuditRecord {
                    category: cat_name.to_string(),
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
        let diff_count = records.iter().filter(|r| r.status == "DIFF").count();
        let wh_only_count = records.iter().filter(|r| r.status == "WH_ONLY").count();
        let ledger_only_count = records.iter().filter(|r| r.status == "LEDGER_ONLY").count();
        let total_diff_amt_cents: i64 = records
            .iter()
            .map(|record| currency_cents(record.diff_amt))
            .sum();
        let match_rate = if total_items > 0 {
            ((equal_count as f64 / total_items as f64) * 1000.0).round() / 10.0
        } else {
            0.0
        };

        categories.push(CategoryAuditResult {
            category: cat_name.to_string(),
            summary: CategorySummary {
                total_items,
                equal_count,
                diff_count,
                wh_only_count,
                ledger_only_count,
                match_rate,
                total_diff_amt: cents_to_currency(total_diff_amt_cents),
            },
            records,
        });
    }

    if categories.is_empty() {
        return Err("核对结果为空，请确认是否提供了有效的在库数据。".into());
    }

    // 汇总全局 KPI
    let grand_total = categories.iter().map(|c| c.summary.total_items).sum();
    let grand_equal = categories.iter().map(|c| c.summary.equal_count).sum();
    let grand_diff = categories.iter().map(|c| c.summary.diff_count).sum();
    let grand_wh_only = categories.iter().map(|c| c.summary.wh_only_count).sum();
    let grand_ledger_only = categories.iter().map(|c| c.summary.ledger_only_count).sum();
    let grand_diff_amt_cents: i64 = categories
        .iter()
        .map(|category| currency_cents(category.summary.total_diff_amt))
        .sum();
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
        total_diff_amt: cents_to_currency(grand_diff_amt_cents),
    };

    let out_path_buf = if let Some(co) = custom_output {
        PathBuf::from(co)
    } else {
        let parent = ledger_p.parent().unwrap_or_else(|| Path::new("."));
        parent.join("账实库存核对分析报告_已生成.xlsx")
    };

    generate_audit_report_excel(&categories, &out_path_buf)?;

    Ok(OverallAuditResult {
        success: true,
        output_file: out_path_buf.to_string_lossy().to_string(),
        overall: overall_summary,
        categories,
        error: None,
    })
}

/// 执行多库账实核对（兼容原有直接核对调用）
pub fn run_inventory_audit(
    ledger_path: &str,
    west_path: Option<&str>,
    tcm_path: Option<&str>,
    hc_path: Option<&str>,
    custom_output: Option<&str>,
    config: &ConfigData,
) -> Result<OverallAuditResult, String> {
    let preview =
        preview_inventory_audit_mapping(ledger_path, west_path, tcm_path, hc_path, config)?;
    let confirmed: Vec<ConfirmedAuditMappingItem> = preview
        .items
        .into_iter()
        .map(|it| ConfirmedAuditMappingItem {
            id: it.id,
            category: it.category,
            wh_name: it.wh_name,
            wh_spec: it.wh_spec,
            wh_factory: it.wh_factory,
            wh_unit: it.wh_unit,
            wh_qty: it.wh_qty,
            wh_price: it.wh_price,
            wh_amount: it.wh_amount,
            checked: it.matched,
            ledger_code: it.ledger_code,
        })
        .collect();

    execute_inventory_audit_with_mapping(ledger_path, confirmed, custom_output, config)
}

const AUDIT_ITEM_HEADERS: [&str; 13] = [
    "库房",
    "核对状态",
    "药品/材料名称",
    "规格",
    "生产厂家",
    "单位",
    "财务存货编码",
    "财务结存数量",
    "库管在库数量",
    "数量差异",
    "财务金额",
    "库管金额",
    "差异金额",
];

fn write_audit_item_headers(worksheet: &mut Worksheet, format: &Format) -> Result<(), String> {
    for (col_idx, header) in AUDIT_ITEM_HEADERS.iter().enumerate() {
        worksheet
            .write_string_with_format(0, col_idx as u16, *header, format)
            .map_err(|e| e.to_string())?;
    }
    Ok(())
}

fn write_audit_record_row(
    worksheet: &mut Worksheet,
    row: u32,
    record: &AuditRecord,
    center_format: &Format,
    left_format: &Format,
    qty_format: &Format,
    money_format: &Format,
) -> Result<(), String> {
    worksheet
        .write_string_with_format(row, 0, &record.category, center_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 1, &record.status_desc, center_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 2, &record.name, left_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 3, &record.spec, left_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 4, &record.factory, left_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 5, &record.unit, center_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_string_with_format(row, 6, &record.ledger_code, center_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 7, record.ledger_qty, qty_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 8, record.wh_qty, qty_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 9, record.diff_qty, qty_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 10, record.ledger_amt, money_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 11, record.wh_amt, money_format)
        .map_err(|e| e.to_string())?;
    worksheet
        .write_number_with_format(row, 12, record.diff_amt, money_format)
        .map_err(|e| e.to_string())?;
    Ok(())
}

fn set_audit_item_column_widths(worksheet: &mut Worksheet) -> Result<(), String> {
    for col in 0..13 {
        worksheet
            .set_column_width(col, 14)
            .map_err(|e| e.to_string())?;
    }
    worksheet
        .set_column_width(2, 24)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(3, 16)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(4, 20)
        .map_err(|e| e.to_string())?;
    Ok(())
}

/// 输出标准 3-Sheet Excel 审计分析报告
fn generate_audit_report_excel(
    categories: &[CategoryAuditResult],
    out_path: &Path,
) -> Result<(), String> {
    let mut out_wb = Workbook::new();

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
    ws_dash
        .set_name("各库房对账汇总看板")
        .map_err(|e| e.to_string())?;

    ws_dash
        .write_string_with_format(0, 0, "石家庄心理医院 - 账实库存核对看板", &fmt_title)
        .map_err(|e| e.to_string())?;

    let dash_headers = [
        "库别",
        "核对品规总数",
        "完全吻合项",
        "数量或金额差异项",
        "仅财务有结存",
        "仅库管有在库",
        "账实吻合率",
        "涉及净差异金额(元)",
    ];
    for (c, h) in dash_headers.iter().enumerate() {
        ws_dash
            .write_string_with_format(2, c as u16, *h, &fmt_header)
            .map_err(|e| e.to_string())?;
    }

    for (idx, cat) in categories.iter().enumerate() {
        let row_d = (3 + idx) as u32;
        ws_dash
            .write_string_with_format(row_d, 0, &cat.category, &fmt_cell_center)
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(row_d, 1, cat.summary.total_items as f64, &fmt_cell_center)
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(row_d, 2, cat.summary.equal_count as f64, &fmt_cell_center)
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(row_d, 3, cat.summary.diff_count as f64, &fmt_cell_center)
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(
                row_d,
                4,
                cat.summary.ledger_only_count as f64,
                &fmt_cell_center,
            )
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(row_d, 5, cat.summary.wh_only_count as f64, &fmt_cell_center)
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(row_d, 6, cat.summary.match_rate / 100.0, &fmt_rate)
            .map_err(|e| e.to_string())?;
        ws_dash
            .write_number_with_format(row_d, 7, cat.summary.total_diff_amt, &fmt_money)
            .map_err(|e| e.to_string())?;
    }

    for c in 0..8 {
        ws_dash.set_column_width(c, 16).map_err(|e| e.to_string())?;
    }

    // Sheet 2: 差异重点排查清单
    let ws_diff = out_wb.add_worksheet();
    ws_diff
        .set_name("差异重点排查清单")
        .map_err(|e| e.to_string())?;

    write_audit_item_headers(ws_diff, &fmt_header)?;

    let mut row_diff = 1;
    for cat in categories {
        for r in &cat.records {
            if r.status == "EQUAL" {
                continue;
            }
            write_audit_record_row(
                ws_diff,
                row_diff,
                r,
                &fmt_cell_center,
                &fmt_cell_left,
                &fmt_qty,
                &fmt_money,
            )?;

            row_diff += 1;
        }
    }

    set_audit_item_column_widths(ws_diff)?;

    // Sheet 3: 全部品规对照底稿
    let ws_all = out_wb.add_worksheet();
    ws_all
        .set_name("全部品规对照底稿")
        .map_err(|e| e.to_string())?;

    write_audit_item_headers(ws_all, &fmt_header)?;

    let mut row_all = 1;
    for cat in categories {
        for r in &cat.records {
            write_audit_record_row(
                ws_all,
                row_all,
                r,
                &fmt_cell_center,
                &fmt_cell_left,
                &fmt_qty,
                &fmt_money,
            )?;

            row_all += 1;
        }
    }

    set_audit_item_column_widths(ws_all)?;

    out_wb
        .save(out_path)
        .map_err(|e| format!("保存审计报告 Excel 失败: {}", e))?;

    Ok(())
}
