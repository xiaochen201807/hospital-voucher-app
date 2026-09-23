use crate::core::audit::{load_warehouse_items, WarehouseItem};
use crate::core::config::ConfigData;
use crate::core::excel_utils::{
    cell_as_f64, cell_as_string, find_col_idx, open_excel, read_sheet_rows,
};
use calamine::{Data, Reader};
use chrono::{Datelike, Local, NaiveDate};
use regex::Regex;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::path::{Path, PathBuf};

use crate::core::models::LedgerCandidateOption;

/// 归一化文本：去除所有空白字符，统一全角半角括号
pub fn normalize_text(text: &str) -> String {
    let s = text.trim();
    let mut s: String = s.chars().filter(|c| !c.is_whitespace()).collect();
    s = s.replace('（', "(").replace('）', ")");
    s = s.replace('【', "[").replace('】', "]");
    s = s.replace('×', "*").replace('x', "*").replace('X', "*");
    s
}

/// 清洗药品名称，去除括号内容获取核心通用名
pub fn clean_drug_name(text: &str) -> String {
    let norm = normalize_text(text);
    if let Ok(re) = Regex::new(r"\(.*?\)|\{.*?\}|\[.*?\]") {
        re.replace_all(&norm, "").to_string()
    } else {
        norm
    }
}

/// 推导当月或指定月份最后一天
pub fn detect_month_end_date(custom_date: Option<&str>, source_hint: Option<&str>) -> String {
    if let Some(d) = custom_date {
        let trimmed = d.trim();
        if !trimmed.is_empty() {
            return trimmed.to_string();
        }
    }

    if let Some(hint) = source_hint {
        if let Ok(re) = Regex::new(r"(\d{4})[^\d]*?(\d{1,2})月?") {
            if let Some(caps) = re.captures(hint) {
                if let (Ok(year), Ok(month)) = (caps[1].parse::<i32>(), caps[2].parse::<u32>()) {
                    if (1..=12).contains(&month) {
                        let next_month = if month == 12 { 1 } else { month + 1 };
                        let next_year = if month == 12 { year + 1 } else { year };
                        if let Some(first_of_next) =
                            NaiveDate::from_ymd_opt(next_year, next_month, 1)
                        {
                            if let Some(last_day) = first_of_next.pred_opt() {
                                return last_day.format("%Y-%m-%d").to_string();
                            }
                        }
                    }
                }
            }
        }
    }

    let today = Local::now().date_naive();
    let year = today.year();
    let month = today.month();
    let next_month = if month == 12 { 1 } else { month + 1 };
    let next_year = if month == 12 { year + 1 } else { year };
    if let Some(first_of_next) = NaiveDate::from_ymd_opt(next_year, next_month, 1) {
        if let Some(last_day) = first_of_next.pred_opt() {
            return last_day.format("%Y-%m-%d").to_string();
        }
    }

    today.format("%Y-%m-%d").to_string()
}

// ----------------------------------------------------
// 1. 数据结构定义
// ----------------------------------------------------

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ExternalVoucherRow {
    pub row_no: usize,
    pub date: String,
    pub voucher_type: String,
    pub voucher_no: String,
    pub summary: String,
    pub subject_code: String,
    pub subject_name: String,
    pub debit_amount: Option<f64>,
    pub credit_amount: Option<f64>,
    pub exchange_rate: Option<f64>,
    pub currency: String,
    pub qty: Option<f64>,
    pub aux_code: String,
    pub aux_name: String,
    pub settle_method: String,
    pub bill_no: String,
    pub occur_date: String,
    pub is_credit: bool,
    pub is_unmatched: bool,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExternalUnmatchedDrug {
    pub id: usize,
    pub row_index: usize,
    pub name: String,
    pub factory: String,
    pub spec: String,
    pub qty: f64,
    pub amount: f64,
    pub supplier: Option<String>,
    pub candidates: Vec<LedgerCandidateOption>,
}

/// 外账凭证生成时由前端确认的人工存货辅助编码。
///
/// `row_index` 使用原始入库单/销售表中的 Excel 行号，因此同一药品即使在
/// 不同来源行出现，也不会因为名称相同而误覆盖另一行的人工选择。
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ConfirmedExternalInventoryMapping {
    pub row_index: usize,
    pub aux_code: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExternalUnmatchedSupplier {
    pub supplier: String,
    pub item_count: usize,
    pub amount: f64,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExternalInboundResult {
    pub success: bool,
    pub output_file: String,
    pub voucher_date: String,
    pub voucher_no: String,
    pub total_items: usize,
    pub supplier_count: usize,
    pub total_entries: usize,
    pub matched_count: usize,
    pub unmatched_count: usize,
    pub unmatched_supplier_count: usize,
    pub match_rate: f64,
    pub total_debit: f64,
    pub total_credit: f64,
    pub diff: f64,
    pub is_balanced: bool,
    pub unmatched_drugs: Vec<ExternalUnmatchedDrug>,
    pub unmatched_suppliers: Vec<ExternalUnmatchedSupplier>,
    pub voucher_rows: Vec<ExternalVoucherRow>,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExternalOutboundResult {
    pub success: bool,
    pub output_file: String,
    pub voucher_date: String,
    pub voucher_no: String,
    pub total_items: usize,
    pub total_entries: usize,
    pub matched_count: usize,
    pub unmatched_count: usize,
    pub match_rate: f64,
    pub total_debit: f64,
    pub total_credit: f64,
    pub diff: f64,
    pub is_balanced: bool,
    pub unmatched_drugs: Vec<ExternalUnmatchedDrug>,
    pub voucher_rows: Vec<ExternalVoucherRow>,
}

#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ExternalAuditRecord {
    pub aux_code: String,
    pub name: String,
    pub spec: String,
    pub factory: String,
    pub ext_qty: f64,
    pub ext_price: f64,
    pub ext_amount: f64,
    pub wh_qty: f64,
    pub diff_qty: f64,
    pub status: String,
}

#[derive(Debug, Serialize, Deserialize)]
pub struct ExternalAuditResult {
    pub success: bool,
    pub output_file: String,
    pub total_items: usize,
    pub equal_count: usize,
    pub diff_count: usize,
    pub ext_only_count: usize,
    pub wh_only_count: usize,
    pub match_rate: f64,
    pub records: Vec<ExternalAuditRecord>,
}

#[derive(Debug, Clone)]
struct ExternalInventoryItem {
    code: String,
    name: String,
    spec: String,
    spec_key: String,
    norm_spec: String,
    norm_name: String,
    clean_name: String,
}

#[derive(Debug, Clone, Default)]
struct ExternalTemplateData {
    supplier_map: HashMap<String, (String, String)>,
    inventory_items: Vec<ExternalInventoryItem>,
    inventory_by_code: HashMap<String, ExternalInventoryItem>,
    balance_by_code: HashMap<String, (f64, f64, f64)>,
}

// ----------------------------------------------------
// 2. 辅助字典与模板解析
// ----------------------------------------------------

fn load_external_template(template_path: &Path) -> Result<ExternalTemplateData, String> {
    let mut excel = open_excel(template_path)?;
    let mut data = ExternalTemplateData::default();

    let info_sheet_name = excel
        .sheet_names()
        .into_iter()
        .find(|s| s.contains("辅助信息"))
        .ok_or_else(|| "外账模板中未找到【辅助信息】Sheet".to_string())?;

    let info_range = excel
        .worksheet_range(&info_sheet_name)
        .map_err(|e| format!("读取【辅助信息】失败: {}", e))?;

    let mut col_type = None;
    let mut col_code = None;
    let mut col_name = None;
    let mut col_spec = None;

    let rows: Vec<Vec<Data>> = info_range.rows().map(|r| r.to_vec()).collect();
    for (r_idx, row) in rows.iter().enumerate().take(5) {
        for (c_idx, cell) in row.iter().enumerate() {
            let s = cell_as_string(cell);
            if s.contains("辅助类型") {
                col_type = Some(c_idx);
            } else if s.contains("辅助编码") {
                col_code = Some(c_idx);
            } else if s.contains("辅助名称") {
                col_name = Some(c_idx);
            } else if s.contains("规格") {
                col_spec = Some(c_idx);
            }
        }
        if col_type.is_some() && col_code.is_some() && col_name.is_some() {
            for data_row in rows.iter().skip(r_idx + 1) {
                let aux_type = col_type
                    .and_then(|c| data_row.get(c))
                    .map(cell_as_string)
                    .unwrap_or_default();
                let aux_code = col_code
                    .and_then(|c| data_row.get(c))
                    .map(cell_as_string)
                    .unwrap_or_default()
                    .trim()
                    .to_string();
                let aux_name = col_name
                    .and_then(|c| data_row.get(c))
                    .map(cell_as_string)
                    .unwrap_or_default()
                    .trim()
                    .to_string();
                let aux_spec = col_spec
                    .and_then(|c| data_row.get(c))
                    .map(cell_as_string)
                    .unwrap_or_default()
                    .trim()
                    .to_string();

                if aux_code.is_empty() || aux_name.is_empty() {
                    continue;
                }

                let formatted_code = if let Ok(num) = aux_code.parse::<i64>() {
                    format!("{:05}", num)
                } else {
                    aux_code
                };

                if aux_type.contains("供应商") {
                    let norm = normalize_text(&aux_name);
                    data.supplier_map
                        .insert(norm.clone(), (formatted_code.clone(), aux_name.clone()));
                    let brief = norm
                        .replace("有限公司", "")
                        .replace("股份有限公司", "")
                        .replace("有限责任公司", "")
                        .replace("医药", "")
                        .replace("石家庄", "");
                    if !brief.is_empty() {
                        data.supplier_map
                            .insert(brief, (formatted_code.clone(), aux_name.clone()));
                    }
                } else if aux_type.contains("存货") {
                    let norm_name = normalize_text(&aux_name);
                    let clean = clean_drug_name(&aux_name);
                    let item = ExternalInventoryItem {
                        code: formatted_code.clone(),
                        name: aux_name.clone(),
                        spec: aux_spec.clone(),
                        spec_key: normalize_inventory_spec_key(&aux_spec),
                        norm_spec: normalize_text(&aux_spec),
                        norm_name,
                        clean_name: clean,
                    };
                    data.inventory_items.push(item.clone());
                    data.inventory_by_code.insert(formatted_code, item);
                }
            }
            break;
        }
    }

    if let Some(bal_sheet_name) = excel
        .sheet_names()
        .into_iter()
        .find(|s| s.contains("辅助余额表"))
    {
        if let Ok(bal_range) = excel.worksheet_range(&bal_sheet_name) {
            let bal_rows: Vec<Vec<Data>> = bal_range.rows().map(|r| r.to_vec()).collect();
            for row in bal_rows.iter().skip(1) {
                let code_cell = row.get(0).map(cell_as_string).unwrap_or_default();
                if code_cell.trim().starts_with("1405") {
                    let aux_code_raw = row.get(3).map(cell_as_string).unwrap_or_default();
                    if aux_code_raw.is_empty() {
                        continue;
                    }
                    let aux_code = if let Ok(n) = aux_code_raw.trim().parse::<i64>() {
                        format!("{:05}", n)
                    } else {
                        aux_code_raw.trim().to_string()
                    };

                    let price = row.get(7).map(cell_as_f64).unwrap_or(0.0);
                    let qty = row.get(8).map(cell_as_f64).unwrap_or(0.0);
                    let amount = row.get(10).map(cell_as_f64).unwrap_or(0.0);

                    data.balance_by_code.insert(aux_code, (price, qty, amount));
                }
            }
        }
    }

    Ok(data)
}

fn find_external_group_column(row: &[Data], keyword: &str) -> Option<usize> {
    row.iter()
        .enumerate()
        .find_map(|(idx, cell)| cell_as_string(cell).contains(keyword).then_some(idx))
}

fn find_external_subheader_column(row: &[Data], start: usize, label: &str) -> Option<usize> {
    row.iter()
        .enumerate()
        .skip(start)
        .find_map(|(idx, cell)| (cell_as_string(cell).trim() == label).then_some(idx))
}

fn normalize_external_aux_code(raw_code: &str) -> String {
    let code = raw_code.trim();
    if let Ok(number) = code.parse::<i64>() {
        format!("{:05}", number)
    } else {
        code.to_string()
    }
}

/// 生成用于外账结存唯一键的规格值。
///
/// 这里仅消除导出格式差异，不做“包含即相等”的模糊比较：统一全半角
/// 分隔符、中文计量单位、包装后缀，以及片/粒/支/t 等数量单位后再比较。
/// 瓶、袋等容器仍然保留，避免把不同包装的同规格药品混为一项。
fn normalize_inventory_spec_key(value: &str) -> String {
    let mut spec = normalize_text(value)
        .to_lowercase()
        .replace('：', ":")
        .replace('；', ":")
        .replace('／', "/")
        .replace('－', "-")
        .replace("毫克", "mg")
        .replace("毫升", "ml")
        .replace("微克", "ug")
        .replace("公斤", "kg")
        .replace("克", "g");

    for suffix in ["盒", "瓶", "袋", "箱", "包"] {
        let suffix = format!("/{}", suffix);
        if spec.ends_with(&suffix) {
            spec.truncate(spec.len() - suffix.len());
            break;
        }
    }

    for unit in ["片", "粒", "支", "t"] {
        spec = spec.replace(unit, "");
    }

    spec.trim_end_matches('/').to_string()
}

/// 将库管表和辅助项目中的剂型归并为可比较的剂型族。
///
/// 数量式明细账没有单独的“剂型”列，因此从辅助项目药名中提取常见剂型；
/// 无法从药名推导时返回空串，避免人为猜一个剂型造成错误匹配。
fn inventory_dosage_family(value: &str) -> String {
    let normalized = normalize_text(value);
    let candidates = [
        ("胶囊", "胶囊"),
        ("注射", "注射"),
        ("口服液", "溶液"),
        ("溶液", "溶液"),
        ("开塞露", "溶液"),
        ("糖浆", "糖浆"),
        ("颗粒", "颗粒"),
        ("滴眼", "滴眼"),
        ("滴耳", "滴耳"),
        ("软膏", "软膏"),
        ("乳膏", "软膏"),
        ("喷雾", "喷雾"),
        ("凝胶", "凝胶"),
        ("栓", "栓"),
        ("丸", "丸"),
        ("散", "散"),
        ("贴", "贴"),
        ("片", "片"),
    ];

    candidates
        .iter()
        .find_map(|(keyword, family)| {
            normalized
                .contains(keyword)
                .then_some((*family).to_string())
        })
        .unwrap_or_default()
}

fn external_inventory_tags(value: &str) -> Vec<String> {
    let normalized = normalize_text(value);
    let Ok(re) = Regex::new(r"\(([^()]*)\)|\[([^\[\]]*)\]") else {
        return Vec::new();
    };

    re.captures_iter(&normalized)
        .filter_map(|caps| caps.get(1).or_else(|| caps.get(2)))
        .flat_map(|m| m.as_str().split(['/', '、', ',', '，']))
        .map(normalize_text)
        .filter(|tag| !tag.is_empty())
        .collect()
}

/// 判断带厂家标签的外账辅助项目是否对应库管厂家。
///
/// 外账辅助项目中的厂家通常写在药名括号里（如“桑寄生（蕴德）”，而库管
/// 表把厂家放在独立列中）。这个判断单独抽出来，供匹配排序使用：当同名同
/// 规格同时存在一个通用项目和一个带厂家项目时，应优先命中厂家项目。
fn external_inventory_factory_matches_warehouse(
    item: &ExternalInventoryItem,
    warehouse: &WarehouseItem,
    factory_map: &HashMap<String, String>,
) -> bool {
    let tags = external_inventory_tags(&item.name);
    if tags.is_empty() {
        return false;
    }

    let warehouse_name = normalize_text(&warehouse.name);
    let warehouse_factory = normalize_text(&warehouse.factory);
    let factory_alias = find_factory_abbreviation(&warehouse_factory, factory_map);

    tags.iter().any(|tag| {
        warehouse_name.contains(tag)
            || warehouse_factory.contains(tag)
            || factory_alias
                .as_ref()
                .is_some_and(|alias| tag.contains(alias) || alias.contains(tag))
    })
}

fn warehouse_inventory_key(warehouse: &WarehouseItem) -> (String, String, String) {
    (
        clean_drug_name(&warehouse.name),
        normalize_inventory_spec_key(&warehouse.spec),
        inventory_dosage_family(&warehouse.dosage_form),
    )
}

/// 按“辅助项目 = 药品名称 + 规格 + 剂型 + 制药厂”匹配外账结存。
///
/// 数量式明细账的辅助项目没有独立的剂型、厂家列，所以剂型从辅助项目药名
/// 推导，厂家用辅助项目括号标签与库管药名/厂家核验；辅助项目没有厂家标签
/// 时，只有在所选三张库管表中该名称、规格、剂型只对应一个厂家才允许命中；
/// 同名同规格同时存在通用项目和厂家项目时，优先使用与库管厂家匹配的厂家项目。
/// 规格必须使用规范化后的完整值相等，禁止旧的 contains/同名兜底。
fn match_external_inventory_key(
    warehouse: &WarehouseItem,
    items: &[ExternalInventoryItem],
    factory_map: &HashMap<String, String>,
    warehouse_factory_counts: &HashMap<(String, String, String), HashSet<String>>,
) -> Option<(String, String)> {
    let (warehouse_name, warehouse_spec, warehouse_dosage) = warehouse_inventory_key(warehouse);

    let candidates: Vec<&ExternalInventoryItem> = items
        .iter()
        .filter(|item| item.clean_name == warehouse_name)
        .filter(|item| item.spec_key == warehouse_spec)
        .filter(|item| {
            let item_dosage = inventory_dosage_family(&item.name);
            warehouse_dosage.is_empty() || item_dosage.is_empty() || item_dosage == warehouse_dosage
        })
        .collect();

    // 与内账比对的厂家匹配规则保持一致：先用库管厂家锁定带厂家后缀的
    // 外账项目。不能直接把“无厂家通用项目”和“厂家项目”一起视为歧义，
    // 否则“桑寄生”会遮蔽“桑寄生（蕴德）”。
    if let Some(item) = unique_inventory_match(candidates.iter().copied().filter(|item| {
        !external_inventory_tags(&item.name).is_empty()
            && external_inventory_factory_matches_warehouse(item, warehouse, factory_map)
    })) {
        return Some((item.code.clone(), item.name.clone()));
    }

    // 如果存在厂家专属项目但没有一个能和库管厂家对应，不能回退到通用项目，
    // 避免把其他厂家的结存错挂到当前厂家。
    if candidates
        .iter()
        .any(|item| !external_inventory_tags(&item.name).is_empty())
    {
        return None;
    }

    // 没有厂家标签时才允许按唯一同名同规格项目匹配；多个厂家共用一个
    // 通用外账项目时仍交给人工核对。
    let item = unique_inventory_match(candidates.iter().copied())?;
    if warehouse_factory_counts
        .get(&(warehouse_name, warehouse_spec, warehouse_dosage))
        .is_some_and(|factories| factories.len() > 1)
    {
        return None;
    }

    Some((item.code.clone(), item.name.clone()))
}

fn split_external_auxiliary(value: &str) -> Option<(String, String, String)> {
    let mut parts = value.trim().splitn(2, |c: char| c.is_whitespace());
    let raw_code = parts.next()?.trim();
    let description = parts.next()?.trim();
    if raw_code.is_empty() || description.is_empty() {
        return None;
    }

    let mut name_parts = description.splitn(2, |c: char| c.is_whitespace());
    let name = name_parts.next()?.trim().to_string();
    let spec = name_parts.next().unwrap_or_default().trim().to_string();
    if name.is_empty() {
        return None;
    }

    Some((normalize_external_aux_code(raw_code), name, spec))
}

/// 读取外账数量式明细账中的 1405 结存数据。
///
/// 该账表不依赖迁账模板的【辅助信息】Sheet，直接使用“辅助项目”里的五位
/// 辅助编码、名称/规格，以及“余额”分组下的数量、单价和金额。
fn load_external_quantity_ledger(ledger_path: &Path) -> Result<ExternalTemplateData, String> {
    let path_display = ledger_path.display().to_string();
    let (_, rows) = read_sheet_rows(ledger_path, &[], "外账数量式明细账")
        .map_err(|e| format!("读取外账数量式明细账 {} 失败: {}", path_display, e))?;

    let header_idx = rows
        .iter()
        .take(12)
        .position(|row| {
            find_col_idx(row, &["科目"]).is_some()
                && find_col_idx(row, &["辅助项目"]).is_some()
                && find_col_idx(row, &["摘要"]).is_some()
        })
        .ok_or_else(|| format!("外账数量式明细账 {} 中未识别到表头", path_display))?;
    let header = &rows[header_idx];
    let subheader = rows
        .get(header_idx + 1)
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少数量子表头", path_display))?;

    let col_subject = find_col_idx(header, &["科目"])
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【科目】列", path_display))?;
    let col_auxiliary = find_col_idx(header, &["辅助项目"])
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【辅助项目】列", path_display))?;
    let col_summary = find_col_idx(header, &["摘要"])
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【摘要】列", path_display))?;
    let balance_start = find_external_group_column(header, "余额")
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【余额】分组", path_display))?;
    let col_qty = find_external_subheader_column(subheader, balance_start, "数量")
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【余额-数量】列", path_display))?;
    let col_price = find_external_subheader_column(subheader, balance_start, "单价")
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【余额-单价】列", path_display))?;
    let col_amount = find_external_subheader_column(subheader, balance_start, "金额")
        .ok_or_else(|| format!("外账数量式明细账 {} 缺少【余额-金额】列", path_display))?;

    let mut data = ExternalTemplateData::default();
    let mut row_priority_by_code = HashMap::new();
    for row in rows.iter().skip(header_idx + 2) {
        let subject = row.get(col_subject).map(cell_as_string).unwrap_or_default();
        if !subject.trim_start().starts_with("1405") {
            continue;
        }

        let auxiliary = row
            .get(col_auxiliary)
            .map(cell_as_string)
            .unwrap_or_default();
        let Some((code, name, spec)) = split_external_auxiliary(&auxiliary) else {
            continue;
        };

        let summary = row.get(col_summary).map(cell_as_string).unwrap_or_default();
        let priority = if summary.contains("本年累计") {
            3
        } else if summary.contains("本月合计") {
            2
        } else if summary.contains("期初余额") {
            1
        } else {
            0
        };
        if row_priority_by_code
            .get(&code)
            .is_some_and(|current| *current > priority)
        {
            continue;
        }
        row_priority_by_code.insert(code.clone(), priority);

        if !data.inventory_by_code.contains_key(&code) {
            let item = ExternalInventoryItem {
                code: code.clone(),
                name: name.clone(),
                spec: spec.clone(),
                spec_key: normalize_inventory_spec_key(&spec),
                norm_spec: normalize_text(&spec),
                norm_name: normalize_text(&name),
                clean_name: clean_drug_name(&name),
            };
            data.inventory_items.push(item.clone());
            data.inventory_by_code.insert(code.clone(), item);
        }

        let price = row.get(col_price).map(cell_as_f64).unwrap_or(0.0);
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let amount = row.get(col_amount).map(cell_as_f64).unwrap_or(0.0);
        data.balance_by_code.insert(code, (price, qty, amount));
    }

    if data.inventory_items.is_empty() {
        return Err(format!(
            "外账数量式明细账 {} 中未读取到 1405 存货结存数据",
            path_display
        ));
    }

    Ok(data)
}

fn get_default_factory_abbr_map() -> HashMap<&'static str, &'static str> {
    let mut m = HashMap::new();
    m.insert("浙江华海药业股份有限公司", "华海");
    m.insert("江苏恩华药业股份有限公司", "恩华");
    m.insert("成都康弘药业集团股份有限公司", "康弘");
    m.insert("上海信谊金朱药业有限公司", "信谊");
    m.insert("瑞阳制药股份有限公司", "瑞阳");
    m.insert("石药集团欧意药业有限公司", "欧意");
    m.insert("沈阳华泰药物研究有限公司", "华泰");
    m.insert("湖南洞庭药业股份有限公司", "洞庭");
    m.insert("齐鲁制药有限公司", "齐鲁");
    m.insert("河北龙海药业有限公司", "龙海");
    m.insert("西南药业股份有限公司", "西南");
    m.insert("华阴市华山制药有限责任公司", "华山");
    m.insert("北京益民药业有限公司", "益民");
    m.insert("江苏联环药业股份有限公司", "联环");
    m.insert("吉林省博大制药有限责任公司", "博大");
    m.insert("海南灵康制药有限公司", "灵康");
    m.insert("山东华鲁制药有限公司", "华鲁");
    m.insert("常州制药厂有限公司", "常州");
    m.insert("鲁南贝特制药有限公司", "贝特");
    m.insert("黑龙江迪龙制药有限公司", "迪龙");
    m.insert("重庆圣华曦药业股份有限公司", "圣华曦");
    m.insert("北京诺华制药有限公司", "诺华");
    m.insert("辉瑞制药有限公司", "辉瑞");
    m.insert("中美天津史克制药有限公司", "史克");
    m.insert("阿斯利康制药有限公司", "阿斯利康");
    m.insert("西安杨森制药有限公司", "杨森");
    m.insert("拜耳医药保健有限公司", "拜耳");
    m.insert("默沙东制药有限公司", "默沙东");
    m.insert("赛诺菲制药有限公司", "赛诺菲");
    m.insert("施贵宝制药有限公司", "施贵宝");
    m.insert("吉林省西点药业科技发展股份有限公司", "西点");
    m.insert("山东新时代药业有限公司", "新时代");
    m.insert("石药集团中诺药业(石家庄)有限公司", "中诺");
    m.insert("正大天晴药业集团股份有限公司", "正大天晴");
    // 这些简称曾经是内置规则的一部分，不能因为增加外部配置而丢失。
    m.insert("扬子江药业集团有限公司", "扬子江");
    m.insert("通化东宝药业股份有限公司", "通化东宝");
    m.insert("远大医药(中国)有限公司", "远大");
    m.insert("长春高新技术产业(集团)股份有限公司", "金赛");
    m
}

fn get_merged_factory_abbr_map(config: Option<&ConfigData>) -> HashMap<String, String> {
    let mut m = HashMap::new();
    for (k, v) in get_default_factory_abbr_map() {
        m.insert(normalize_text(k), normalize_text(v));
    }
    if let Some(cfg) = config {
        for (k, v) in &cfg.factory_abbreviations {
            let key = normalize_text(k);
            let value = normalize_text(v);
            if !key.is_empty() && !value.is_empty() {
                m.insert(key, value);
            }
        }
    }
    m
}

fn find_factory_abbreviation(
    normalized_factory: &str,
    factory_map: &HashMap<String, String>,
) -> Option<String> {
    let mut aliases: Vec<(&String, &String)> = factory_map.iter().collect();
    // 优先使用更长、更具体的厂家名称，避免 HashMap 遍历顺序影响结果。
    aliases.sort_by(|(full_a, abbr_a), (full_b, abbr_b)| {
        full_b
            .len()
            .cmp(&full_a.len())
            .then_with(|| full_a.cmp(full_b))
            .then_with(|| abbr_a.cmp(abbr_b))
    });

    aliases
        .into_iter()
        .find(|(full, abbr)| {
            normalized_factory.contains(full.as_str()) || normalized_factory.contains(abbr.as_str())
        })
        .map(|(_, abbr)| abbr.clone())
}

fn unique_inventory_match<'a>(
    candidates: impl IntoIterator<Item = &'a ExternalInventoryItem>,
) -> Option<&'a ExternalInventoryItem> {
    let mut candidates: Vec<&ExternalInventoryItem> = candidates.into_iter().collect();
    candidates.sort_by(|a, b| {
        a.code
            .cmp(&b.code)
            .then_with(|| a.norm_name.cmp(&b.norm_name))
            .then_with(|| a.norm_spec.cmp(&b.norm_spec))
    });
    candidates.dedup_by(|a, b| a.code == b.code);
    (candidates.len() == 1).then(|| candidates[0])
}

fn inventory_spec_matches(source_spec: &str, item_spec: &str) -> bool {
    source_spec.is_empty()
        || item_spec.is_empty()
        || item_spec.contains(source_spec)
        || source_spec.contains(item_spec)
}

#[derive(Debug, Clone, Copy, Default)]
struct InboundColumns {
    header_idx: usize,
    name: Option<usize>,
    spec: Option<usize>,
    factory: Option<usize>,
    qty: Option<usize>,
    price: Option<usize>,
    amount: Option<usize>,
    supplier: Option<usize>,
}

fn detect_inbound_columns(rows: &[Vec<Data>]) -> Option<InboundColumns> {
    for (r_idx, row) in rows.iter().enumerate().take(20) {
        let mut columns = InboundColumns {
            header_idx: r_idx,
            ..InboundColumns::default()
        };

        for (c_idx, cell) in row.iter().enumerate() {
            let text = cell_as_string(cell);
            if text.contains("通用名") || text.contains("药品名称") || text.contains("品名")
            {
                columns.name = Some(c_idx);
            } else if text.contains("规格") {
                columns.spec = Some(c_idx);
            } else if text.contains("厂家") || text.contains("生产企业") {
                columns.factory = Some(c_idx);
            } else if text.contains("入库数量") || text == "数量" || text.contains("实收数量")
            {
                columns.qty = Some(c_idx);
            } else if text.contains("进价单价") || text.contains("购进单价") || text == "进价"
            {
                columns.price = Some(c_idx);
            } else if text.contains("进价金额")
                || text.contains("购进金额")
                || text.contains("金额(进)")
            {
                columns.amount = Some(c_idx);
            } else if text.contains("供货单位")
                || text.contains("供应商")
                || text.contains("供货商")
            {
                columns.supplier = Some(c_idx);
            }
        }

        if columns.name.is_some()
            && columns.qty.is_some()
            && (columns.amount.is_some() || columns.price.is_some())
        {
            return Some(columns);
        }
    }

    None
}

fn match_external_inventory_code(
    name: &str,
    spec: &str,
    factory: &str,
    items: &[ExternalInventoryItem],
    factory_map: &HashMap<String, String>,
) -> Option<(String, String)> {
    let norm_name = normalize_text(name);
    let clean_name = clean_drug_name(name);
    let norm_spec = normalize_text(spec);
    let norm_factory = normalize_text(factory);

    let abbr = find_factory_abbreviation(&norm_factory, factory_map);
    let factory_matches = |it: &&ExternalInventoryItem| {
        abbr.as_ref()
            .map(|value| it.norm_name.contains(value))
            .unwrap_or(false)
    };
    let spec_matches =
        |it: &&ExternalInventoryItem| inventory_spec_matches(&norm_spec, &it.norm_spec);

    // 1. 厂家简称 + 通用名 + 规格，是外账辅助档案最可靠的组合。
    if let Some(it) = unique_inventory_match(
        items
            .iter()
            .filter(|it| it.clean_name == clean_name)
            .filter(factory_matches)
            .filter(spec_matches),
    ) {
        return Some((it.code.clone(), it.name.clone()));
    }

    // 2. 全称相同也必须先通过规格约束，避免同名不同规格时直接命中错误编码。
    if let Some(it) = unique_inventory_match(
        items
            .iter()
            .filter(|it| it.norm_name == norm_name)
            .filter(spec_matches),
    ) {
        return Some((it.code.clone(), it.name.clone()));
    }

    // 3. 通用名 + 规格。若仍有多个厂家候选，不再猜测，交给人工处理。
    if let Some(it) = unique_inventory_match(
        items
            .iter()
            .filter(|it| it.clean_name == clean_name)
            .filter(spec_matches),
    ) {
        return Some((it.code.clone(), it.name.clone()));
    }

    // 4. 没有规格或档案规格为空时，厂家简称 + 通用名仍可作为明确匹配。
    if let Some(it) = unique_inventory_match(
        items
            .iter()
            .filter(|it| it.clean_name == clean_name)
            .filter(factory_matches),
    ) {
        return Some((it.code.clone(), it.name.clone()));
    }

    // 5. 只有一个同名存货时才允许放宽到通用名匹配。
    if let Some(it) = unique_inventory_match(items.iter().filter(|it| it.clean_name == clean_name))
    {
        return Some((it.code.clone(), it.name.clone()));
    }

    // 6. 模糊匹配同样必须唯一，并且要满足规格约束；不再返回第一个候选。
    if clean_name.chars().count() >= 3 {
        if let Some(it) = unique_inventory_match(
            items
                .iter()
                .filter(|it| {
                    it.norm_name.contains(&clean_name) || clean_name.contains(&it.norm_name)
                })
                .filter(spec_matches),
        ) {
            return Some((it.code.clone(), it.name.clone()));
        }
    }

    None
}

fn match_external_supplier_code(
    supplier: &str,
    supplier_map: &HashMap<String, (String, String)>,
) -> Option<(String, String)> {
    let norm = normalize_text(supplier);
    if let Some(res) = supplier_map.get(&norm) {
        return Some(res.clone());
    }

    let mut candidates: Vec<(&String, &(String, String))> = supplier_map
        .iter()
        .filter(|(k, _)| norm.contains(k.as_str()) || k.as_str().contains(&norm))
        .collect();
    candidates.sort_by(|(a, _), (b, _)| b.len().cmp(&a.len()).then_with(|| a.cmp(b)));

    let Some((best_key, best_value)) = candidates.first() else {
        return None;
    };
    let best_len = best_key.len();
    let best_code = best_value.0.clone();
    if candidates
        .iter()
        .filter(|(key, _)| key.len() == best_len)
        .all(|(_, value)| value.0 == best_code)
    {
        Some((*candidates[0].1).clone())
    } else {
        None
    }
}

fn external_candidate_option(
    item: &ExternalInventoryItem,
    template_data: &ExternalTemplateData,
) -> LedgerCandidateOption {
    let (price, qty, amount) = template_data
        .balance_by_code
        .get(&item.code)
        .cloned()
        .unwrap_or((0.0, 0.0, 0.0));

    LedgerCandidateOption {
        code: item.code.clone(),
        name: item.name.clone(),
        spec: item.spec.clone(),
        qty,
        price,
        amount,
    }
}

fn external_candidate_score(query: &str, item: &ExternalInventoryItem) -> i32 {
    let normalized_query = normalize_text(query);
    if normalized_query.is_empty() {
        return 0;
    }

    let clean_query = clean_drug_name(&normalized_query);
    let mut score = 0;

    if item.code == normalized_query {
        score += 200;
    } else if item.code.contains(&normalized_query) {
        score += 120;
    }

    if item.norm_name == normalized_query || item.clean_name == clean_query {
        score += 160;
    } else if item.norm_name.contains(&normalized_query)
        || normalized_query.contains(&item.norm_name)
        || item.clean_name.contains(&clean_query)
        || clean_query.contains(&item.clean_name)
    {
        score += 100;
    }

    if !item.norm_spec.is_empty()
        && (item.norm_spec.contains(&normalized_query)
            || normalized_query.contains(&item.norm_spec))
    {
        score += 60;
    }

    score
}

/// 在当前外账模板的【辅助信息】存货字典中搜索人工匹配候选。
///
/// 外账使用的是模板内的五位辅助编码，不依赖内账数量金额总账，因此单独
/// 提供搜索入口，且候选编码全部来自当前模板，避免把内账编码误带入外账。
pub fn search_external_inventory_candidates(
    template_path: &Path,
    query: &str,
    limit: usize,
) -> Result<Vec<LedgerCandidateOption>, String> {
    let template_data = load_external_template(template_path)?;
    let mut scored: Vec<(i32, &ExternalInventoryItem)> = template_data
        .inventory_items
        .iter()
        .filter_map(|item| {
            let score = external_candidate_score(query, item);
            (score > 0).then_some((score, item))
        })
        .collect();

    scored.sort_by(|(score_a, item_a), (score_b, item_b)| {
        score_b
            .cmp(score_a)
            .then_with(|| item_a.code.cmp(&item_b.code))
            .then_with(|| item_a.name.cmp(&item_b.name))
    });

    Ok(scored
        .into_iter()
        .take(limit.max(1).min(200))
        .map(|(_, item)| external_candidate_option(item, &template_data))
        .collect())
}

fn build_external_inventory_candidates(
    name: &str,
    spec: &str,
    factory: &str,
    template_data: &ExternalTemplateData,
) -> Vec<LedgerCandidateOption> {
    let factory_query = normalize_text(factory);
    let mut scored: Vec<(i32, &ExternalInventoryItem)> = template_data
        .inventory_items
        .iter()
        .filter_map(|item| {
            let mut score = external_candidate_score(name, item);
            if !spec.trim().is_empty()
                && inventory_spec_matches(&normalize_text(spec), &item.norm_spec)
            {
                score += 50;
            }
            if !factory_query.is_empty() && item.norm_name.contains(&factory_query) {
                score += 20;
            }
            (score > 0).then_some((score, item))
        })
        .collect();

    scored.sort_by(|(score_a, item_a), (score_b, item_b)| {
        score_b
            .cmp(score_a)
            .then_with(|| item_a.code.cmp(&item_b.code))
            .then_with(|| item_a.name.cmp(&item_b.name))
    });

    scored
        .into_iter()
        .take(25)
        .map(|(_, item)| external_candidate_option(item, template_data))
        .collect()
}

fn find_confirmed_external_inventory_match(
    row_index: usize,
    confirmed_mappings: Option<&[ConfirmedExternalInventoryMapping]>,
    template_data: &ExternalTemplateData,
) -> Option<(String, String)> {
    let mapping = confirmed_mappings?
        .iter()
        .find(|mapping| mapping.row_index == row_index)?;
    let code = mapping.aux_code.trim();
    if code.is_empty() {
        return None;
    }

    template_data
        .inventory_by_code
        .get(code)
        .map(|item| (item.code.clone(), item.name.clone()))
}

/// 外账输出路径与内账保持一致：未指定目录的文件名保存到对应输入文件同级目录；
/// 用户填写绝对路径时保留用户指定位置。
fn resolve_external_output_path(
    source_path: &Path,
    output_path: Option<&Path>,
    default_name: &str,
) -> PathBuf {
    let parent = source_path
        .parent()
        .filter(|path| !path.as_os_str().is_empty())
        .unwrap_or_else(|| Path::new("."));

    match output_path.filter(|path| !path.as_os_str().is_empty()) {
        Some(path) if path.is_absolute() => path.to_path_buf(),
        Some(path) => parent.join(path),
        None => parent.join(default_name),
    }
}

// ----------------------------------------------------
// 3. 外账入库凭证生成 (纯 Rust 原生实现)
// ----------------------------------------------------

pub fn generate_external_inbound_voucher(
    inbound_path: &Path,
    template_path: &Path,
    output_path: Option<&Path>,
    custom_date: Option<&str>,
    voucher_no: Option<&str>,
    _config: Option<&ConfigData>,
    confirmed_mappings: Option<&[ConfirmedExternalInventoryMapping]>,
) -> Result<ExternalInboundResult, String> {
    let template_data = load_external_template(template_path)?;
    let factory_abbr = get_merged_factory_abbr_map(_config);

    let mut in_excel = open_excel(inbound_path)?;
    let sheet_names = in_excel.sheet_names();
    if sheet_names.is_empty() {
        return Err("入库单没有工作表".to_string());
    }

    // 入库单可能包含封面、说明页或空白页，不能假设第一个 Sheet 就是数据页。
    let mut selected_sheet = None;
    for sheet_name in sheet_names {
        let Ok(range) = in_excel.worksheet_range(&sheet_name) else {
            continue;
        };
        let rows: Vec<Vec<Data>> = range.rows().map(|r| r.to_vec()).collect();
        if let Some(columns) = detect_inbound_columns(&rows) {
            selected_sheet = Some((sheet_name, rows, columns));
            break;
        }
    }

    let (in_sheet, rows, columns) = selected_sheet
        .ok_or_else(|| "入库单所有工作表中均未识别到药品名称、数量及金额/进价表头".to_string())?;
    let _ = in_sheet;
    let header_idx = columns.header_idx;
    let col_name = columns.name.expect("入库表头已校验药品名称列");
    let col_qty = columns.qty.expect("入库表头已校验数量列");
    let col_spec = columns.spec;
    let col_factory = columns.factory;
    let col_price = columns.price;
    let col_amt = columns.amount;
    let col_supplier = columns.supplier;

    struct RawInboundRow {
        row_no: usize,
        name: String,
        spec: String,
        factory: String,
        qty: f64,
        amount: f64,
        #[allow(dead_code)]
        supplier: String,
    }

    let mut total_raw_count: usize = 0;
    let mut supplier_names_order = Vec::new();
    let mut supplier_grouped_rows: HashMap<String, Vec<RawInboundRow>> = HashMap::new();

    for (r_idx, row) in rows.iter().enumerate().skip(header_idx + 1) {
        let name = cell_as_string(row.get(col_name).unwrap_or(&Data::Empty))
            .trim()
            .to_string();
        if name.is_empty()
            || name.contains("合计")
            || name.contains("总计")
            || name.contains("制表")
            || name.starts_with("报表")
        {
            continue;
        }

        let spec = col_spec
            .and_then(|c| row.get(c))
            .map(cell_as_string)
            .unwrap_or_default()
            .trim()
            .to_string();
        let factory = col_factory
            .and_then(|c| row.get(c))
            .map(cell_as_string)
            .unwrap_or_default()
            .trim()
            .to_string();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let amount = if let Some(c) = col_amt {
            let direct_amt = row.get(c).map(cell_as_f64).unwrap_or(0.0);
            if direct_amt > 0.0 {
                (direct_amt * 100.0).round() / 100.0
            } else {
                let price = col_price
                    .and_then(|cp| row.get(cp))
                    .map(cell_as_f64)
                    .unwrap_or(0.0);
                (qty * price * 100.0).round() / 100.0
            }
        } else {
            let price = col_price
                .and_then(|cp| row.get(cp))
                .map(cell_as_f64)
                .unwrap_or(0.0);
            (qty * price * 100.0).round() / 100.0
        };

        let supplier = col_supplier
            .and_then(|c| row.get(c))
            .map(cell_as_string)
            .unwrap_or_default()
            .trim()
            .to_string();
        let s_key = if supplier.is_empty() {
            "未指定供应商".to_string()
        } else {
            supplier.clone()
        };

        if !supplier_grouped_rows.contains_key(&s_key) {
            supplier_names_order.push(s_key.clone());
        }

        let item = RawInboundRow {
            row_no: r_idx + 1,
            name,
            spec,
            factory,
            qty,
            amount,
            supplier: s_key.clone(),
        };

        supplier_grouped_rows.entry(s_key).or_default().push(item);
        total_raw_count += 1;
    }

    let voucher_date = detect_month_end_date(custom_date, inbound_path.to_str());
    let v_no = voucher_no.unwrap_or("1").to_string();

    let mut voucher_rows = Vec::new();
    let mut unmatched_drugs = Vec::new();
    let mut unmatched_suppliers = Vec::new();
    let mut total_debit_cents: i64 = 0;
    let mut total_credit_cents: i64 = 0;
    let mut row_counter = 1;

    for s_name in &supplier_names_order {
        let items = &supplier_grouped_rows[s_name];
        let mut supplier_sum_cents: i64 = 0;

        for item in items {
            let cents = (item.amount * 100.0).round() as i64;
            supplier_sum_cents += cents;
            total_debit_cents += cents;

            let matched = find_confirmed_external_inventory_match(
                item.row_no,
                confirmed_mappings,
                &template_data,
            )
            .or_else(|| {
                match_external_inventory_code(
                    &item.name,
                    &item.spec,
                    &item.factory,
                    &template_data.inventory_items,
                    &factory_abbr,
                )
            });
            let (aux_code, aux_name, is_unmatched) = match matched {
                Some((code, name)) => (code, name, false),
                None => {
                    unmatched_drugs.push(ExternalUnmatchedDrug {
                        id: item.row_no,
                        row_index: item.row_no,
                        name: item.name.clone(),
                        factory: item.factory.clone(),
                        spec: item.spec.clone(),
                        qty: item.qty,
                        amount: item.amount,
                        supplier: Some(s_name.clone()),
                        candidates: build_external_inventory_candidates(
                            &item.name,
                            &item.spec,
                            &item.factory,
                            &template_data,
                        ),
                    });
                    (String::new(), item.name.clone(), true)
                }
            };

            let brief_supplier = s_name
                .replace("有限公司", "")
                .replace("股份有限公司", "")
                .replace("有限责任公司", "")
                .replace("医药", "")
                .replace("石家庄", "");

            voucher_rows.push(ExternalVoucherRow {
                row_no: row_counter,
                date: voucher_date.clone(),
                voucher_type: "记".to_string(),
                voucher_no: v_no.clone(),
                summary: format!("{}到货", brief_supplier),
                subject_code: "1405".to_string(),
                subject_name: "库存商品".to_string(),
                debit_amount: Some(item.amount),
                credit_amount: None,
                exchange_rate: Some(1.0),
                currency: "人民币".to_string(),
                qty: Some(item.qty),
                aux_code,
                aux_name,
                settle_method: String::new(),
                bill_no: String::new(),
                occur_date: voucher_date.clone(),
                is_credit: false,
                is_unmatched,
            });
            row_counter += 1;
        }

        let credit_amount = (supplier_sum_cents as f64) / 100.0;
        total_credit_cents += supplier_sum_cents;

        let (sup_code, sup_name, supplier_is_unmatched) =
            match match_external_supplier_code(s_name, &template_data.supplier_map) {
                Some((code, name)) => (code, name, false),
                None => {
                    unmatched_suppliers.push(ExternalUnmatchedSupplier {
                        supplier: s_name.clone(),
                        item_count: items.len(),
                        amount: credit_amount,
                    });
                    (String::new(), s_name.clone(), true)
                }
            };

        let brief_supplier = s_name
            .replace("有限公司", "")
            .replace("股份有限公司", "")
            .replace("有限责任公司", "")
            .replace("医药", "")
            .replace("石家庄", "");

        voucher_rows.push(ExternalVoucherRow {
            row_no: row_counter,
            date: voucher_date.clone(),
            voucher_type: "记".to_string(),
            voucher_no: v_no.clone(),
            summary: format!("{}到货", brief_supplier),
            subject_code: "2202".to_string(),
            subject_name: "应付账款".to_string(),
            debit_amount: None,
            credit_amount: Some(credit_amount),
            exchange_rate: Some(1.0),
            currency: "人民币".to_string(),
            qty: None,
            aux_code: sup_code,
            aux_name: sup_name,
            settle_method: String::new(),
            bill_no: String::new(),
            occur_date: voucher_date.clone(),
            is_credit: true,
            is_unmatched: supplier_is_unmatched,
        });
        row_counter += 1;
    }

    let total_debit = (total_debit_cents as f64) / 100.0;
    let total_credit = (total_credit_cents as f64) / 100.0;
    let diff = ((total_debit_cents - total_credit_cents) as f64) / 100.0;
    let is_balanced = total_debit_cents == total_credit_cents;

    let out_file_path = resolve_external_output_path(
        inbound_path,
        output_path,
        "表格迁账参考模板_西药入库_已生成.xlsx",
    );

    write_external_voucher_workbook(template_path, &out_file_path, &voucher_rows)?;

    let total_items = total_raw_count;
    let unmatched_count = unmatched_drugs.len();
    let matched_count = total_items.saturating_sub(unmatched_count);
    let match_rate = if total_items > 0 {
        ((matched_count as f64) / (total_items as f64) * 100.0 * 10.0).round() / 10.0
    } else {
        0.0
    };

    Ok(ExternalInboundResult {
        success: true,
        output_file: out_file_path.to_string_lossy().to_string(),
        voucher_date,
        voucher_no: v_no,
        total_items,
        supplier_count: supplier_names_order.len(),
        total_entries: voucher_rows.len(),
        matched_count,
        unmatched_count,
        unmatched_supplier_count: unmatched_suppliers.len(),
        match_rate,
        total_debit,
        total_credit,
        diff,
        is_balanced,
        unmatched_drugs,
        unmatched_suppliers,
        voucher_rows,
    })
}

// ----------------------------------------------------
// 4. 外账出库凭证生成 (纯 Rust 原生实现)
// ----------------------------------------------------

pub fn generate_external_outbound_voucher(
    sales_path: &Path,
    template_path: &Path,
    output_path: Option<&Path>,
    custom_date: Option<&str>,
    voucher_no: Option<&str>,
    _config: Option<&ConfigData>,
    confirmed_mappings: Option<&[ConfirmedExternalInventoryMapping]>,
) -> Result<ExternalOutboundResult, String> {
    let template_data = load_external_template(template_path)?;
    let factory_abbr = get_merged_factory_abbr_map(_config);

    let mut sales_excel = open_excel(sales_path)?;
    let sheet_names = sales_excel.sheet_names();
    let sheet_name = sheet_names
        .iter()
        .find(|s| s.contains("销售明细"))
        .cloned()
        .or_else(|| {
            if sheet_names.len() > 1 {
                Some(sheet_names[1].clone())
            } else {
                sheet_names.first().cloned()
            }
        })
        .ok_or_else(|| "销售表没有工作表".to_string())?;

    let range = sales_excel
        .worksheet_range(&sheet_name)
        .map_err(|e| format!("读取销售表失败: {}", e))?;

    let rows: Vec<Vec<Data>> = range.rows().map(|r| r.to_vec()).collect();
    if rows.is_empty() {
        return Err("销售表内容为空".to_string());
    }

    let mut header_idx = 0;
    let mut col_name = None;
    let mut col_spec = None;
    let mut col_factory = None;
    let mut col_qty = None;
    let mut col_price = None;
    let mut col_amt = None;

    for (r_idx, row) in rows.iter().enumerate().take(10) {
        for (c_idx, cell) in row.iter().enumerate() {
            let s = cell_as_string(cell);
            if s.contains("通用名") || s.contains("药品名称") || s.contains("品名") {
                col_name = Some(c_idx);
            } else if s.contains("规格") {
                col_spec = Some(c_idx);
            } else if s.contains("厂家") || s.contains("生产企业") || s.contains("制药厂")
            {
                col_factory = Some(c_idx);
            } else if s.contains("实发数量") || s.contains("销售数量") || s == "数量" {
                col_qty = Some(c_idx);
            } else if s.contains("成本进价") || s.contains("进价单价") || s == "进价" {
                col_price = Some(c_idx);
            } else if s.contains("进价金额") || s.contains("成本金额") {
                col_amt = Some(c_idx);
            }
        }
        if col_name.is_some() && col_qty.is_some() {
            header_idx = r_idx;
            break;
        }
    }

    let col_name = col_name.ok_or_else(|| "销售表中未识别到【药品名称】列".to_string())?;
    let col_qty = col_qty.ok_or_else(|| "销售表中未识别到【数量】列".to_string())?;

    struct RawSaleItem {
        row_no: usize,
        name: String,
        spec: String,
        factory: String,
        qty: f64,
        #[allow(dead_code)]
        price: f64,
        amount: f64,
    }

    let mut raw_items: Vec<RawSaleItem> = Vec::new();
    let mut group_map: HashMap<(String, String, String), usize> = HashMap::new();

    for (r_idx, row) in rows.iter().enumerate().skip(header_idx + 1) {
        let name = cell_as_string(row.get(col_name).unwrap_or(&Data::Empty))
            .trim()
            .to_string();
        if name.is_empty()
            || name.contains("合计")
            || name.contains("总计")
            || name.contains("制表")
        {
            continue;
        }

        let spec = col_spec
            .and_then(|c| row.get(c))
            .map(cell_as_string)
            .unwrap_or_default()
            .trim()
            .to_string();
        let factory = col_factory
            .and_then(|c| row.get(c))
            .map(cell_as_string)
            .unwrap_or_default()
            .trim()
            .to_string();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let price = col_price
            .and_then(|c| row.get(c))
            .map(cell_as_f64)
            .unwrap_or(0.0);
        let amount = if let Some(c) = col_amt {
            let direct = row.get(c).map(cell_as_f64).unwrap_or(0.0);
            if direct > 0.0 {
                (direct * 100.0).round() / 100.0
            } else {
                (qty * price * 100.0).round() / 100.0
            }
        } else {
            (qty * price * 100.0).round() / 100.0
        };

        let key = (
            normalize_text(&name),
            normalize_text(&spec),
            normalize_text(&factory),
        );
        if let Some(&idx) = group_map.get(&key) {
            raw_items[idx].qty += qty;
            raw_items[idx].amount = ((raw_items[idx].amount + amount) * 100.0).round() / 100.0;
        } else {
            let idx = raw_items.len();
            raw_items.push(RawSaleItem {
                row_no: r_idx + 1,
                name,
                spec,
                factory,
                qty,
                price,
                amount,
            });
            group_map.insert(key, idx);
        }
    }

    let voucher_date = detect_month_end_date(custom_date, sales_path.to_str());
    let v_no = voucher_no.unwrap_or("1").to_string();

    let mut voucher_rows = Vec::new();
    let mut unmatched_drugs = Vec::new();
    let mut total_credit_cents: i64 = 0;

    let mut credit_rows = Vec::new();
    for item in &raw_items {
        let matched = find_confirmed_external_inventory_match(
            item.row_no,
            confirmed_mappings,
            &template_data,
        )
        .or_else(|| {
            match_external_inventory_code(
                &item.name,
                &item.spec,
                &item.factory,
                &template_data.inventory_items,
                &factory_abbr,
            )
        });
        let (aux_code, aux_name, is_unmatched) = match matched {
            Some((code, name)) => (code, name, false),
            None => {
                unmatched_drugs.push(ExternalUnmatchedDrug {
                    id: item.row_no,
                    row_index: item.row_no,
                    name: item.name.clone(),
                    factory: item.factory.clone(),
                    spec: item.spec.clone(),
                    qty: item.qty,
                    amount: item.amount,
                    supplier: None,
                    candidates: build_external_inventory_candidates(
                        &item.name,
                        &item.spec,
                        &item.factory,
                        &template_data,
                    ),
                });
                (String::new(), item.name.clone(), true)
            }
        };

        let cost_amount = item.amount;

        let line_cents = (cost_amount * 100.0).round() as i64;
        total_credit_cents += line_cents;

        credit_rows.push(ExternalVoucherRow {
            row_no: 0,
            date: voucher_date.clone(),
            voucher_type: "记".to_string(),
            voucher_no: v_no.clone(),
            summary: "结转销售成本".to_string(),
            subject_code: "1405".to_string(),
            subject_name: "库存商品".to_string(),
            debit_amount: None,
            credit_amount: Some(cost_amount),
            exchange_rate: Some(1.0),
            currency: "人民币".to_string(),
            qty: Some(item.qty),
            aux_code,
            aux_name,
            settle_method: String::new(),
            bill_no: String::new(),
            occur_date: voucher_date.clone(),
            is_credit: true,
            is_unmatched,
        });
    }

    let total_cost = (total_credit_cents as f64) / 100.0;

    let summary_text = if let Ok(d) = NaiveDate::parse_from_str(&voucher_date, "%Y-%m-%d") {
        format!("结转{}年{}月西药销售成本", d.year(), d.month())
    } else {
        "结转西药销售成本".to_string()
    };

    voucher_rows.push(ExternalVoucherRow {
        row_no: 1,
        date: voucher_date.clone(),
        voucher_type: "记".to_string(),
        voucher_no: v_no.clone(),
        summary: summary_text,
        subject_code: "5401".to_string(),
        subject_name: "主营业务成本".to_string(),
        debit_amount: Some(total_cost),
        credit_amount: None,
        exchange_rate: Some(1.0),
        currency: "人民币".to_string(),
        qty: None,
        aux_code: String::new(),
        aux_name: String::new(),
        settle_method: String::new(),
        bill_no: String::new(),
        occur_date: voucher_date.clone(),
        is_credit: false,
        is_unmatched: false,
    });

    for (idx, mut cr) in credit_rows.into_iter().enumerate() {
        cr.row_no = idx + 2;
        voucher_rows.push(cr);
    }

    let out_file_path = resolve_external_output_path(
        sales_path,
        output_path,
        "表格迁账参考模板_西药出库_已生成.xlsx",
    );

    write_external_voucher_workbook(template_path, &out_file_path, &voucher_rows)?;

    let total_items = raw_items.len();
    let unmatched_count = unmatched_drugs.len();
    let matched_count = total_items.saturating_sub(unmatched_count);
    let match_rate = if total_items > 0 {
        ((matched_count as f64) / (total_items as f64) * 100.0 * 10.0).round() / 10.0
    } else {
        0.0
    };

    Ok(ExternalOutboundResult {
        success: true,
        output_file: out_file_path.to_string_lossy().to_string(),
        voucher_date,
        voucher_no: v_no,
        total_items,
        total_entries: voucher_rows.len(),
        matched_count,
        unmatched_count,
        match_rate,
        total_debit: total_cost,
        total_credit: total_cost,
        diff: 0.0,
        is_balanced: true,
        unmatched_drugs,
        voucher_rows,
    })
}

// ----------------------------------------------------
// 5. 外账结存数智能比对 (纯 Rust 原生实现)
// ----------------------------------------------------

pub fn generate_external_inventory_audit(
    ledger_path: &Path,
    west_path: &Path,
    tcm_path: &Path,
    hc_path: &Path,
    output_path: Option<&Path>,
    _config: Option<&ConfigData>,
) -> Result<ExternalAuditResult, String> {
    let ledger_data = load_external_quantity_ledger(ledger_path)?;
    let factory_abbr = get_merged_factory_abbr_map(_config);

    let mut wh_items = Vec::new();
    let mut seen_warehouse_paths = HashSet::new();
    for warehouse_path in [west_path, tcm_path, hc_path] {
        // 兼容测试或历史调用中重复传入同一张表，避免把同一库存重复计入。
        if seen_warehouse_paths.insert(warehouse_path.to_path_buf()) {
            // 外账结存核对与内账使用同一套库管报表解析规则，兼容 .xls/.xlsx
            // 以及药品、耗材等不同名称列。
            wh_items.extend(load_warehouse_items(warehouse_path)?);
        }
    }

    let mut warehouse_factory_counts: HashMap<(String, String, String), HashSet<String>> =
        HashMap::new();
    for warehouse in &wh_items {
        let key = warehouse_inventory_key(warehouse);
        warehouse_factory_counts
            .entry(key)
            .or_default()
            .insert(normalize_text(&warehouse.factory));
    }

    let mut records = Vec::new();
    let mut matched_ext_codes = HashSet::new();
    let mut matched_wh_indices = HashSet::new();

    for (wh_idx, wh_item) in wh_items.iter().enumerate() {
        if let Some((code, _)) = match_external_inventory_key(
            wh_item,
            &ledger_data.inventory_items,
            &factory_abbr,
            &warehouse_factory_counts,
        ) {
            matched_wh_indices.insert(wh_idx);
            matched_ext_codes.insert(code.clone());

            let (ext_price, ext_qty, ext_amt) = ledger_data
                .balance_by_code
                .get(&code)
                .cloned()
                .unwrap_or((0.0, 0.0, 0.0));

            let diff_qty = wh_item.qty - ext_qty;
            let status = if (diff_qty.abs()) < 0.0001 {
                "完全吻合".to_string()
            } else {
                "数量差异".to_string()
            };

            let std_item = ledger_data.inventory_by_code.get(&code);
            let display_name = std_item
                .map(|i| i.name.clone())
                .unwrap_or_else(|| wh_item.name.clone());
            let display_spec = std_item
                .map(|i| i.spec.clone())
                .unwrap_or_else(|| wh_item.spec.clone());

            records.push(ExternalAuditRecord {
                aux_code: code,
                name: display_name,
                spec: display_spec,
                factory: wh_item.factory.clone(),
                ext_qty,
                ext_price,
                ext_amount: ext_amt,
                wh_qty: wh_item.qty,
                diff_qty: (diff_qty * 100.0).round() / 100.0,
                status,
            });
        }
    }

    for (wh_idx, wh_item) in wh_items.iter().enumerate() {
        if !matched_wh_indices.contains(&wh_idx) {
            records.push(ExternalAuditRecord {
                aux_code: "-".to_string(),
                name: wh_item.name.clone(),
                spec: wh_item.spec.clone(),
                factory: wh_item.factory.clone(),
                ext_qty: 0.0,
                ext_price: 0.0,
                ext_amount: 0.0,
                wh_qty: wh_item.qty,
                diff_qty: wh_item.qty,
                status: "仅库管有".to_string(),
            });
        }
    }

    for (code, (ext_price, ext_qty, ext_amt)) in &ledger_data.balance_by_code {
        // 数量为 0 但金额非 0 的期末余额可能是历史尾差，不能在对账结果中静默丢失。
        if (ext_qty.abs() > 0.0001 || ext_amt.abs() > 0.0001) && !matched_ext_codes.contains(code) {
            let std_item = ledger_data.inventory_by_code.get(code);
            let display_name = std_item
                .map(|i| i.name.clone())
                .unwrap_or_else(|| format!("外账存货{}", code));
            let display_spec = std_item.map(|i| i.spec.clone()).unwrap_or_default();

            records.push(ExternalAuditRecord {
                aux_code: code.clone(),
                name: display_name,
                spec: display_spec,
                factory: "-".to_string(),
                ext_qty: *ext_qty,
                ext_price: *ext_price,
                ext_amount: *ext_amt,
                wh_qty: 0.0,
                diff_qty: -(*ext_qty),
                status: "仅外账有".to_string(),
            });
        }
    }

    records.sort_by_key(|r| match r.status.as_str() {
        "数量差异" => 1,
        "仅外账有" => 2,
        "仅库管有" => 3,
        _ => 4,
    });

    let total_items = records.len();
    let equal_count = records.iter().filter(|r| r.status == "完全吻合").count();
    let diff_count = records.iter().filter(|r| r.status == "数量差异").count();
    let ext_only_count = records.iter().filter(|r| r.status == "仅外账有").count();
    let wh_only_count = records.iter().filter(|r| r.status == "仅库管有").count();
    let match_rate = if total_items > 0 {
        ((equal_count as f64) / (total_items as f64) * 100.0 * 10.0).round() / 10.0
    } else {
        0.0
    };

    let out_file_path =
        resolve_external_output_path(ledger_path, output_path, "外账账实库存核对分析报告.xlsx");

    export_external_audit_excel(&out_file_path, &records)?;

    Ok(ExternalAuditResult {
        success: true,
        output_file: out_file_path.to_string_lossy().to_string(),
        total_items,
        equal_count,
        diff_count,
        ext_only_count,
        wh_only_count,
        match_rate,
        records,
    })
}

// ----------------------------------------------------
// 6. Excel 文件写出支持 (凭证 Sheet 与比对报告)
// ----------------------------------------------------

fn escape_xml(s: &str) -> String {
    s.chars()
        .filter(|c| {
            // XML 1.0 不允许控制字符；Excel 遇到这些字符会拒绝打开文件。
            matches!(*c, '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
        })
        .collect::<String>()
        .replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn xml_attribute_value<'a>(element: &'a str, attribute: &str) -> Option<&'a str> {
    let marker = format!("{}=\"", attribute);
    let start = element.find(&marker)? + marker.len();
    let end = element[start..].find('"')? + start;
    Some(&element[start..end])
}

fn replace_xml_attribute(element: &str, attribute: &str, value: &str) -> String {
    let marker = format!("{}=\"", attribute);
    if let Some(marker_pos) = element.find(&marker) {
        let value_start = marker_pos + marker.len();
        if let Some(relative_end) = element[value_start..].find('"') {
            let value_end = value_start + relative_end;
            let mut result = String::with_capacity(element.len() + value.len());
            result.push_str(&element[..value_start]);
            result.push_str(&escape_xml(value));
            result.push_str(&element[value_end..]);
            return result;
        }
    }

    let insert_at = if element.ends_with("/>") {
        element.len() - 2
    } else {
        element.len().saturating_sub(1)
    };
    format!(
        "{} {}=\"{}\"{}",
        &element[..insert_at],
        attribute,
        escape_xml(value),
        &element[insert_at..]
    )
}

fn remove_xml_attribute(element: &str, attribute: &str) -> String {
    let marker = format!(" {}=\"", attribute);
    let Some(marker_pos) = element.find(&marker) else {
        return element.to_string();
    };
    let value_start = marker_pos + marker.len();
    let Some(relative_end) = element[value_start..].find('"') else {
        return element.to_string();
    };
    let value_end = value_start + relative_end + 1;
    format!("{}{}", &element[..marker_pos], &element[value_end..])
}

fn normalize_open_tag(element: &str) -> String {
    if element.ends_with("/>") {
        format!("{}>", &element[..element.len() - 2])
    } else {
        element.to_string()
    }
}

fn find_row_xml(sheet_xml: &str, row_number: usize) -> Option<(usize, usize)> {
    let mut search_from = 0;
    while let Some(relative_start) = sheet_xml[search_from..].find("<row") {
        let start = search_from + relative_start;
        let open_end = sheet_xml[start..].find('>')? + start;
        let open_tag = &sheet_xml[start..=open_end];
        let row_value = row_number.to_string();
        if xml_attribute_value(open_tag, "r") == Some(row_value.as_str()) {
            if open_tag.ends_with("/>") {
                return Some((start, open_end + 1));
            }
            let close_tag = "</row>";
            let close_start = sheet_xml[open_end + 1..].find(close_tag)? + open_end + 1;
            return Some((start, close_start + close_tag.len()));
        }
        search_from = open_end + 1;
    }
    None
}

fn cell_column(cell_open_tag: &str) -> Option<String> {
    let reference = xml_attribute_value(cell_open_tag, "r")?;
    let column: String = reference
        .chars()
        .take_while(|c| c.is_ascii_alphabetic())
        .collect();
    (!column.is_empty()).then_some(column)
}

enum CellContent {
    Text(String),
    Number(String),
    Empty,
}

fn render_cell_with_template_style(
    cell_xml: &str,
    content: CellContent,
    style_override: Option<&str>,
) -> String {
    let Some(open_end) = cell_xml.find('>') else {
        return cell_xml.to_string();
    };
    let open_tag = normalize_open_tag(&cell_xml[..=open_end]);
    let mut open_tag = remove_xml_attribute(&open_tag, "t");
    if let Some(style_id) = style_override {
        // 外账模板的 42 号样式是黄色告警样式，用于提示待人工补全的辅助编码。
        open_tag = replace_xml_attribute(&open_tag, "s", style_id);
    }

    let body = match content {
        CellContent::Text(value) => {
            open_tag = replace_xml_attribute(&open_tag, "t", "inlineStr");
            format!("<is><t>{}</t></is>", escape_xml(&value))
        }
        CellContent::Number(value) => format!("<v>{}</v>", escape_xml(&value)),
        CellContent::Empty => String::new(),
    };

    format!("{}{}", open_tag, body) + "</c>"
}

fn cell_content_for_row(row: &ExternalVoucherRow, column: &str) -> CellContent {
    match column {
        "A" => CellContent::Text(row.date.clone()),
        "B" => CellContent::Text(row.voucher_type.clone()),
        "C" => CellContent::Text(row.voucher_no.clone()),
        "D" => CellContent::Text(row.summary.clone()),
        "E" => CellContent::Text(row.subject_code.clone()),
        "F" => CellContent::Text(row.subject_name.clone()),
        "G" => row
            .debit_amount
            .map(|value| CellContent::Number(format!("{value:.2}")))
            .unwrap_or(CellContent::Empty),
        "H" => row
            .credit_amount
            .map(|value| CellContent::Number(format!("{value:.2}")))
            .unwrap_or(CellContent::Empty),
        "I" | "J" => CellContent::Empty,
        "K" => row
            .qty
            .map(|value| CellContent::Number(format!("{value}")))
            .unwrap_or(CellContent::Empty),
        "L" => {
            if row.aux_code.is_empty() {
                CellContent::Empty
            } else {
                CellContent::Text(row.aux_code.clone())
            }
        }
        "M" => {
            if row.aux_name.is_empty() {
                CellContent::Empty
            } else {
                CellContent::Text(row.aux_name.clone())
            }
        }
        // N/O/P 由模板保留为空白，避免再次写入不应导入的字段。
        "N" | "O" | "P" => CellContent::Empty,
        _ => CellContent::Empty,
    }
}

fn render_voucher_row(prototype_xml: &str, row_number: usize, row: &ExternalVoucherRow) -> String {
    let Some(open_end) = prototype_xml.find('>') else {
        return prototype_xml.to_string();
    };
    let Some(close_start) = prototype_xml.rfind("</row>") else {
        return prototype_xml.to_string();
    };

    let mut row_open =
        replace_xml_attribute(&prototype_xml[..=open_end], "r", &row_number.to_string());
    row_open = replace_xml_attribute(&row_open, "spans", "1:16");

    let mut result = String::with_capacity(prototype_xml.len() + 128);
    result.push_str(&row_open);
    let mut cursor = open_end + 1;
    while cursor < close_start {
        let Some(relative_start) = prototype_xml[cursor..close_start].find("<c") else {
            break;
        };
        let cell_start = cursor + relative_start;
        let Some(cell_open_end_relative) = prototype_xml[cell_start..close_start].find('>') else {
            break;
        };
        let cell_open_end = cell_start + cell_open_end_relative;
        let cell_open_tag = &prototype_xml[cell_start..=cell_open_end];
        let cell_end = if cell_open_tag.ends_with("/>") {
            cell_open_end + 1
        } else {
            let Some(relative_cell_close) =
                prototype_xml[cell_open_end + 1..close_start].find("</c>")
            else {
                break;
            };
            cell_open_end + 1 + relative_cell_close + "</c>".len()
        };

        if let Some(column) = cell_column(cell_open_tag) {
            let template_cell = &prototype_xml[cell_start..cell_end];
            let warning_style = (row.is_unmatched && column == "L").then_some("42");
            let mut rendered = render_cell_with_template_style(
                template_cell,
                cell_content_for_row(row, &column),
                warning_style,
            );
            rendered = replace_xml_attribute(&rendered, "r", &format!("{}{}", column, row_number));
            result.push_str(&rendered);
        }
        cursor = cell_end;
    }
    result.push_str("</row>");
    result
}

fn normalize_zip_entry_path(target: &str) -> String {
    let mut parts = Vec::new();
    let normalized_target = target.replace('\\', "/");
    for part in normalized_target.split('/') {
        match part {
            "" | "." => {}
            ".." => {
                parts.pop();
            }
            value => parts.push(value),
        }
    }
    parts.join("/")
}

fn resolve_voucher_sheet_entry(archive: &mut zip::ZipArchive<std::fs::File>) -> String {
    let mut workbook_xml = String::new();
    if let Ok(mut entry) = archive.by_name("xl/workbook.xml") {
        let _ = entry.read_to_string(&mut workbook_xml);
    }
    let mut relationships_xml = String::new();
    if let Ok(mut entry) = archive.by_name("xl/_rels/workbook.xml.rels") {
        let _ = entry.read_to_string(&mut relationships_xml);
    }

    let Some(sheet_pos) = workbook_xml.find("name=\"凭证\"") else {
        return "xl/worksheets/sheet6.xml".to_string();
    };
    let sheet_fragment = &workbook_xml[sheet_pos..];
    let Some(rid_pos) = sheet_fragment.find("r:id=\"") else {
        return "xl/worksheets/sheet6.xml".to_string();
    };
    let rid_fragment = &sheet_fragment[rid_pos + "r:id=\"".len()..];
    let Some(rid_end) = rid_fragment.find('"') else {
        return "xl/worksheets/sheet6.xml".to_string();
    };
    let rid = &rid_fragment[..rid_end];
    let relationship_marker = format!("Id=\"{}\"", rid);
    let Some(relationship_pos) = relationships_xml.find(&relationship_marker) else {
        return "xl/worksheets/sheet6.xml".to_string();
    };
    let relationship_fragment = &relationships_xml[relationship_pos..];
    let Some(target_pos) = relationship_fragment.find("Target=\"") else {
        return "xl/worksheets/sheet6.xml".to_string();
    };
    let target_fragment = &relationship_fragment[target_pos + "Target=\"".len()..];
    let Some(target_end) = target_fragment.find('"') else {
        return "xl/worksheets/sheet6.xml".to_string();
    };
    let target = &target_fragment[..target_end];
    let path = if target.starts_with('/') {
        target.trim_start_matches('/').to_string()
    } else {
        format!("xl/{}", target)
    };
    normalize_zip_entry_path(&path)
}

fn write_external_voucher_workbook(
    template_path: &Path,
    output_path: &Path,
    voucher_rows: &[ExternalVoucherRow],
) -> Result<(), String> {
    let template_abs =
        std::fs::canonicalize(template_path).map_err(|e| format!("解析模板路径失败: {}", e))?;
    let output_abs = if output_path.exists() {
        std::fs::canonicalize(output_path).map_err(|e| format!("解析输出路径失败: {}", e))?
    } else if output_path.is_absolute() {
        output_path.to_path_buf()
    } else {
        std::env::current_dir()
            .map_err(|e| format!("获取当前目录失败: {}", e))?
            .join(output_path)
    };
    if template_abs == output_abs {
        return Err("输出文件不能覆盖外账模板，请选择其他文件名".to_string());
    }

    let file =
        std::fs::File::open(template_path).map_err(|e| format!("打开模板文件失败: {}", e))?;
    let mut archive =
        zip::ZipArchive::new(file).map_err(|e| format!("解析模板 zip 失败: {}", e))?;

    // 1. 动态寻找凭证 Sheet 的 XML 内部路径。
    let voucher_sheet_entry_name = resolve_voucher_sheet_entry(&mut archive);

    // 2. 读取原凭证 Sheet，保留表头、分录样式及尾部元数据，只替换分录数据。
    let mut original_sheet_xml = String::new();
    {
        let mut sheet_entry = archive.by_name(&voucher_sheet_entry_name).map_err(|e| {
            format!(
                "未在模板中找到凭证工作表 {}: {}",
                voucher_sheet_entry_name, e
            )
        })?;
        sheet_entry
            .read_to_string(&mut original_sheet_xml)
            .map_err(|e| format!("读取凭证工作表失败: {}", e))?;
    }

    let (_, header_end_idx) = find_row_xml(&original_sheet_xml, 3)
        .ok_or_else(|| "模板凭证 sheet 中未找到第 3 行表头".to_string())?;

    let sheet_data_end_tag = "</sheetData>";
    let tail_start_idx = original_sheet_xml
        .find(sheet_data_end_tag)
        .ok_or_else(|| "模板凭证 sheet 中未找到 </sheetData>".to_string())?;

    let header_xml = &original_sheet_xml[..header_end_idx];
    let tail_xml = &original_sheet_xml[tail_start_idx..];

    let debit_prototype = find_row_xml(&original_sheet_xml, 4)
        .map(|(start, end)| original_sheet_xml[start..end].to_string())
        .ok_or_else(|| "模板凭证 sheet 中未找到借方分录样式行".to_string())?;
    let credit_prototype = find_row_xml(&original_sheet_xml, 6)
        .map(|(start, end)| original_sheet_xml[start..end].to_string())
        .unwrap_or_else(|| debit_prototype.clone());

    // 构建所有分录行 XML。每一行都从模板样式行克隆，I/J/P 等字段只清空内容而不丢失样式。
    let mut new_rows_xml = String::with_capacity(voucher_rows.len() * 500);
    for (idx, row) in voucher_rows.iter().enumerate() {
        let r = idx + 4;
        let prototype = if row.is_credit {
            &credit_prototype
        } else {
            &debit_prototype
        };
        new_rows_xml.push_str(&render_voucher_row(prototype, r, row));
    }

    let mut new_sheet_xml = format!("{}{}{}", header_xml, new_rows_xml, tail_xml);
    if let Some(dimension_start) = new_sheet_xml.find("<dimension ") {
        if let Some(dimension_end_rel) = new_sheet_xml[dimension_start..].find('>') {
            let dimension_end = dimension_start + dimension_end_rel;
            let dimension_tag = &new_sheet_xml[dimension_start..=dimension_end];
            let last_row = 3 + voucher_rows.len();
            let replacement =
                replace_xml_attribute(dimension_tag, "ref", &format!("A1:P{}", last_row.max(3)));
            new_sheet_xml.replace_range(dimension_start..=dimension_end, &replacement);
        }
    }

    // 3. 将原模板除凭证外的所有 sheet 及元数据原封不动写入输出文件 (100% 字节级保真)
    if let Some(parent) = output_path.parent().filter(|p| !p.as_os_str().is_empty()) {
        std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    let out_file =
        std::fs::File::create(output_path).map_err(|e| format!("创建输出文件失败: {}", e))?;
    let mut zip_writer = zip::ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let mut entry = archive
            .by_index(i)
            .map_err(|e| format!("读取 zip entry 失败: {}", e))?;
        let entry_name = entry.name().to_string();

        let options = zip::write::SimpleFileOptions::default()
            .compression_method(entry.compression())
            .unix_permissions(entry.unix_mode().unwrap_or(0o644));

        if entry_name == voucher_sheet_entry_name {
            zip_writer
                .start_file(&entry_name, options)
                .map_err(|e| format!("写入新凭证 sheet 失败: {}", e))?;
            std::io::Write::write_all(&mut zip_writer, new_sheet_xml.as_bytes())
                .map_err(|e| format!("写入凭证 xml 数据失败: {}", e))?;
        } else {
            zip_writer
                .start_file(&entry_name, options)
                .map_err(|e| format!("写入 entry {} 失败: {}", entry_name, e))?;
            std::io::copy(&mut entry, &mut zip_writer)
                .map_err(|e| format!("复制 entry {} 失败: {}", entry_name, e))?;
        }
    }

    zip_writer
        .finish()
        .map_err(|e| format!("完成 zip 压缩失败: {}", e))?;
    Ok(())
}

fn export_external_audit_excel(
    output_path: &Path,
    records: &[ExternalAuditRecord],
) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet
        .set_name("外账账实库存核对明细")
        .map_err(|e| e.to_string())?;

    let font_family = "微软雅黑";

    let title_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(14)
        .set_bold();

    let header_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_bold()
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin)
        .set_background_color(Color::RGB(0xEAEEF4));

    let text_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_border(FormatBorder::Thin);

    let center_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_align(FormatAlign::Center)
        .set_border(FormatBorder::Thin);

    let number_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_num_format("#,##0.##")
        .set_align(FormatAlign::Right)
        .set_border(FormatBorder::Thin);

    let money_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_num_format("#,##0.00")
        .set_align(FormatAlign::Right)
        .set_border(FormatBorder::Thin);

    let warn_row_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_background_color(Color::RGB(0xFDEAEA))
        .set_border(FormatBorder::Thin);

    worksheet
        .write_string_with_format(0, 0, "外账账实库存核对分析报告", &title_format)
        .map_err(|e| e.to_string())?;

    let headers = [
        "序号",
        "外账存货编码",
        "药品标准名称",
        "规格型号",
        "生产厂家",
        "外账账面数量",
        "外账成本单价",
        "外账账面金额",
        "库管在库数量",
        "数量差异 (库管-外账)",
        "核对状态",
    ];

    for (c, h) in headers.iter().enumerate() {
        worksheet
            .write_string_with_format(2, c as u16, *h, &header_format)
            .map_err(|e| e.to_string())?;
    }

    for (idx, r) in records.iter().enumerate() {
        let row_idx = (idx + 3) as u32;
        let is_diff = r.status == "数量差异";
        let base_fmt = if is_diff {
            &warn_row_format
        } else {
            &text_format
        };

        worksheet
            .write_number_with_format(row_idx, 0, (idx + 1) as f64, &center_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_string_with_format(row_idx, 1, &r.aux_code, &center_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_string_with_format(row_idx, 2, &r.name, base_fmt)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_string_with_format(row_idx, 3, &r.spec, base_fmt)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_string_with_format(row_idx, 4, &r.factory, base_fmt)
            .map_err(|e| e.to_string())?;

        worksheet
            .write_number_with_format(row_idx, 5, r.ext_qty, &number_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_number_with_format(row_idx, 6, r.ext_price, &money_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_number_with_format(row_idx, 7, r.ext_amount, &money_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_number_with_format(row_idx, 8, r.wh_qty, &number_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_number_with_format(row_idx, 9, r.diff_qty, &number_format)
            .map_err(|e| e.to_string())?;
        worksheet
            .write_string_with_format(row_idx, 10, &r.status, &center_format)
            .map_err(|e| e.to_string())?;
    }

    worksheet
        .set_column_width(1, 14)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(2, 26)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(3, 16)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(4, 22)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(5, 13)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(7, 14)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(8, 13)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(9, 18)
        .map_err(|e| e.to_string())?;
    worksheet
        .set_column_width(10, 12)
        .map_err(|e| e.to_string())?;

    workbook
        .save(output_path)
        .map_err(|e| format!("保存核对报告失败: {}", e))?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sample_inventory_item() -> ExternalInventoryItem {
        ExternalInventoryItem {
            code: "00042".to_string(),
            name: "测试药片".to_string(),
            spec: "10mg*10片".to_string(),
            spec_key: normalize_inventory_spec_key("10mg*10片"),
            norm_spec: "10mg*10片".to_string(),
            norm_name: "测试药片".to_string(),
            clean_name: "测试药片".to_string(),
        }
    }

    fn inventory_item(code: &str, name: &str, spec: &str) -> ExternalInventoryItem {
        ExternalInventoryItem {
            code: code.to_string(),
            name: name.to_string(),
            spec: spec.to_string(),
            spec_key: normalize_inventory_spec_key(spec),
            norm_spec: normalize_text(spec),
            norm_name: normalize_text(name),
            clean_name: clean_drug_name(name),
        }
    }

    fn warehouse_item(name: &str, spec: &str, dosage_form: &str, factory: &str) -> WarehouseItem {
        WarehouseItem {
            name: name.to_string(),
            spec: spec.to_string(),
            dosage_form: dosage_form.to_string(),
            factory: factory.to_string(),
            unit: "盒".to_string(),
            qty: 1.0,
            price: 1.0,
            amount: 1.0,
        }
    }

    fn warehouse_factory_counts(
        warehouses: &[WarehouseItem],
    ) -> HashMap<(String, String, String), HashSet<String>> {
        let mut counts = HashMap::new();
        for warehouse in warehouses {
            counts
                .entry(warehouse_inventory_key(warehouse))
                .or_insert_with(HashSet::new)
                .insert(normalize_text(&warehouse.factory));
        }
        counts
    }

    #[test]
    fn external_inventory_key_prefers_exact_spec_and_factory_tag() {
        let items = vec![
            inventory_item("00004", "盐酸贝那普利片", "10mg"),
            inventory_item("00965", "盐酸贝那普利片（湖南千金）", "10mg*28片/盒"),
        ];
        let warehouse =
            warehouse_item("盐酸贝那普利片", "10mg*28片/盒", "片剂", "湖南千金湘江药业");
        let warehouses = vec![warehouse.clone()];
        let counts = warehouse_factory_counts(&warehouses);
        let factory_map = get_merged_factory_abbr_map(None);

        assert_eq!(
            match_external_inventory_key(&warehouse, &items, &factory_map, &counts)
                .map(|(code, _)| code),
            Some("00965".to_string())
        );
    }

    #[test]
    fn external_inventory_key_prefers_factory_item_over_generic_item() {
        let items = vec![
            inventory_item("00999", "桑寄生", "1克*1000克/袋"),
            inventory_item("01110", "桑寄生（蕴德）", "1克*1000克/袋"),
        ];
        let warehouse = warehouse_item("桑寄生", "1克*1000克/袋", "饮片", "河北蕴德药业有限公司");
        let counts = warehouse_factory_counts(std::slice::from_ref(&warehouse));
        let factory_map = get_merged_factory_abbr_map(None);

        assert_eq!(
            match_external_inventory_key(&warehouse, &items, &factory_map, &counts)
                .map(|(code, _)| code),
            Some("01110".to_string())
        );

        let wrong_factory =
            warehouse_item("桑寄生", "1克*1000克/袋", "饮片", "河北国瑞堂药业有限公司");
        let wrong_counts = warehouse_factory_counts(std::slice::from_ref(&wrong_factory));
        assert!(
            match_external_inventory_key(&wrong_factory, &items, &factory_map, &wrong_counts,)
                .is_none()
        );
    }

    #[test]
    fn external_inventory_key_does_not_use_contains_or_wrong_dosage() {
        let items = vec![inventory_item("00051", "地西泮注射液", "2ml：10mg")];
        let warehouse = warehouse_item(
            "地西泮注射液",
            "10mg*10支/盒",
            "注射剂",
            "国药集团容生制药有限公司",
        );
        let counts = warehouse_factory_counts(std::slice::from_ref(&warehouse));
        let factory_map = get_merged_factory_abbr_map(None);
        assert!(match_external_inventory_key(&warehouse, &items, &factory_map, &counts).is_none());

        let dosage_mismatch = warehouse_item(
            "地西泮注射液",
            "2ml：10mg",
            "片剂",
            "国药集团容生制药有限公司",
        );
        let mismatch_counts = warehouse_factory_counts(std::slice::from_ref(&dosage_mismatch));
        assert!(match_external_inventory_key(
            &dosage_mismatch,
            &items,
            &factory_map,
            &mismatch_counts,
        )
        .is_none());
    }

    #[test]
    fn external_inventory_key_rejects_untagged_factory_ambiguity() {
        let items = vec![inventory_item("01037", "盐酸异丙嗪注射液", "50mg*10支/盒")];
        let first = warehouse_item(
            "盐酸异丙嗪注射液",
            "50mg*10支/盒",
            "注射剂",
            "武汉福星生物药业有限公司",
        );
        let second = warehouse_item(
            "盐酸异丙嗪注射液",
            "50mg*10支/盒",
            "注射剂",
            "遂成药业股份有限公司",
        );
        let counts = warehouse_factory_counts(&[first.clone(), second.clone()]);
        let factory_map = get_merged_factory_abbr_map(None);

        assert!(match_external_inventory_key(&first, &items, &factory_map, &counts).is_none());
        assert!(match_external_inventory_key(&second, &items, &factory_map, &counts).is_none());
    }

    #[test]
    fn confirmed_external_mapping_accepts_only_template_inventory_codes() {
        let item = sample_inventory_item();
        let mut template_data = ExternalTemplateData::default();
        template_data
            .inventory_by_code
            .insert(item.code.clone(), item);

        let valid = [ConfirmedExternalInventoryMapping {
            row_index: 12,
            aux_code: "00042".to_string(),
        }];
        assert_eq!(
            find_confirmed_external_inventory_match(12, Some(&valid), &template_data)
                .map(|(code, _)| code),
            Some("00042".to_string())
        );

        let invalid = [ConfirmedExternalInventoryMapping {
            row_index: 12,
            aux_code: "99999".to_string(),
        }];
        assert!(
            find_confirmed_external_inventory_match(12, Some(&invalid), &template_data).is_none()
        );
    }

    #[test]
    fn external_candidate_search_scores_code_name_and_spec() {
        let item = sample_inventory_item();
        assert!(external_candidate_score("00042", &item) > 0);
        assert!(external_candidate_score("测试药", &item) > 0);
        assert!(external_candidate_score("10mg", &item) > 0);
        assert_eq!(external_candidate_score("", &item), 0);
    }

    #[test]
    fn external_relative_output_stays_next_to_source_file() {
        let relative = resolve_external_output_path(
            Path::new("/var/data/入库单.xlsx"),
            Some(Path::new("生成结果.xlsx")),
            "默认结果.xlsx",
        );
        assert_eq!(relative, Path::new("/var/data/生成结果.xlsx"));

        let default =
            resolve_external_output_path(Path::new("/var/data/入库单.xlsx"), None, "默认结果.xlsx");
        assert_eq!(default, Path::new("/var/data/默认结果.xlsx"));

        let absolute = resolve_external_output_path(
            Path::new("/var/data/入库单.xlsx"),
            Some(Path::new("/tmp/用户指定结果.xlsx")),
            "默认结果.xlsx",
        );
        assert_eq!(absolute, Path::new("/tmp/用户指定结果.xlsx"));
    }
}
