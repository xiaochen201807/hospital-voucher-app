use crate::core::config::ConfigData;
use crate::core::excel_utils::{cell_as_f64, cell_as_string, open_excel};
use calamine::{Data, Reader};
use chrono::{Datelike, Local, NaiveDate};
use regex::Regex;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
use std::io::Read;
use std::path::{Path, PathBuf};

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
                        if let Some(first_of_next) = NaiveDate::from_ymd_opt(next_year, next_month, 1) {
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
    pub row_index: usize,
    pub name: String,
    pub factory: String,
    pub spec: String,
    pub qty: f64,
    pub amount: f64,
    pub supplier: Option<String>,
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
    pub match_rate: f64,
    pub total_debit: f64,
    pub total_credit: f64,
    pub diff: f64,
    pub is_balanced: bool,
    pub unmatched_drugs: Vec<ExternalUnmatchedDrug>,
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
                let aux_type = col_type.and_then(|c| data_row.get(c)).map(cell_as_string).unwrap_or_default();
                let aux_code = col_code.and_then(|c| data_row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
                let aux_name = col_name.and_then(|c| data_row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
                let aux_spec = col_spec.and_then(|c| data_row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();

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
                    data.supplier_map.insert(norm.clone(), (formatted_code.clone(), aux_name.clone()));
                    let brief = norm
                        .replace("有限公司", "")
                        .replace("股份有限公司", "")
                        .replace("有限责任公司", "")
                        .replace("医药", "")
                        .replace("石家庄", "");
                    if !brief.is_empty() {
                        data.supplier_map.insert(brief, (formatted_code.clone(), aux_name.clone()));
                    }
                } else if aux_type.contains("存货") {
                    let norm_name = normalize_text(&aux_name);
                    let clean = clean_drug_name(&aux_name);
                    let item = ExternalInventoryItem {
                        code: formatted_code.clone(),
                        name: aux_name.clone(),
                        spec: aux_spec,
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

    if let Some(bal_sheet_name) = excel.sheet_names().into_iter().find(|s| s.contains("辅助余额表")) {
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
    m.insert("长春高新技术产业(集团)股份有限公司", "金赛");
    m
}

fn get_merged_factory_abbr_map(config: Option<&ConfigData>) -> HashMap<String, String> {
    let mut m = HashMap::new();
    for (k, v) in get_default_factory_abbr_map() {
        m.insert(k.to_string(), v.to_string());
    }
    if let Some(cfg) = config {
        for (k, v) in &cfg.factory_abbreviations {
            m.insert(k.clone(), v.clone());
        }
    }
    m
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

    let abbr = factory_map.iter().find_map(|(full, b)| {
        if norm_factory.contains(full) || norm_factory.contains(b) {
            Some(b.as_str())
        } else {
            None
        }
    });

    for it in items {
        if it.norm_name == norm_name {
            return Some((it.code.clone(), it.name.clone()));
        }
    }

    if let Some(b) = abbr {
        let candidate_with_factory = format!("{}({})", clean_name, b);
        for it in items {
            if it.norm_name == candidate_with_factory {
                return Some((it.code.clone(), it.name.clone()));
            }
        }
    }

    if !norm_spec.is_empty() {
        for it in items {
            if (it.norm_name == clean_name || it.clean_name == clean_name)
                && (it.spec.contains(&norm_spec) || norm_spec.contains(&it.spec))
            {
                return Some((it.code.clone(), it.name.clone()));
            }
        }
    }

    let matches: Vec<&ExternalInventoryItem> = items
        .iter()
        .filter(|it| it.clean_name == clean_name)
        .collect();

    if matches.len() == 1 {
        return Some((matches[0].code.clone(), matches[0].name.clone()));
    }

    for it in items {
        if (it.norm_name.contains(&clean_name) || clean_name.contains(&it.norm_name))
            && (norm_spec.is_empty() || it.spec.contains(&norm_spec) || norm_spec.contains(&it.spec))
        {
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

    for (k, v) in supplier_map {
        if norm.contains(k) || k.contains(&norm) {
            return Some(v.clone());
        }
    }

    None
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
) -> Result<ExternalInboundResult, String> {
    let template_data = load_external_template(template_path)?;
    let factory_abbr = get_merged_factory_abbr_map(_config);

    let mut in_excel = open_excel(inbound_path)?;
    let in_sheet = in_excel
        .sheet_names()
        .into_iter()
        .next()
        .ok_or_else(|| "入库单没有工作表".to_string())?;

    let in_range = in_excel
        .worksheet_range(&in_sheet)
        .map_err(|e| format!("读取入库单失败: {}", e))?;

    let rows: Vec<Vec<Data>> = in_range.rows().map(|r| r.to_vec()).collect();
    if rows.is_empty() {
        return Err("入库单内容为空".to_string());
    }

    let mut header_idx = 0;
    let mut col_name = None;
    let mut col_spec = None;
    let mut col_factory = None;
    let mut col_qty = None;
    let mut col_price = None;
    let mut col_amt = None;
    let mut col_supplier = None;

    for (r_idx, row) in rows.iter().enumerate().take(10) {
        for (c_idx, cell) in row.iter().enumerate() {
            let s = cell_as_string(cell);
            if s.contains("通用名") || s.contains("药品名称") || s.contains("品名") {
                col_name = Some(c_idx);
            } else if s.contains("规格") {
                col_spec = Some(c_idx);
            } else if s.contains("厂家") || s.contains("生产企业") {
                col_factory = Some(c_idx);
            } else if s.contains("入库数量") || s == "数量" || s.contains("实收数量") {
                col_qty = Some(c_idx);
            } else if s.contains("进价单价") || s.contains("购进单价") || s == "进价" {
                col_price = Some(c_idx);
            } else if s.contains("进价金额") || s.contains("购进金额") || s.contains("金额(进)") {
                col_amt = Some(c_idx);
            } else if s.contains("供货单位") || s.contains("供应商") || s.contains("供货商") {
                col_supplier = Some(c_idx);
            }
        }
        if col_name.is_some() && col_qty.is_some() && (col_amt.is_some() || col_price.is_some()) {
            header_idx = r_idx;
            break;
        }
    }

    let col_name = col_name.ok_or_else(|| "入库单中未识别到【药品名称】列".to_string())?;
    let col_qty = col_qty.ok_or_else(|| "入库单中未识别到【数量】列".to_string())?;

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
        let name = cell_as_string(row.get(col_name).unwrap_or(&Data::Empty)).trim().to_string();
        if name.is_empty() || name.contains("合计") || name.contains("总计") || name.contains("制表") || name.starts_with("报表") {
            continue;
        }

        let spec = col_spec.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let factory = col_factory.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let amount = if let Some(c) = col_amt {
            let direct_amt = row.get(c).map(cell_as_f64).unwrap_or(0.0);
            if direct_amt > 0.0 {
                (direct_amt * 100.0).round() / 100.0
            } else {
                let price = col_price.and_then(|cp| row.get(cp)).map(cell_as_f64).unwrap_or(0.0);
                (qty * price * 100.0).round() / 100.0
            }
        } else {
            let price = col_price.and_then(|cp| row.get(cp)).map(cell_as_f64).unwrap_or(0.0);
            (qty * price * 100.0).round() / 100.0
        };

        let supplier = col_supplier.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let s_key = if supplier.is_empty() { "未指定供应商".to_string() } else { supplier.clone() };

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

            let (aux_code, aux_name, is_unmatched) = match match_external_inventory_code(
                &item.name,
                &item.spec,
                &item.factory,
                &template_data.inventory_items,
                &factory_abbr,
            ) {
                Some((code, name)) => (code, name, false),
                None => {
                    unmatched_drugs.push(ExternalUnmatchedDrug {
                        row_index: item.row_no,
                        name: item.name.clone(),
                        factory: item.factory.clone(),
                        spec: item.spec.clone(),
                        qty: item.qty,
                        amount: item.amount,
                        supplier: Some(s_name.clone()),
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

        let (sup_code, sup_name) = match match_external_supplier_code(s_name, &template_data.supplier_map) {
            Some((code, name)) => (code, name),
            None => (String::new(), s_name.clone()),
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
            is_unmatched: false,
        });
        row_counter += 1;
    }

    let total_debit = (total_debit_cents as f64) / 100.0;
    let total_credit = (total_credit_cents as f64) / 100.0;
    let diff = ((total_debit_cents - total_credit_cents) as f64) / 100.0;
    let is_balanced = total_debit_cents == total_credit_cents;

    let out_file_path = output_path
        .map(|p| p.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("表格迁账参考模板_西药入库_已生成.xlsx"));

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
        match_rate,
        total_debit,
        total_credit,
        diff,
        is_balanced,
        unmatched_drugs,
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
            } else if s.contains("厂家") || s.contains("生产企业") || s.contains("制药厂") {
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
        let name = cell_as_string(row.get(col_name).unwrap_or(&Data::Empty)).trim().to_string();
        if name.is_empty() || name.contains("合计") || name.contains("总计") || name.contains("制表") {
            continue;
        }

        let spec = col_spec.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let factory = col_factory.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);
        let price = col_price.and_then(|c| row.get(c)).map(cell_as_f64).unwrap_or(0.0);
        let amount = if let Some(c) = col_amt {
            let direct = row.get(c).map(cell_as_f64).unwrap_or(0.0);
            if direct > 0.0 { (direct * 100.0).round() / 100.0 } else { (qty * price * 100.0).round() / 100.0 }
        } else {
            (qty * price * 100.0).round() / 100.0
        };

        let key = (normalize_text(&name), normalize_text(&spec), normalize_text(&factory));
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
        let (aux_code, aux_name, is_unmatched) = match match_external_inventory_code(
            &item.name,
            &item.spec,
            &item.factory,
            &template_data.inventory_items,
            &factory_abbr,
        ) {
            Some((code, name)) => (code, name, false),
            None => {
                unmatched_drugs.push(ExternalUnmatchedDrug {
                    row_index: item.row_no,
                    name: item.name.clone(),
                    factory: item.factory.clone(),
                    spec: item.spec.clone(),
                    qty: item.qty,
                    amount: item.amount,
                    supplier: None,
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

    let out_file_path = output_path
        .map(|p| p.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("表格迁账参考模板_西药出库_已生成.xlsx"));

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
    template_path: &Path,
    warehouse_path: &Path,
    output_path: Option<&Path>,
    _config: Option<&ConfigData>,
) -> Result<ExternalAuditResult, String> {
    let template_data = load_external_template(template_path)?;
    let factory_abbr = get_merged_factory_abbr_map(_config);

    let mut wh_excel = open_excel(warehouse_path)?;
    let wh_sheet = wh_excel
        .sheet_names()
        .into_iter()
        .next()
        .ok_or_else(|| "库管库存报表为空".to_string())?;

    let wh_range = wh_excel
        .worksheet_range(&wh_sheet)
        .map_err(|e| format!("读取库管报表失败: {}", e))?;

    let wh_rows: Vec<Vec<Data>> = wh_range.rows().map(|r| r.to_vec()).collect();
    if wh_rows.is_empty() {
        return Err("库管报表无内容".to_string());
    }

    let mut header_idx = 0;
    let mut col_name = None;
    let mut col_spec = None;
    let mut col_factory = None;
    let mut col_qty = None;

    for (r_idx, row) in wh_rows.iter().enumerate().take(10) {
        for (c_idx, cell) in row.iter().enumerate() {
            let s = cell_as_string(cell);
            if s.contains("药品名称") || s.contains("通用名") || s.contains("品名") {
                col_name = Some(c_idx);
            } else if s.contains("规格") {
                col_spec = Some(c_idx);
            } else if s.contains("厂家") || s.contains("产地") {
                col_factory = Some(c_idx);
            } else if s.contains("在库数量") || s.contains("结存数量") || s.contains("现存量") || s.contains("数量") {
                col_qty = Some(c_idx);
            }
        }
        if col_name.is_some() && col_qty.is_some() {
            header_idx = r_idx;
            break;
        }
    }

    let col_name = col_name.ok_or_else(|| "库管报表中未识别到【药品名称】列".to_string())?;
    let col_qty = col_qty.ok_or_else(|| "库管报表中未识别到【数量】列".to_string())?;

    struct WhItem {
        name: String,
        spec: String,
        factory: String,
        qty: f64,
    }

    let mut wh_items = Vec::new();
    for row in wh_rows.iter().skip(header_idx + 1) {
        let name = cell_as_string(row.get(col_name).unwrap_or(&Data::Empty)).trim().to_string();
        if name.is_empty() || name.contains("合计") || name.contains("总计") {
            continue;
        }

        let spec = col_spec.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let factory = col_factory.and_then(|c| row.get(c)).map(cell_as_string).unwrap_or_default().trim().to_string();
        let qty = row.get(col_qty).map(cell_as_f64).unwrap_or(0.0);

        wh_items.push(WhItem {
            name,
            spec,
            factory,
            qty,
        });
    }

    let mut records = Vec::new();
    let mut matched_ext_codes = HashSet::new();
    let mut matched_wh_indices = HashSet::new();

    for (wh_idx, wh_item) in wh_items.iter().enumerate() {
        if let Some((code, _)) = match_external_inventory_code(
            &wh_item.name,
            &wh_item.spec,
            &wh_item.factory,
            &template_data.inventory_items,
            &factory_abbr,
        ) {
            matched_wh_indices.insert(wh_idx);
            matched_ext_codes.insert(code.clone());

            let (ext_price, ext_qty, ext_amt) = template_data
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

            let std_item = template_data.inventory_by_code.get(&code);
            let display_name = std_item.map(|i| i.name.clone()).unwrap_or_else(|| wh_item.name.clone());
            let display_spec = std_item.map(|i| i.spec.clone()).unwrap_or_else(|| wh_item.spec.clone());

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

    for (code, (ext_price, ext_qty, ext_amt)) in &template_data.balance_by_code {
        if *ext_qty > 0.0 && !matched_ext_codes.contains(code) {
            let std_item = template_data.inventory_by_code.get(code);
            let display_name = std_item.map(|i| i.name.clone()).unwrap_or_else(|| format!("外账存货{}", code));
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

    let out_file_path = output_path
        .map(|p| p.to_path_buf())
        .unwrap_or_else(|| PathBuf::from("外账账实库存核对分析报告.xlsx"));

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
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

fn write_external_voucher_workbook(
    template_path: &Path,
    output_path: &Path,
    voucher_rows: &[ExternalVoucherRow],
) -> Result<(), String> {
    let file = std::fs::File::open(template_path).map_err(|e| format!("打开模板文件失败: {}", e))?;
    let mut archive = zip::ZipArchive::new(file).map_err(|e| format!("解析模板 zip 失败: {}", e))?;

    // 1. 动态寻找凭证 sheet 的 xml 内部路径 (通常是 xl/worksheets/sheet6.xml)
    let mut voucher_sheet_entry_name = "xl/worksheets/sheet6.xml".to_string();
    let mut wb_xml = String::new();
    if let Ok(mut wb_entry) = archive.by_name("xl/workbook.xml") {
        let _ = wb_entry.read_to_string(&mut wb_xml);
    }
    let mut rels_xml = String::new();
    if let Ok(mut rels_entry) = archive.by_name("xl/_rels/workbook.xml.rels") {
        let _ = rels_entry.read_to_string(&mut rels_xml);
    }

    if let Some(pos) = wb_xml.find("name=\"凭证\"") {
        let rest = &wb_xml[pos..];
        if let Some(rid_pos) = rest.find("r:id=\"") {
            let rid_str = &rest[rid_pos + 6..];
            if let Some(rid_end) = rid_str.find('"') {
                let rid = &rid_str[..rid_end];
                let search_rid = format!("Id=\"{}\"", rid);
                if let Some(r_pos) = rels_xml.find(&search_rid) {
                    let r_rest = &rels_xml[r_pos..];
                    if let Some(target_pos) = r_rest.find("Target=\"") {
                        let target_str = &r_rest[target_pos + 8..];
                        if let Some(target_end) = target_str.find('"') {
                            let target = &target_str[..target_end];
                            let full_target = if target.starts_with('/') {
                                target.trim_start_matches('/').to_string()
                            } else {
                                format!("xl/{}", target)
                            };
                            voucher_sheet_entry_name = full_target;
                        }
                    }
                }
            }
        }
    }

    // 2. 读取原凭证 sheet 的 xml，保留前 3 行表头及尾部标签，只替换分录数据
    let mut original_sheet_xml = String::new();
    {
        let mut sheet_entry = archive
            .by_name(&voucher_sheet_entry_name)
            .map_err(|e| format!("未在模板中找到凭证工作表 {}: {}", voucher_sheet_entry_name, e))?;
        sheet_entry
            .read_to_string(&mut original_sheet_xml)
            .map_err(|e| format!("读取凭证工作表失败: {}", e))?;
    }

    let row3_tag = "<row r=\"3\"";
    let row3_pos = original_sheet_xml
        .find(row3_tag)
        .ok_or_else(|| "模板凭证 sheet 中未找到第 3 行表头".to_string())?;
    let row3_end_tag = "</row>";
    let row3_end_rel = original_sheet_xml[row3_pos..]
        .find(row3_end_tag)
        .ok_or_else(|| "模板凭证 sheet 第 3 行闭合标签未找到".to_string())?;
    let header_end_idx = row3_pos + row3_end_rel + row3_end_tag.len();

    let sheet_data_end_tag = "</sheetData>";
    let tail_start_idx = original_sheet_xml
        .find(sheet_data_end_tag)
        .ok_or_else(|| "模板凭证 sheet 中未找到 </sheetData>".to_string())?;

    let header_xml = &original_sheet_xml[..header_end_idx];
    let tail_xml = &original_sheet_xml[tail_start_idx..];

    // 构建所有分录行 XML (从第 4 行开始，I/J 不填，P 列不写)
    let mut new_rows_xml = String::with_capacity(voucher_rows.len() * 350);
    for (idx, row) in voucher_rows.iter().enumerate() {
        let r = idx + 4;
        new_rows_xml.push_str(&format!("<row r=\"{}\" spans=\"1:15\" customHeight=\"1\">", r));

        // A 列: *记账日期
        new_rows_xml.push_str(&format!("<c r=\"A{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.date)));
        // B 列: *凭证类型
        new_rows_xml.push_str(&format!("<c r=\"B{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.voucher_type)));
        // C 列: *凭证号
        new_rows_xml.push_str(&format!("<c r=\"C{}\"><v>{}</v></c>", r, escape_xml(&row.voucher_no)));
        // D 列: *摘要
        new_rows_xml.push_str(&format!("<c r=\"D{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.summary)));
        // E 列: *科目编码
        new_rows_xml.push_str(&format!("<c r=\"E{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.subject_code)));
        // F 列: *科目名称
        new_rows_xml.push_str(&format!("<c r=\"F{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.subject_name)));

        // G 列: *本币借方金额
        if let Some(amt) = row.debit_amount {
            new_rows_xml.push_str(&format!("<c r=\"G{}\"><v>{:.2}</v></c>", r, amt));
        }

        // H 列: *本币贷方金额
        if let Some(amt) = row.credit_amount {
            new_rows_xml.push_str(&format!("<c r=\"H{}\"><v>{:.2}</v></c>", r, amt));
        }

        // I 列和 J 列：不填 (留空)

        // K 列: 数量
        if let Some(qty) = row.qty {
            new_rows_xml.push_str(&format!("<c r=\"K{}\"><v>{}</v></c>", r, qty));
        }

        // L 列: 辅助编码
        if !row.aux_code.is_empty() {
            new_rows_xml.push_str(&format!("<c r=\"L{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.aux_code)));
        }

        // M 列: 辅助名称
        if !row.aux_name.is_empty() {
            new_rows_xml.push_str(&format!("<c r=\"M{}\" t=\"inlineStr\"><is><t>{}</t></is></c>", r, escape_xml(&row.aux_name)));
        }

        // N 列 (附件数), O 列 (制单人), P 列 (扩展列) 均不写入任何内容！
        new_rows_xml.push_str("</row>");
    }

    let new_sheet_xml = format!("{}{}{}", header_xml, new_rows_xml, tail_xml);

    // 3. 将原模板除凭证外的所有 sheet 及元数据原封不动写入输出文件 (100% 字节级保真)
    if let Some(parent) = output_path.parent() {
        std::fs::create_dir_all(parent).map_err(|e| e.to_string())?;
    }
    let out_file = std::fs::File::create(output_path).map_err(|e| format!("创建输出文件失败: {}", e))?;
    let mut zip_writer = zip::ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let mut entry = archive.by_index(i).map_err(|e| format!("读取 zip entry 失败: {}", e))?;
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

    zip_writer.finish().map_err(|e| format!("完成 zip 压缩失败: {}", e))?;
    Ok(())
}

fn export_external_audit_excel(output_path: &Path, records: &[ExternalAuditRecord]) -> Result<(), String> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();
    worksheet.set_name("外账账实库存核对明细").map_err(|e| e.to_string())?;

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

    worksheet.write_string_with_format(0, 0, "外账账实库存核对分析报告", &title_format).map_err(|e| e.to_string())?;

    let headers = [
        "序号", "外账存货编码", "药品标准名称", "规格型号", "生产厂家",
        "外账账面数量", "外账成本单价", "外账账面金额", "库管在库数量",
        "数量差异 (库管-外账)", "核对状态",
    ];

    for (c, h) in headers.iter().enumerate() {
        worksheet.write_string_with_format(2, c as u16, *h, &header_format).map_err(|e| e.to_string())?;
    }

    for (idx, r) in records.iter().enumerate() {
        let row_idx = (idx + 3) as u32;
        let is_diff = r.status == "数量差异";
        let base_fmt = if is_diff { &warn_row_format } else { &text_format };

        worksheet.write_number_with_format(row_idx, 0, (idx + 1) as f64, &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(row_idx, 1, &r.aux_code, &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(row_idx, 2, &r.name, base_fmt).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(row_idx, 3, &r.spec, base_fmt).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(row_idx, 4, &r.factory, base_fmt).map_err(|e| e.to_string())?;

        worksheet.write_number_with_format(row_idx, 5, r.ext_qty, &number_format).map_err(|e| e.to_string())?;
        worksheet.write_number_with_format(row_idx, 6, r.ext_price, &money_format).map_err(|e| e.to_string())?;
        worksheet.write_number_with_format(row_idx, 7, r.ext_amount, &money_format).map_err(|e| e.to_string())?;
        worksheet.write_number_with_format(row_idx, 8, r.wh_qty, &number_format).map_err(|e| e.to_string())?;
        worksheet.write_number_with_format(row_idx, 9, r.diff_qty, &number_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(row_idx, 10, &r.status, &center_format).map_err(|e| e.to_string())?;
    }

    worksheet.set_column_width(1, 14).map_err(|e| e.to_string())?;
    worksheet.set_column_width(2, 26).map_err(|e| e.to_string())?;
    worksheet.set_column_width(3, 16).map_err(|e| e.to_string())?;
    worksheet.set_column_width(4, 22).map_err(|e| e.to_string())?;
    worksheet.set_column_width(5, 13).map_err(|e| e.to_string())?;
    worksheet.set_column_width(7, 14).map_err(|e| e.to_string())?;
    worksheet.set_column_width(8, 13).map_err(|e| e.to_string())?;
    worksheet.set_column_width(9, 18).map_err(|e| e.to_string())?;
    worksheet.set_column_width(10, 12).map_err(|e| e.to_string())?;

    workbook.save(output_path).map_err(|e| format!("保存核对报告失败: {}", e))?;
    Ok(())
}
