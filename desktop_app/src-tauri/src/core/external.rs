use crate::core::config::ConfigData;
use crate::core::excel_utils::{cell_as_f64, cell_as_string, open_excel};
use crate::core::voucher_writer::{copy_template_sheets, load_template_sheets};
use calamine::{Data, Reader};
use chrono::{Datelike, Local, NaiveDate};
use regex::Regex;
use rust_xlsxwriter::{Color, Format, FormatAlign, FormatBorder, Workbook};
use serde::{Deserialize, Serialize};
use std::collections::{HashMap, HashSet};
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
    m.insert("扬子江药业集团有限公司", "扬子江");
    m.insert("通化东宝药业股份有限公司", "通化东宝");
    m.insert("远大医药(中国)有限公司", "远大");
    m.insert("长春高新技术产业(集团)股份有限公司", "金赛");
    m
}

fn match_external_inventory_code(
    name: &str,
    spec: &str,
    factory: &str,
    items: &[ExternalInventoryItem],
    factory_map: &HashMap<&str, &str>,
) -> Option<(String, String)> {
    let norm_name = normalize_text(name);
    let clean_name = clean_drug_name(name);
    let norm_spec = normalize_text(spec);
    let norm_factory = normalize_text(factory);

    let abbr = factory_map.iter().find_map(|(full, b)| {
        if norm_factory.contains(*full) || norm_factory.contains(*b) {
            Some(*b)
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
    let factory_abbr = get_default_factory_abbr_map();

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
    let factory_abbr = get_default_factory_abbr_map();

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

    let mut raw_items = Vec::new();
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

        raw_items.push(RawSaleItem {
            row_no: r_idx + 1,
            name,
            spec,
            factory,
            qty,
            price,
            amount,
        });
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
    let factory_abbr = get_default_factory_abbr_map();

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

fn write_external_voucher_workbook(
    template_path: &Path,
    output_path: &Path,
    voucher_rows: &[ExternalVoucherRow],
) -> Result<(), String> {
    let sheets = load_template_sheets(template_path)?;
    let mut workbook = Workbook::new();

    copy_template_sheets(&mut workbook, &sheets)?;

    let worksheet = workbook.add_worksheet();
    worksheet.set_name("凭证").map_err(|e| e.to_string())?;

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

    let money_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_num_format("#,##0.00")
        .set_align(FormatAlign::Right)
        .set_border(FormatBorder::Thin);

    let qty_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_num_format("#,##0.##")
        .set_align(FormatAlign::Right)
        .set_border(FormatBorder::Thin);

    let warn_format = Format::new()
        .set_font_name(font_family)
        .set_font_size(10)
        .set_background_color(Color::RGB(0xFFFAD2))
        .set_border(FormatBorder::Thin);

    let original_voucher = sheets.iter().find(|s| s.name.contains("凭证"));
    if let Some(orig) = original_voucher {
        for (r_idx, row) in orig.rows.iter().enumerate().take(3) {
            for (c_idx, cell) in row.iter().enumerate() {
                let text = cell_as_string(cell);
                if r_idx == 2 {
                    worksheet
                        .write_string_with_format(r_idx as u32, c_idx as u16, &text, &header_format)
                        .map_err(|e| e.to_string())?;
                } else if r_idx == 0 {
                    worksheet
                        .write_string_with_format(r_idx as u32, c_idx as u16, &text, &title_format)
                        .map_err(|e| e.to_string())?;
                } else {
                    worksheet
                        .write_string(r_idx as u32, c_idx as u16, &text)
                        .map_err(|e| e.to_string())?;
                }
            }
        }
    } else {
        let headers = [
            "*记账日期", "*凭证类型", "*凭证号", "*摘要", "*科目编码", "*科目名称",
            "*本币借方金额", "*本币贷方金额", "汇率", "币别", "数量", "辅助编码",
            "辅助名称", "结算方式", "票号", "发生日期",
        ];
        for (c_idx, h) in headers.iter().enumerate() {
            worksheet
                .write_string_with_format(2, c_idx as u16, *h, &header_format)
                .map_err(|e| e.to_string())?;
        }
    }

    for (idx, row) in voucher_rows.iter().enumerate() {
        let r = (idx + 3) as u32;

        worksheet.write_string_with_format(r, 0, &row.date, &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 1, &row.voucher_type, &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 2, &row.voucher_no, &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 3, &row.summary, &text_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 4, &row.subject_code, &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 5, &row.subject_name, &text_format).map_err(|e| e.to_string())?;

        if let Some(amt) = row.debit_amount {
            worksheet.write_number_with_format(r, 6, amt, &money_format).map_err(|e| e.to_string())?;
        } else {
            worksheet.write_blank(r, 6, &text_format).map_err(|e| e.to_string())?;
        }

        if let Some(amt) = row.credit_amount {
            worksheet.write_number_with_format(r, 7, amt, &money_format).map_err(|e| e.to_string())?;
        } else {
            worksheet.write_blank(r, 7, &text_format).map_err(|e| e.to_string())?;
        }

        worksheet.write_number_with_format(r, 8, row.exchange_rate.unwrap_or(1.0), &center_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 9, &row.currency, &center_format).map_err(|e| e.to_string())?;

        if let Some(qty) = row.qty {
            worksheet.write_number_with_format(r, 10, qty, &qty_format).map_err(|e| e.to_string())?;
        } else {
            worksheet.write_blank(r, 10, &text_format).map_err(|e| e.to_string())?;
        }

        let code_fmt = if row.is_unmatched { &warn_format } else { &center_format };
        worksheet.write_string_with_format(r, 11, &row.aux_code, code_fmt).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 12, &row.aux_name, &text_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 13, &row.settle_method, &text_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 14, &row.bill_no, &text_format).map_err(|e| e.to_string())?;
        worksheet.write_string_with_format(r, 15, &row.occur_date, &center_format).map_err(|e| e.to_string())?;
    }

    worksheet.set_column_width(0, 13).map_err(|e| e.to_string())?;
    worksheet.set_column_width(1, 8).map_err(|e| e.to_string())?;
    worksheet.set_column_width(2, 8).map_err(|e| e.to_string())?;
    worksheet.set_column_width(3, 24).map_err(|e| e.to_string())?;
    worksheet.set_column_width(4, 12).map_err(|e| e.to_string())?;
    worksheet.set_column_width(5, 14).map_err(|e| e.to_string())?;
    worksheet.set_column_width(6, 14).map_err(|e| e.to_string())?;
    worksheet.set_column_width(7, 14).map_err(|e| e.to_string())?;
    worksheet.set_column_width(10, 10).map_err(|e| e.to_string())?;
    worksheet.set_column_width(11, 12).map_err(|e| e.to_string())?;
    worksheet.set_column_width(12, 28).map_err(|e| e.to_string())?;

    workbook.save(output_path).map_err(|e| format!("保存 Excel 失败: {}", e))?;
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
