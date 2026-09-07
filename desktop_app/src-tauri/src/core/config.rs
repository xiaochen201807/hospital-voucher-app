use chrono::{Datelike, Local, NaiveDate};
use regex::Regex;
use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};

#[derive(Debug, Clone, Serialize, Deserialize, Default)]
pub struct ConfigData {
    #[serde(default)]
    pub factory_abbreviations: HashMap<String, String>,
    #[serde(default)]
    pub vendor_abbr_map: HashMap<String, String>,
    #[serde(default)]
    pub strict_vendor_suffix_drugs: Vec<String>,
    #[serde(default)]
    pub drug_code_overrides: HashMap<String, String>,
}

/// 清洗文本：去除所有空格、中英文括号、斜杠、星号、下划线、横杠并转小写
pub fn clean_text(text: &str) -> String {
    text.chars()
        .filter(|c| !matches!(*c, ' ' | '\t' | '\r' | '\n' | '(' | ')' | '（' | '）' | '-' | '*' | '/' | '_' | '—' | '　'))
        .collect::<String>()
        .to_lowercase()
}

fn parse_date_value(value: &str) -> Option<NaiveDate> {
    let separators = Regex::new(
        r"(?P<year>\d{4})\s*(?:年|[-/.])\s*(?P<month>\d{1,2})\s*(?:月|[-/.])\s*(?P<day>\d{1,2})",
    )
    .unwrap();
    if let Some(caps) = separators.captures(value) {
        let year = caps.name("year")?.as_str().parse::<i32>().ok()?;
        let month = caps.name("month")?.as_str().parse::<u32>().ok()?;
        let day = caps.name("day")?.as_str().parse::<u32>().ok()?;
        return NaiveDate::from_ymd_opt(year, month, day);
    }

    let compact = Regex::new(r"(?P<year>\d{4})(?P<month>\d{2})(?P<day>\d{2})").unwrap();
    let caps = compact.captures(value)?;
    let year = caps.name("year")?.as_str().parse::<i32>().ok()?;
    let month = caps.name("month")?.as_str().parse::<u32>().ok()?;
    let day = caps.name("day")?.as_str().parse::<u32>().ok()?;
    NaiveDate::from_ymd_opt(year, month, day)
}

fn detect_voucher_date_from_file_names(file_paths: &[&str]) -> Option<String> {
    let re_month = Regex::new(r"(\d{4})[^\d]*?(\d{1,2})月").unwrap();
    let re_period = Regex::new(r"(\d{4})(\d{2})期?").unwrap();
    let re_dash = Regex::new(r"(\d{4})[-_](\d{1,2})").unwrap();

    for path in file_paths {
        let file_name = Path::new(path)
            .file_name()
            .and_then(|n| n.to_str())
            .unwrap_or("");

        if let Some(caps) = re_month.captures(file_name) {
            if let (Ok(y), Ok(m)) = (caps[1].parse::<i32>(), caps[2].parse::<u32>()) {
                if (1..=12).contains(&m) {
                    return Some(get_month_end_date_str(y, m));
                }
            }
        }

        if let Some(caps) = re_period.captures(file_name) {
            if let (Ok(y), Ok(m)) = (caps[1].parse::<i32>(), caps[2].parse::<u32>()) {
                if (1..=12).contains(&m) {
                    return Some(get_month_end_date_str(y, m));
                }
            }
        }

        if let Some(caps) = re_dash.captures(file_name) {
            if let (Ok(y), Ok(m)) = (caps[1].parse::<i32>(), caps[2].parse::<u32>()) {
                if (1..=12).contains(&m) {
                    return Some(get_month_end_date_str(y, m));
                }
            }
        }
    }

    None
}

/// 自动推断月末入账日期。
///
/// 优先级为：用户指定日期 > 来源表格中的最大业务日期 > 文件名中的期间 > 当前月月末。
pub fn detect_voucher_date_with_source_dates(
    custom_date: Option<&str>,
    file_paths: &[&str],
    source_dates: &[String],
) -> String {
    if let Some(d) = custom_date {
        let trimmed = d.trim();
        if !trimmed.is_empty() {
            return trimmed.to_string();
        }
    }

    if let Some(max_date) = source_dates.iter().filter_map(|value| parse_date_value(value)).max() {
        return get_month_end_date_str(max_date.year(), max_date.month());
    }

    if let Some(date) = detect_voucher_date_from_file_names(file_paths) {
        return date;
    }

    // 默认当月最后一天
    let today = Local::now().date_naive();
    get_month_end_date_str(today.year(), today.month())
}

/// 自动推断月末入账日期（兼容仅能从文件名推断的调用方）。
pub fn detect_voucher_date(custom_date: Option<&str>, file_paths: &[&str]) -> String {
    detect_voucher_date_with_source_dates(custom_date, file_paths, &[])
}

fn get_month_end_date_str(year: i32, month: u32) -> String {
    let (next_year, next_month) = if month == 12 {
        (year + 1, 1)
    } else {
        (year, month + 1)
    };
    if let Some(first_of_next) = NaiveDate::from_ymd_opt(next_year, next_month, 1) {
        if let Some(last_day) = first_of_next.pred_opt() {
            return last_day.format("%Y-%m-%d").to_string();
        }
    }
    format!("{:04}-{:02}-28", year, month)
}

/// 查找 factory_mapping.json 路径
pub fn find_config_path(custom: Option<&str>) -> PathBuf {
    if let Some(p) = custom {
        let pb = PathBuf::from(p);
        if pb.exists() {
            return pb;
        }
    }

    let mut candidates = vec![
        PathBuf::from("factory_mapping.json"),
        PathBuf::from("../factory_mapping.json"),
        PathBuf::from("../../factory_mapping.json"),
        PathBuf::from("resources/factory_mapping.json"),
    ];

    if let Ok(exe_p) = std::env::current_exe() {
        if let Some(dir) = exe_p.parent() {
            candidates.push(dir.join("factory_mapping.json"));
            candidates.push(dir.join("resources/factory_mapping.json"));
            candidates.push(dir.join("../Resources/resources/factory_mapping.json"));
            candidates.push(dir.join("../Resources/factory_mapping.json"));
        }
    }

    for c in &candidates {
        if c.exists() {
            return c.canonicalize().unwrap_or_else(|_| c.clone());
        }
    }

    PathBuf::from("factory_mapping.json")
}

/// 加载配置文件
pub fn load_config(custom: Option<&str>) -> (ConfigData, PathBuf) {
    let path = find_config_path(custom);
    load_config_from_path(&path)
}

/// 从指定路径加载配置。路径不存在或内容无效时返回默认配置，但仍保留该路径用于后续保存。
pub fn load_config_from_path(path: &Path) -> (ConfigData, PathBuf) {
    if path.exists() {
        if let Ok(content) = fs::read_to_string(path) {
            if let Ok(cfg) = serde_json::from_str::<ConfigData>(&content) {
                return (cfg, path.to_path_buf());
            }
        }
    }
    (ConfigData::default(), path.to_path_buf())
}

/// 保存配置文件
pub fn save_config_to_file(data: &ConfigData, custom: Option<&str>) -> Result<PathBuf, String> {
    let path = find_config_path(custom);
    save_config_to_path(data, &path)
}

/// 将配置保存到指定路径，并自动创建用户配置目录。
pub fn save_config_to_path(data: &ConfigData, path: &Path) -> Result<PathBuf, String> {
    let json_str = serde_json::to_string_pretty(data).map_err(|e| format!("序列化 JSON 失败: {}", e))?;
    if let Some(parent) = path.parent().filter(|p| !p.as_os_str().is_empty()) {
        fs::create_dir_all(parent).map_err(|e| format!("创建配置目录失败: {}", e))?;
    }
    fs::write(path, json_str).map_err(|e| format!("写入配置文件失败: {}", e))?;
    Ok(path.to_path_buf())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_dates_take_priority_over_ledger_filename_period() {
        let source_dates = vec!["2026/08/05".to_string(), "2026-08-12".to_string()];
        let result = detect_voucher_date_with_source_dates(
            None,
            &["西药-药品入库单.xlsx", "总账_20260907173619.xlsx"],
            &source_dates,
        );
        assert_eq!(result, "2026-08-31");
    }

    #[test]
    fn custom_date_wins_over_source_dates() {
        let source_dates = vec!["2026/08/12".to_string()];
        let result = detect_voucher_date_with_source_dates(
            Some("2026-09-15"),
            &["总账_20260907173619.xlsx"],
            &source_dates,
        );
        assert_eq!(result, "2026-09-15");
    }
}
