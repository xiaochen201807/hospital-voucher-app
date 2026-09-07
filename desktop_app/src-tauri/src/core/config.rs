use serde::{Deserialize, Serialize};
use std::collections::HashMap;
use std::fs;
use std::path::PathBuf;

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

/// 查找 factory_mapping.json 路径
pub fn find_config_path(custom: Option<&str>) -> PathBuf {
    if let Some(p) = custom {
        let pb = PathBuf::from(p);
        if pb.exists() {
            return pb;
        }
    }

    let candidates = [
        PathBuf::from("factory_mapping.json"),
        PathBuf::from("../factory_mapping.json"),
        PathBuf::from("../../factory_mapping.json"),
        PathBuf::from("resources/factory_mapping.json"),
    ];

    for c in &candidates {
        if c.exists() {
            return c.canonicalize().unwrap_or(c.clone());
        }
    }

    PathBuf::from("factory_mapping.json")
}

/// 加载配置文件
pub fn load_config(custom: Option<&str>) -> (ConfigData, PathBuf) {
    let path = find_config_path(custom);
    if path.exists() {
        if let Ok(content) = fs::read_to_string(&path) {
            if let Ok(cfg) = serde_json::from_str::<ConfigData>(&content) {
                return (cfg, path);
            }
        }
    }
    (ConfigData::default(), path)
}

/// 保存配置文件
pub fn save_config_to_file(data: &ConfigData, custom: Option<&str>) -> Result<PathBuf, String> {
    let path = find_config_path(custom);
    let json_str = serde_json::to_string_pretty(data).map_err(|e| format!("序列化 JSON 失败: {}", e))?;
    fs::write(&path, json_str).map_err(|e| format!("写入配置文件失败: {}", e))?;
    Ok(path)
}
