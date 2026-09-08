use crate::core::config::clean_text;
use crate::core::excel_utils::cell_as_string;
use crate::core::voucher_writer::TemplateSheet;
use std::collections::HashMap;

/// 从供应商全称生成凭证摘要，例如“河北蕴德药业有限公司” -> “河北蕴德到货”。
pub fn get_supplier_brief(sup_name: &str) -> String {
    if sup_name.is_empty() {
        return "药品到货".to_string();
    }

    let cleaned = sup_name
        .replace("股份有限公司", "")
        .replace("医药有限公司", "")
        .replace("药业有限公司", "")
        .replace("有限责任公司", "")
        .replace("有限公司", "")
        .replace("责任公司", "");
    format!("{}到货", cleaned.trim())
}

/// 从已读取的凭证模板附表中提取供应商字典。
pub fn load_supplier_dict(sheets: &[TemplateSheet]) -> HashMap<String, String> {
    let mut dict = HashMap::new();

    if let Some(sheet) = sheets.iter().find(|sheet| sheet.name.contains("辅助核算")) {
        for row in &sheet.rows {
            if row.len() < 3 {
                continue;
            }
            let category = cell_as_string(&row[0]);
            if !category.contains("供应商") {
                continue;
            }

            let code = cell_as_string(&row[1]);
            let name = cell_as_string(&row[2]);
            if !code.is_empty() && !name.is_empty() {
                dict.insert(code, name);
            }
        }
    }

    // 与历史脚本保持兼容；正式模板有供应商表时不会走此兜底。
    if dict.is_empty() {
        dict.insert("001".into(), "国药乐仁堂医药有限公司".into());
        dict.insert("002".into(), "华润益生制药有限公司".into());
        dict.insert("004".into(), "河北国泰医药有限责任公司".into());
        dict.insert("005".into(), "河北蕴德药业有限公司".into());
        dict.insert("007".into(), "石药集团中诚医药字号".into());
    }

    dict
}

/// 匹配供应商编码，并按匹配强度确定性地选择结果。
pub fn match_supplier_code(sup_name: &str, dict: &HashMap<String, String>) -> (String, String) {
    let clean_sup = clean_text(sup_name);
    if clean_sup.is_empty() {
        return (String::new(), sup_name.to_string());
    }

    let mut candidates: Vec<(u8, usize, String, String)> = Vec::new();
    let core_sup = clean_text(&strip_company_suffix(sup_name));

    for (code, name) in dict {
        let clean_name = clean_text(name);
        let score = if clean_name == clean_sup {
            Some(3)
        } else if clean_name.contains(&clean_sup) || clean_sup.contains(&clean_name) {
            Some(2)
        } else if !core_sup.is_empty() && clean_name.contains(&core_sup) {
            Some(1)
        } else {
            None
        };

        if let Some(score) = score {
            candidates.push((score, clean_name.len(), code.clone(), name.clone()));
        }
    }

    candidates.sort_by(|a, b| {
        b.0.cmp(&a.0)
            .then_with(|| b.1.cmp(&a.1))
            .then_with(|| a.2.cmp(&b.2))
    });

    candidates
        .into_iter()
        .next()
        .map(|(_, _, code, name)| (code, name))
        .unwrap_or_else(|| (String::new(), sup_name.to_string()))
}

fn strip_company_suffix(name: &str) -> String {
    name.replace("股份有限公司", "")
        .replace("医药有限公司", "")
        .replace("药业有限公司", "")
        .replace("有限责任公司", "")
        .replace("有限公司", "")
        .replace("责任公司", "")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn supplier_matching_is_deterministic_for_overlapping_names() {
        let dict = HashMap::from([
            ("001".to_string(), "河北国泰医药有限公司".to_string()),
            ("002".to_string(), "河北国泰".to_string()),
        ]);

        assert_eq!(
            match_supplier_code("河北国泰", &dict).0,
            "002",
            "短名称完整匹配应优先于较长名称的包含匹配"
        );
    }
}
