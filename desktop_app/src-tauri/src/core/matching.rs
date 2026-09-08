use crate::core::config::{clean_text, ConfigData};
use crate::core::models::LedgerEntry;

/// 常用中药炮制前缀列表。
const TCM_PREFIXES: &[&str] = &[
    "制", "炒", "麸炒", "炙", "煅", "酒", "醋", "生", "清", "法", "姜", "焦", "蜜炙", "盐", "熟",
];

/// 剥离中药炮制前缀，获取核心本草名称。
pub fn strip_tcm_prefix(name: &str) -> &str {
    for prefix in TCM_PREFIXES {
        if let Some(stripped) = name.strip_prefix(prefix) {
            if !stripped.is_empty() {
                return stripped;
            }
        }
    }
    name
}

/// 从药名中解析基础品名与括号内的厂家标签。
pub fn extract_name_and_vendor_tag(name: &str) -> (String, Option<String>) {
    if let Some((start_idx, open_char)) = name.char_indices().find(|(_, c)| *c == '（' || *c == '(')
    {
        let after_open = &name[start_idx + open_char.len_utf8()..];
        if let Some(end_rel_idx) = after_open.find(['）', ')']) {
            let base = name[..start_idx].trim().to_string();
            let tag = after_open[..end_rel_idx].trim().to_string();
            if !tag.is_empty() {
                return (base, Some(tag));
            }
        }
    }
    (name.to_string(), None)
}

/// 匹配药品存货科目，并返回匹配方式说明。
pub fn match_drug_with_method(
    drug_name: &str,
    spec: &str,
    factory: &str,
    ledger: &[LedgerEntry],
    config: &ConfigData,
) -> Option<(LedgerEntry, String)> {
    let clean_drug = clean_text(drug_name);
    let clean_sp = clean_text(spec);
    let clean_fac = clean_text(factory);

    // 1. 提取厂家简称（三重智能容错）。
    let mut vendor_suffix = String::new();
    if !clean_fac.is_empty() {
        for (full, brief) in &config.factory_abbreviations {
            let clean_full = clean_text(full);
            let clean_brief = clean_text(brief);
            if clean_fac.contains(&clean_full)
                || (!clean_brief.is_empty() && clean_fac.contains(&clean_brief))
                || (!clean_fac.is_empty()
                    && clean_full.contains(&clean_fac)
                    && clean_fac.len() >= 6)
            {
                vendor_suffix = brief.clone();
                break;
            }
        }
    }

    // 0. 特殊手动指定覆盖。
    let mut override_cands = Vec::new();
    if !vendor_suffix.is_empty() {
        override_cands.push(clean_text(&format!("{}({})", drug_name, vendor_suffix)));
        override_cands.push(clean_text(&format!("{}（{}）", drug_name, vendor_suffix)));
        override_cands.push(clean_text(&format!("{}{}", drug_name, vendor_suffix)));
    }
    if !factory.is_empty() {
        override_cands.push(clean_text(&format!("{}({})", drug_name, factory)));
        override_cands.push(clean_text(&format!("{}（{}）", drug_name, factory)));
    }
    if !spec.is_empty() {
        override_cands.push(clean_text(&format!("{}{}", drug_name, spec)));
    }
    override_cands.push(clean_drug.clone());

    for (key, value) in &config.drug_code_overrides {
        let clean_key = clean_text(key);
        if override_cands
            .iter()
            .any(|candidate| candidate == &clean_key)
        {
            let trimmed_value = value.trim();
            // 空配置代表明确阻断自动匹配，避免穿透到模糊匹配。
            if trimmed_value.is_empty() {
                return None;
            }
            if let Some(target) = ledger
                .iter()
                .find(|entry| entry.code == trimmed_value || entry.aux_code == trimmed_value)
            {
                return Some((target.clone(), "手动配置指定".to_string()));
            }
            return None;
        }
    }

    // 严格厂家后缀保护。
    let is_strict_drug = config
        .strict_vendor_suffix_drugs
        .iter()
        .any(|drug| clean_text(drug) == clean_drug);

    // 优先级 1：全名 + 厂家全词精准匹配。
    if !vendor_suffix.is_empty() {
        let expected_with_vendor = format!("{}{}", clean_drug, clean_text(&vendor_suffix));
        for entry in ledger {
            let entry_name = clean_text(&entry.drug_name);
            if entry_name == expected_with_vendor {
                return Some((
                    entry.clone(),
                    format!("厂家简称精准匹配 ({})", vendor_suffix),
                ));
            }
        }
    }

    // 优先级 1.5：财务总账品名括号厂家反向解析。
    if !clean_fac.is_empty() {
        for entry in ledger {
            let (ledger_base, ledger_tag) = extract_name_and_vendor_tag(&entry.drug_name);
            if let Some(ledger_tag) = ledger_tag {
                let clean_ledger_base = clean_text(&ledger_base);
                let clean_ledger_tag = clean_text(&ledger_tag);
                if clean_ledger_base == clean_drug && clean_fac.contains(&clean_ledger_tag) {
                    return Some((
                        entry.clone(),
                        format!("总账括号厂家反向匹配 ({})", ledger_tag),
                    ));
                }
            }
        }
    }

    // 严格厂家药品不能回退到同名通用科目。
    if is_strict_drug {
        return None;
    }

    // 优先级 2：全名 + 规格精准一致。
    for entry in ledger {
        let entry_name = clean_text(&entry.drug_name);
        let entry_spec = clean_text(&entry.spec);
        if entry_name == clean_drug && !clean_sp.is_empty() && entry_spec == clean_sp {
            return Some((entry.clone(), "品名与规格完全一致".to_string()));
        }
    }

    // 优先级 3：仅全名一致；多条候选时必须通过规格明确排他。
    let candidates: Vec<&LedgerEntry> = ledger
        .iter()
        .filter(|entry| clean_text(&entry.drug_name) == clean_drug)
        .collect();

    if candidates.len() == 1 {
        return Some((candidates[0].clone(), "同名单品规自动关联".to_string()));
    } else if candidates.len() > 1 {
        let mut matched_candidate = None;
        let mut match_count = 0;
        for candidate in &candidates {
            let candidate_spec = clean_text(&candidate.spec);
            if !candidate_spec.is_empty()
                && !clean_sp.is_empty()
                && (clean_sp == candidate_spec
                    || clean_sp.contains(&candidate_spec)
                    || candidate_spec.contains(&clean_sp))
            {
                matched_candidate = Some((*candidate).clone());
                match_count += 1;
            }
        }
        if match_count == 1 {
            return matched_candidate.map(|candidate| (candidate, "同名规格排他命中".to_string()));
        }
    }

    // 优先级 4：中药炮制前缀兼容匹配。
    let stripped_clean_drug = clean_text(strip_tcm_prefix(&clean_drug));
    let mut tcm_candidates = Vec::new();
    for entry in ledger {
        let (entry_base, entry_tag) = extract_name_and_vendor_tag(&entry.drug_name);
        let clean_entry_base = clean_text(&entry_base);
        let stripped_clean_entry = clean_text(strip_tcm_prefix(&clean_entry_base));

        if stripped_clean_drug == stripped_clean_entry
            || clean_drug == stripped_clean_entry
            || stripped_clean_drug == clean_entry_base
        {
            let mut vendor_ok = true;
            if let Some(ledger_tag) = entry_tag {
                let clean_tag = clean_text(&ledger_tag);
                if !clean_fac.is_empty() && !clean_fac.contains(&clean_tag) {
                    vendor_ok = false;
                }
            }

            if vendor_ok {
                tcm_candidates.push(entry);
            }
        }
    }

    if tcm_candidates.len() == 1 {
        return Some((
            tcm_candidates[0].clone(),
            format!("中药炮制前缀兼容 ({})", tcm_candidates[0].drug_name),
        ));
    } else if tcm_candidates.len() > 1 {
        let mut spec_matched = Vec::new();
        for candidate in &tcm_candidates {
            let candidate_spec = clean_text(&candidate.spec);
            if !candidate_spec.is_empty()
                && !clean_sp.is_empty()
                && (clean_sp == candidate_spec
                    || clean_sp.contains(&candidate_spec)
                    || candidate_spec.contains(&clean_sp))
            {
                spec_matched.push(*candidate);
            }
        }
        if spec_matched.len() == 1 {
            return Some((
                spec_matched[0].clone(),
                format!("中药炮制规格排他命中 ({})", spec_matched[0].drug_name),
            ));
        }
    }

    None
}

pub fn match_drug(
    drug_name: &str,
    spec: &str,
    factory: &str,
    ledger: &[LedgerEntry],
    config: &ConfigData,
) -> Option<LedgerEntry> {
    match_drug_with_method(drug_name, spec, factory, ledger, config).map(|(entry, _)| entry)
}
