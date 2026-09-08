pub mod core;

use serde::{Deserialize, Serialize};
use serde_json::{json, Value};
use std::fs;
use std::path::{Path, PathBuf};
use std::process::Command;
use tauri::Manager;

const CONFIG_FILE_NAME: &str = "factory_mapping.json";

#[derive(Debug, Serialize, Deserialize, Clone)]
struct ScannedItem {
    name: String,
    path: String,
}

#[derive(Debug, Serialize, Deserialize, Default)]
struct ScanResultData {
    sales_files: Vec<ScannedItem>,
    inbound_files: Vec<ScannedItem>,
    ledger_files: Vec<ScannedItem>,
    template_files: Vec<ScannedItem>,
    west_wh_files: Vec<ScannedItem>,
    tcm_wh_files: Vec<ScannedItem>,
    hc_wh_files: Vec<ScannedItem>,
    all_excel: Vec<ScannedItem>,
}

#[derive(Debug, Serialize, Deserialize)]
#[allow(dead_code)]
struct ScanResponse {
    success: bool,
    data: Option<ScanResultData>,
    error: Option<String>,
}

fn runtime_config_path(app: &tauri::AppHandle, custom: Option<&str>) -> Result<PathBuf, String> {
    if let Some(path) = custom.map(str::trim).filter(|path| !path.is_empty()) {
        return Ok(PathBuf::from(path));
    }

    app.path()
        .app_config_dir()
        .map(|dir| dir.join(CONFIG_FILE_NAME))
        .map_err(|e| format!("获取应用配置目录失败: {}", e))
}

fn load_runtime_config(
    app: &tauri::AppHandle,
    custom: Option<&str>,
) -> (core::config::ConfigData, PathBuf) {
    match runtime_config_path(app, custom) {
        Ok(path) if path.exists() || custom.is_some() => core::config::load_config_from_path(&path),
        Ok(path) => {
            // 首次运行时以安装包内的只读配置作为默认值，但保存位置仍固定为用户配置目录。
            let (cfg, _) = core::config::load_config(None);
            (cfg, path)
        }
        Err(_) => core::config::load_config(custom),
    }
}

fn serialize_command_result<T: Serialize>(result: Result<T, String>) -> Result<Value, String> {
    match result {
        Ok(value) => serde_json::to_value(value).map_err(|e| e.to_string()),
        Err(error) => Ok(json!({ "success": false, "error": error })),
    }
}

fn run_with_runtime_config<T, F>(
    app: &tauri::AppHandle,
    custom_config: Option<&str>,
    action: F,
) -> Result<Value, String>
where
    T: Serialize,
    F: FnOnce(&core::config::ConfigData) -> Result<T, String>,
{
    let (config, _) = load_runtime_config(app, custom_config);
    serialize_command_result(action(&config))
}

fn save_runtime_config(
    app: &tauri::AppHandle,
    data: &core::config::ConfigData,
    custom: Option<&str>,
) -> Result<PathBuf, String> {
    let path = runtime_config_path(app, custom)?;
    core::config::save_config_to_path(data, &path)
}

#[tauri::command]
fn scan_files(dir: Option<String>) -> Result<Value, String> {
    let scan_path = if let Some(d) = dir {
        PathBuf::from(d)
    } else {
        let cur = std::env::current_dir().unwrap_or_else(|_| PathBuf::from("."));
        // 检查当前目录下是否有 xls/xlsx 文件
        let has_excel = fs::read_dir(&cur)
            .map(|entries| {
                entries.filter_map(|e| e.ok()).any(|e| {
                    let n = e.file_name().to_string_lossy().to_string();
                    (n.ends_with(".xlsx") || n.ends_with(".xls")) && !n.starts_with("~$")
                })
            })
            .unwrap_or(false);

        if has_excel {
            cur
        } else {
            // 尝试用户主目录下的 Downloads 或 Desktop 目录
            let home_opt = std::env::var("HOME")
                .or_else(|_| std::env::var("USERPROFILE"))
                .ok();
            let mut fallback = cur.clone();
            if let Some(h) = home_opt {
                let dl = PathBuf::from(&h).join("Downloads");
                let dt = PathBuf::from(&h).join("Desktop");
                if dl.exists() {
                    fallback = dl;
                } else if dt.exists() {
                    fallback = dt;
                }
            }
            fallback
        }
    };

    if !scan_path.exists() {
        return Ok(json!({
            "success": false,
            "error": format!("目录 '{:?}' 不存在", scan_path)
        }));
    }

    let mut res = ScanResultData::default();

    if let Ok(entries) = fs::read_dir(&scan_path) {
        let mut files: Vec<PathBuf> = entries
            .filter_map(|e| e.ok().map(|ent| ent.path()))
            .filter(|p| {
                if let Some(name) = p.file_name().and_then(|n| n.to_str()) {
                    !name.starts_with("~$")
                        && (name.ends_with(".xlsx") || name.ends_with(".xls"))
                        && !name.contains("已生成")
                        && !name.contains("backup")
                } else {
                    false
                }
            })
            .collect();

        files.sort();

        for p in files {
            let name = p
                .file_name()
                .and_then(|n| n.to_str())
                .unwrap_or("")
                .to_string();
            let abs_path = p
                .canonicalize()
                .unwrap_or(p.clone())
                .to_string_lossy()
                .to_string();
            let item = ScannedItem {
                name: name.clone(),
                path: abs_path,
            };

            res.all_excel.push(item.clone());

            if name.contains("销售") {
                res.sales_files.push(item.clone());
            }
            if name.contains("入库") && !name.contains("模板") {
                res.inbound_files.push(item.clone());
            }
            if name.contains("总账") || name.contains("数量金额") {
                res.ledger_files.push(item.clone());
            }
            if name.contains("模板") {
                res.template_files.push(item.clone());
            }
            if name.contains("西药")
                && (name.contains("库存") || name.contains("报表") || name.contains("房"))
            {
                res.west_wh_files.push(item.clone());
            } else if name.contains("中药")
                && (name.contains("库存") || name.contains("报表") || name.contains("房"))
            {
                res.tcm_wh_files.push(item.clone());
            } else if (name.contains("耗材") || name.contains("材料"))
                && (name.contains("库存") || name.contains("报表") || name.contains("库"))
            {
                res.hc_wh_files.push(item.clone());
            }
        }
    }

    Ok(json!({
        "success": true,
        "data": res
    }))
}

#[tauri::command(rename_all = "camelCase")]
fn execute_sales_process(
    file: String,
    output: Option<String>,
    sheet_name: Option<String>,
) -> Result<Value, String> {
    serialize_command_result(core::sales::process_sales_file(
        &file,
        output.as_deref(),
        sheet_name.as_deref(),
    ))
}

#[tauri::command(rename_all = "camelCase")]
fn search_ledger_candidates(
    ledger: String,
    query: String,
    limit: Option<usize>,
    category: Option<String>,
) -> Result<Value, String> {
    let entries = core::ledger::load_ledger_entries(Path::new(&ledger))?;
    let candidates = core::matching::search_ledger_candidates_for_category(
        &query,
        &entries,
        category.as_deref(),
        limit.unwrap_or(50),
    );
    Ok(json!({
        "success": true,
        "candidates": candidates
    }))
}

#[tauri::command(rename_all = "camelCase")]
#[allow(clippy::too_many_arguments)]
fn execute_outbound_voucher(
    app: tauri::AppHandle,
    sales: String,
    ledger: String,
    template: String,
    output: Option<String>,
    date: Option<String>,
    fallback_price: Option<bool>,
    confirmed_items: Option<Vec<core::models::ConfirmedLedgerMapping>>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::outbound::generate_outbound_voucher(
            &sales,
            &ledger,
            &template,
            output.as_deref(),
            date.as_deref(),
            fallback_price.unwrap_or(true),
            cfg,
            confirmed_items.as_deref(),
        )
    })
}

#[tauri::command(rename_all = "camelCase")]
#[allow(clippy::too_many_arguments)]
fn execute_inbound_voucher(
    app: tauri::AppHandle,
    inbound: String,
    ledger: String,
    template: String,
    output: Option<String>,
    date: Option<String>,
    voucher_no: Option<String>,
    confirmed_items: Option<Vec<core::models::ConfirmedLedgerMapping>>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::inbound::generate_inbound_voucher(
            &inbound,
            &ledger,
            &template,
            output.as_deref(),
            date.as_deref(),
            voucher_no.as_deref(),
            cfg,
            confirmed_items.as_deref(),
        )
    })
}

#[tauri::command]
fn preview_inventory_audit_mapping(
    app: tauri::AppHandle,
    ledger: String,
    west: Option<String>,
    tcm: Option<String>,
    hc: Option<String>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::audit::preview_inventory_audit_mapping(
            &ledger,
            west.as_deref(),
            tcm.as_deref(),
            hc.as_deref(),
            cfg,
        )
    })
}

#[tauri::command(rename_all = "camelCase")]
fn execute_inventory_audit_with_mapping(
    app: tauri::AppHandle,
    ledger: String,
    confirmed_items: Vec<core::audit::ConfirmedAuditMappingItem>,
    output: Option<String>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::audit::execute_inventory_audit_with_mapping(
            &ledger,
            confirmed_items,
            output.as_deref(),
            cfg,
        )
    })
}

#[tauri::command]
fn execute_inventory_audit(
    app: tauri::AppHandle,
    ledger: String,
    west: Option<String>,
    tcm: Option<String>,
    hc: Option<String>,
    config: Option<String>,
    output: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::audit::run_inventory_audit(
            &ledger,
            west.as_deref(),
            tcm.as_deref(),
            hc.as_deref(),
            output.as_deref(),
            cfg,
        )
    })
}

#[tauri::command]
fn get_config(app: tauri::AppHandle, config: Option<String>) -> Result<Value, String> {
    let (cfg, path) = load_runtime_config(&app, config.as_deref());
    Ok(json!({
        "success": true,
        "data": cfg,
        "config_file": path.to_string_lossy().to_string()
    }))
}

#[tauri::command]
fn save_config(
    app: tauri::AppHandle,
    data: String,
    config: Option<String>,
) -> Result<Value, String> {
    let cfg_data: core::config::ConfigData = match serde_json::from_str(&data) {
        Ok(d) => d,
        Err(e) => {
            return Ok(
                json!({ "success": false, "error": format!("解析配置 JSON 格式失败: {}", e) }),
            )
        }
    };

    match save_runtime_config(&app, &cfg_data, config.as_deref()) {
        Ok(p) => Ok(json!({
            "success": true,
            "message": "配置保存成功",
            "config_file": p.to_string_lossy().to_string()
        })),
        Err(e) => Ok(json!({ "success": false, "error": e })),
    }
}

#[tauri::command]
fn open_in_system(path: String) -> Result<bool, String> {
    let target = Path::new(&path);
    if !target.exists() {
        return Err(format!("文件或路径不存在: {}", path));
    }

    #[cfg(target_os = "macos")]
    {
        Command::new("open")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "windows")]
    {
        Command::new("explorer")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "linux")]
    {
        Command::new("xdg-open")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }

    Ok(true)
}

#[tauri::command]
fn show_in_folder(path: String) -> Result<bool, String> {
    let target = Path::new(&path);
    if !target.exists() {
        return Err(format!("文件不存在: {}", path));
    }

    #[cfg(target_os = "macos")]
    {
        Command::new("open")
            .arg("-R")
            .arg(&path)
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "windows")]
    {
        Command::new("explorer")
            .arg(format!("/select,{}", path))
            .spawn()
            .map_err(|e| e.to_string())?;
    }
    #[cfg(target_os = "linux")]
    {
        if let Some(parent) = target.parent() {
            Command::new("xdg-open")
                .arg(parent)
                .spawn()
                .map_err(|e| e.to_string())?;
        }
    }

    Ok(true)
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_opener::init())
        .invoke_handler(tauri::generate_handler![
            scan_files,
            execute_sales_process,
            search_ledger_candidates,
            execute_outbound_voucher,
            execute_inbound_voucher,
            execute_inventory_audit,
            preview_inventory_audit_mapping,
            execute_inventory_audit_with_mapping,
            get_config,
            save_config,
            open_in_system,
            show_in_folder
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

#[cfg(test)]
mod tests {
    use super::*;

    fn project_file(name: &str) -> String {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(name)
            .to_string_lossy()
            .to_string()
    }

    #[test]
    fn test_scan_files() {
        let res = scan_files(None).unwrap();
        assert!(res["success"].as_bool().unwrap_or(false));
    }

    #[test]
    fn test_sales_and_audit() {
        let (cfg, _) = core::config::load_config(None);
        let sales_path = project_file("2026.8月西药销售表.xls");
        let ledger_path = project_file("石家庄心理医院_数量金额总账_20260903150759.xlsx");
        let west_path = project_file("石家庄心理医院新西药房库存汇总报表2026831.xls");
        let sales_res =
            core::sales::process_sales_file(&sales_path, None, None).expect("销售汇总处理必须成功");
        assert_eq!(sales_res.totals.unique_count, 54);
        assert_eq!(sales_res.totals.original_count, 82);
        assert_eq!(sales_res.totals.total_qty as i64, 139353);
        println!(
            ">>> 纯 Rust 销售汇总验证成功: {} 种去重药品 (原 {} 笔), 总件数: {}",
            sales_res.totals.unique_count,
            sales_res.totals.original_count,
            sales_res.totals.total_qty
        );

        let audit_res = core::audit::run_inventory_audit(
            &ledger_path,
            Some(&west_path),
            None,
            None,
            None,
            &cfg,
        )
        .expect("账实核对处理必须成功");

        println!(
            ">>> 纯 Rust 账实核对成功: 总品规 {}, 吻合 {}, 差异 {}, 吻合率 {}%",
            audit_res.overall.total_items,
            audit_res.overall.equal_count,
            audit_res.overall.diff_count,
            audit_res.overall.match_rate
        );
        assert!(audit_res.overall.total_items > 0);
        assert_eq!(
            audit_res.overall.total_items - audit_res.overall.equal_count,
            audit_res.overall.diff_count
                + audit_res.overall.wh_only_count
                + audit_res.overall.ledger_only_count
        );
        assert!((0.0..=100.0).contains(&audit_res.overall.match_rate));
        let historical_balance = audit_res
            .categories
            .iter()
            .flat_map(|category| category.records.iter())
            .find(|record| record.ledger_code == "1201_XY0014")
            .expect("审计结果应保留历史数量为 0 但金额非 0 的总账记录");
        assert_eq!(historical_balance.ledger_amt, 5.21);
    }
}
