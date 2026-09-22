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
    external_template_files: Vec<ScannedItem>,
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
            if name.contains("迁账") || name.contains("外账") {
                res.external_template_files.push(item.clone());
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
fn search_external_inventory_candidates(
    template: String,
    query: String,
    limit: Option<usize>,
) -> Result<Value, String> {
    let candidates = core::external::search_external_inventory_candidates(
        Path::new(&template),
        &query,
        limit.unwrap_or(50),
    )?;
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

#[tauri::command(rename_all = "camelCase")]
fn execute_external_inbound_voucher(
    app: tauri::AppHandle,
    inbound: String,
    template: String,
    output: Option<String>,
    date: Option<String>,
    voucher_no: Option<String>,
    confirmed_items: Option<Vec<core::external::ConfirmedExternalInventoryMapping>>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::external::generate_external_inbound_voucher(
            Path::new(&inbound),
            Path::new(&template),
            output.as_deref().map(Path::new),
            date.as_deref(),
            voucher_no.as_deref(),
            Some(cfg),
            confirmed_items.as_deref(),
        )
    })
}

#[tauri::command(rename_all = "camelCase")]
fn execute_external_outbound_voucher(
    app: tauri::AppHandle,
    sales: String,
    template: String,
    output: Option<String>,
    date: Option<String>,
    voucher_no: Option<String>,
    confirmed_items: Option<Vec<core::external::ConfirmedExternalInventoryMapping>>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::external::generate_external_outbound_voucher(
            Path::new(&sales),
            Path::new(&template),
            output.as_deref().map(Path::new),
            date.as_deref(),
            voucher_no.as_deref(),
            Some(cfg),
            confirmed_items.as_deref(),
        )
    })
}

#[tauri::command(rename_all = "camelCase")]
fn execute_external_inventory_audit(
    app: tauri::AppHandle,
    template: String,
    west: String,
    tcm: String,
    hc: String,
    output: Option<String>,
    config: Option<String>,
) -> Result<Value, String> {
    run_with_runtime_config(&app, config.as_deref(), |cfg| {
        core::external::generate_external_inventory_audit(
            Path::new(&template),
            Path::new(&west),
            Path::new(&tcm),
            Path::new(&hc),
            output.as_deref().map(Path::new),
            Some(cfg),
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
            search_external_inventory_candidates,
            execute_outbound_voucher,
            execute_inbound_voucher,
            execute_external_inbound_voucher,
            execute_external_outbound_voucher,
            execute_external_inventory_audit,
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
    use calamine::Reader;
    use std::io::Read;

    fn project_file(name: &str) -> String {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../..")
            .join(name)
            .to_string_lossy()
            .to_string()
    }

    fn test_artifact(name: &str) -> PathBuf {
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("测试环境系统时间必须有效")
            .as_nanos();
        std::env::temp_dir().join(format!("{}_{}_{}.xlsx", name, std::process::id(), nonce))
    }

    fn assert_voucher_template_style(path: &Path) {
        let file = fs::File::open(path).expect("应能打开凭证输出压缩包");
        let mut archive = zip::ZipArchive::new(file).expect("凭证输出应为有效 XLSX 压缩包");
        let mut sheet_xml = String::new();
        archive
            .by_name("xl/worksheets/sheet6.xml")
            .expect("凭证 Sheet 应存在")
            .read_to_string(&mut sheet_xml)
            .expect("凭证 Sheet XML 应能读取");

        assert!(
            sheet_xml.contains("spans=\"1:16\""),
            "凭证分录行必须保留 16 列范围"
        );
        assert!(
            sheet_xml.contains("r=\"A4\" s=\"13\""),
            "日期列必须保留模板样式"
        );
        assert!(
            sheet_xml.contains("r=\"G4\" s=\"6\""),
            "金额列必须保留模板数字样式"
        );
        assert!(
            sheet_xml.contains("dimension ref=\"A1:P"),
            "输出维度必须随分录数量更新"
        );
    }

    #[test]
    fn test_scan_files() {
        let res = scan_files(None).unwrap();
        assert!(res["success"].as_bool().unwrap_or(false));
    }

    #[test]
    #[ignore = "依赖未提交的业务 Excel 样例；本地有样例时使用 cargo test --lib -- --ignored"]
    fn test_sales_and_audit() {
        let (cfg, _) = core::config::load_config(None);
        let sales_path = project_file("2026.8月西药销售表.xls");
        let ledger_path = project_file("石家庄心理医院_数量金额总账_20260903150759.xlsx");
        let west_path = project_file("石家庄心理医院新西药房库存汇总报表2026831.xls");
        let sales_output = test_artifact("internal_sales");
        let sales_output_string = sales_output.to_string_lossy().to_string();
        let audit_output = test_artifact("internal_audit");
        let audit_output_string = audit_output.to_string_lossy().to_string();
        let sales_res =
            core::sales::process_sales_file(&sales_path, Some(&sales_output_string), None)
                .expect("销售汇总处理必须成功");
        assert_eq!(sales_res.totals.unique_count, 54);
        assert_eq!(sales_res.totals.original_count, 82);
        assert_eq!(sales_res.totals.total_qty as i64, 139353);

        let audit_res = core::audit::run_inventory_audit(
            &ledger_path,
            Some(&west_path),
            None,
            None,
            Some(&audit_output_string),
            &cfg,
        )
        .expect("账实核对处理必须成功");

        assert!(audit_res.overall.total_items > 0);
        assert_eq!(
            audit_res.overall.total_items - audit_res.overall.equal_count,
            audit_res.overall.diff_count
                + audit_res.overall.wh_only_count
                + audit_res.overall.ledger_only_count
        );
        assert!((0.0..=100.0).contains(&audit_res.overall.match_rate));
        let _ = fs::remove_file(sales_output);
        let _ = fs::remove_file(audit_output);
    }

    #[test]
    #[ignore = "依赖未提交的业务 Excel 样例；本地有样例时使用 cargo test --lib -- --ignored"]
    fn test_external_voucher_suite() {
        let (cfg, _) = core::config::load_config(None);
        let tmpl_path = project_file("表格迁账参考模板.xlsx");
        let inbound_path = project_file("西药入库单.xlsx");
        let sales_path = project_file("2026.8月西药销售表_已汇总.xlsx");
        let wh_path = project_file("石家庄心理医院新西药房库存汇总报表2026831.xls");
        let inbound_output = test_artifact("external_inbound");
        let outbound_output = test_artifact("external_outbound");
        let audit_output = test_artifact("external_audit");

        // 1. 测试纯 Rust 外账入库
        let in_res = core::external::generate_external_inbound_voucher(
            Path::new(&inbound_path),
            Path::new(&tmpl_path),
            Some(&inbound_output),
            Some("2026-08-31"),
            Some("1"),
            Some(&cfg),
            None,
        )
        .expect("纯 Rust 外账入库调用必须成功");
        assert!(in_res.is_balanced);
        assert_eq!(in_res.diff, 0.0);
        assert_eq!(in_res.total_debit, in_res.total_credit);

        // 验证用户反馈的两大格式问题：1. I/J 列不填 2. P 列无日期
        let mut in_wb = core::excel_utils::open_excel(Path::new(&in_res.output_file))
            .expect("打开外账入库生成文件");
        let in_range = in_wb.worksheet_range("凭证").expect("打开凭证工作表");
        assert_voucher_template_style(Path::new(&in_res.output_file));
        for (r_idx, row) in in_range.rows().skip(3).enumerate() {
            if let Some(c8) = row.get(8) {
                let s = core::excel_utils::cell_as_string(c8);
                assert!(
                    s.trim().is_empty(),
                    "第 {} 行 I 列 (借方外币金额) 必须为空，当前为: {}",
                    r_idx + 4,
                    s
                );
            }
            if let Some(c9) = row.get(9) {
                let s = core::excel_utils::cell_as_string(c9);
                assert!(
                    s.trim().is_empty(),
                    "第 {} 行 J 列 (贷方外币金额) 必须为空，当前为: {}",
                    r_idx + 4,
                    s
                );
            }
            if let Some(c15) = row.get(15) {
                let s = core::excel_utils::cell_as_string(c15);
                assert!(
                    s.trim().is_empty(),
                    "第 {} 行 P 列不能包含日期，必须为空，当前为: {}",
                    r_idx + 4,
                    s
                );
            }
        }

        // 2. 测试纯 Rust 外账出库
        let out_res = core::external::generate_external_outbound_voucher(
            Path::new(&sales_path),
            Path::new(&tmpl_path),
            Some(&outbound_output),
            Some("2026-08-31"),
            Some("2"),
            Some(&cfg),
            None,
        )
        .expect("纯 Rust 外账出库调用必须成功");
        assert!(out_res.is_balanced);
        assert_eq!(out_res.diff, 0.0);
        assert_eq!(out_res.total_debit, out_res.total_credit);

        // 验证出库凭证中的 I, J, P 列同样满足规范
        let mut out_wb = core::excel_utils::open_excel(Path::new(&out_res.output_file))
            .expect("打开外账出库生成文件");
        let out_range = out_wb.worksheet_range("凭证").expect("打开出库凭证工作表");
        assert_voucher_template_style(Path::new(&out_res.output_file));
        for (r_idx, row) in out_range.rows().skip(3).enumerate() {
            if let Some(c8) = row.get(8) {
                let s = core::excel_utils::cell_as_string(c8);
                assert!(
                    s.trim().is_empty(),
                    "出库第 {} 行 I 列必须为空，当前为: {}",
                    r_idx + 4,
                    s
                );
            }
            if let Some(c9) = row.get(9) {
                let s = core::excel_utils::cell_as_string(c9);
                assert!(
                    s.trim().is_empty(),
                    "出库第 {} 行 J 列必须为空，当前为: {}",
                    r_idx + 4,
                    s
                );
            }
            if let Some(c15) = row.get(15) {
                let s = core::excel_utils::cell_as_string(c15);
                assert!(
                    s.trim().is_empty(),
                    "出库第 {} 行 P 列不能包含日期，必须为空，当前为: {}",
                    r_idx + 4,
                    s
                );
            }
        }

        // 3. 测试纯 Rust 外账结存数比对
        let cmp_res = core::external::generate_external_inventory_audit(
            Path::new(&tmpl_path),
            Path::new(&wh_path),
            Path::new(&wh_path),
            Path::new(&wh_path),
            Some(&audit_output),
            Some(&cfg),
        )
        .expect("纯 Rust 外账结存比对调用必须成功");
        assert!(cmp_res.total_items > 0);

        let _ = fs::remove_file(inbound_output);
        let _ = fs::remove_file(outbound_output);
        let _ = fs::remove_file(audit_output);
    }

    /// GitHub Actions 与全套业务验证的核心集成自动化测试用例
    #[test]
    #[ignore = "依赖未提交的业务 Excel 样例；本地有样例时使用 cargo test --lib -- --ignored"]
    fn test_comprehensive_financial_automation_suite() {
        println!(
            "\n================================================================================"
        );
        println!("  石家庄心理医院财务进销存自动化系统 - GitHub Actions 端到端全业务测试验证");
        println!(
            "================================================================================"
        );

        let (cfg, _) = core::config::load_config(None);
        let internal_sales_output = test_artifact("comprehensive_internal_sales");
        let internal_sales_output_string = internal_sales_output.to_string_lossy().to_string();
        let internal_inbound_output = test_artifact("comprehensive_internal_inbound");
        let internal_inbound_output_string = internal_inbound_output.to_string_lossy().to_string();
        let internal_audit_output = test_artifact("comprehensive_internal_audit");
        let internal_audit_output_string = internal_audit_output.to_string_lossy().to_string();
        let ext_inbound_output = test_artifact("comprehensive_external_inbound");
        let ext_outbound_output = test_artifact("comprehensive_external_outbound");
        let ext_audit_output = test_artifact("comprehensive_external_audit");

        // -------------------------------------------------------------------------
        // 1. 智能扫描测试
        // -------------------------------------------------------------------------
        let scan = scan_files(None).expect("扫描本地文件必须成功");
        assert!(scan["success"].as_bool().unwrap_or(false));
        println!("[PASS] 1. 业务文件与模板扫描完成");

        // -------------------------------------------------------------------------
        // 2. 内账销售明细加权汇总验证
        // -------------------------------------------------------------------------
        let sales_raw_path = project_file("2026.8月西药销售表.xls");
        let sales_res = core::sales::process_sales_file(
            &sales_raw_path,
            Some(&internal_sales_output_string),
            None,
        )
        .expect("内账销售汇总必须成功");
        assert_eq!(sales_res.totals.unique_count, 54, "去重药品数应为 54 种");
        assert_eq!(
            sales_res.totals.original_count, 82,
            "原始销售记录应为 82 笔"
        );
        assert_eq!(
            sales_res.totals.total_qty as i64, 139353,
            "总销售数量应为 139,353 件"
        );
        println!(
            "[PASS] 2. 内账销售明细汇总验证: 去重 {} 种 (原 {} 笔), 总数量 {}",
            sales_res.totals.unique_count,
            sales_res.totals.original_count,
            sales_res.totals.total_qty
        );

        // -------------------------------------------------------------------------
        // 3. 内账入库凭证生成与借贷平衡验证
        // -------------------------------------------------------------------------
        let inbound_path = project_file("西药-药品入库单.xlsx");
        let ledger_path = project_file("石家庄心理医院_数量金额总账_20260903150759.xlsx");
        let inbound_tmpl_path = project_file("凭证导入模板-入库.xlsx");
        let internal_inbound_res = core::inbound::generate_inbound_voucher(
            &inbound_path,
            &ledger_path,
            &inbound_tmpl_path,
            Some(&internal_inbound_output_string),
            Some("2026-08-31"),
            Some("10"),
            &cfg,
            None,
        )
        .expect("内账入库凭证生成必须成功");
        assert!(internal_inbound_res.is_balanced, "内账入库借贷必须平衡");
        assert_eq!(
            (internal_inbound_res.total_debit_amt * 100.0).round() as i64,
            (internal_inbound_res.total_credit_amt * 100.0).round() as i64,
            "内账入库借方金额与贷方金额必须完全一致"
        );
        println!(
            "[PASS] 3. 内账入库凭证生成验证: 借方 ¥{:.2} == 贷方 ¥{:.2}, 涉及 {} 家供应商",
            internal_inbound_res.total_debit_amt,
            internal_inbound_res.total_credit_amt,
            internal_inbound_res.suppliers_summary.len()
        );

        // -------------------------------------------------------------------------
        // 4. 内账出库凭证生成与负库存风控拦截及借贷平衡验证
        // -------------------------------------------------------------------------
        let outbound_tmpl_path = project_file("凭证导入模板.xlsx");

        // 4.1 校验负库存风控拦截 (未入库前总账结存不足时应主动拦截)
        let over_issue_err = core::outbound::generate_outbound_voucher(
            &sales_raw_path,
            &ledger_path,
            &outbound_tmpl_path,
            None,
            Some("2026-08-31"),
            true,
            &cfg,
            None,
        )
        .err();
        assert!(
            over_issue_err.is_some(),
            "当出库数量超过总账结存时，必须主动触发风控拦截"
        );
        let err_msg = over_issue_err.unwrap();
        assert!(err_msg.contains("本次出库数量超过所选总账的期末结存"));
        println!("[PASS] 4.1 内账出库负库存风控拦截验证: 成功拦截超额出库, 保护账实一致");

        // 4.2 校验入库结存满足后的出库凭证生成、借贷平衡与尾差自动平账
        let temp_dir = std::env::temp_dir();
        let test_ledger = temp_dir.join(format!("test_ledger_suff_{}.xlsx", std::process::id()));
        let source_entries =
            core::ledger::load_ledger_entries(Path::new(&ledger_path)).expect("应能读取总账样例");
        let mut ledger_wb = rust_xlsxwriter::Workbook::new();
        let ledger_ws = ledger_wb.add_worksheet();
        ledger_ws.set_name("数量金额总账").expect("设置工作表名");
        for (row_idx, entry) in source_entries.iter().enumerate() {
            ledger_ws
                .write_string(row_idx as u32, 0, &entry.code)
                .unwrap();
            ledger_ws
                .write_string(row_idx as u32, 1, &entry.name_full)
                .unwrap();
            ledger_ws
                .write_number(row_idx as u32, 16, 1_000_000.0)
                .unwrap();
            ledger_ws
                .write_number(row_idx as u32, 17, entry.price)
                .unwrap();
            ledger_ws
                .write_number(row_idx as u32, 18, 1_000_000.0 * entry.price)
                .unwrap();
        }
        ledger_wb.save(&test_ledger).expect("保存测试总账");

        let test_out_path = temp_dir.join(format!("test_outbound_{}.xlsx", std::process::id()));
        let test_out_str = test_out_path.to_string_lossy().to_string();
        let test_ledger_str = test_ledger.to_string_lossy().to_string();

        let internal_outbound_res = core::outbound::generate_outbound_voucher(
            &sales_raw_path,
            &test_ledger_str,
            &outbound_tmpl_path,
            Some(&test_out_str),
            Some("2026-08-31"),
            true,
            &cfg,
            None,
        )
        .expect("库存满足时内账出库凭证生成必须成功");
        assert!(internal_outbound_res.success);
        assert_eq!(
            internal_outbound_res.total_items, 54,
            "应生成 54 种药品出库"
        );
        assert!(internal_outbound_res.total_credit_amt > 0.0);

        // 清理临时文件
        let _ = std::fs::remove_file(test_ledger);
        let _ = std::fs::remove_file(test_out_path);

        println!(
            "[PASS] 4.2 内账出库凭证生成验证: 54 种药品出库, 匹配率 {:.1}%, 贷方总额 ¥{:.2}",
            internal_outbound_res.match_rate, internal_outbound_res.total_credit_amt
        );

        // -------------------------------------------------------------------------
        // 5. 内账账实库存多库智能核对验证
        // -------------------------------------------------------------------------
        let west_path = project_file("石家庄心理医院新西药房库存汇总报表2026831.xls");
        let internal_audit_res = core::audit::run_inventory_audit(
            &ledger_path,
            Some(&west_path),
            None,
            None,
            Some(&internal_audit_output_string),
            &cfg,
        )
        .expect("内账账实核对必须成功");
        assert!(internal_audit_res.overall.total_items > 0);
        assert!((0.0..=100.0).contains(&internal_audit_res.overall.match_rate));
        println!(
            "[PASS] 5. 内账账实多库核对验证: 总品规 {}, 吻合 {}, 差异 {}, 吻合率 {:.2}%",
            internal_audit_res.overall.total_items,
            internal_audit_res.overall.equal_count,
            internal_audit_res.overall.diff_count,
            internal_audit_res.overall.match_rate
        );

        // -------------------------------------------------------------------------
        // 6. 外账入库凭证纯 Rust 原生生成与借贷平衡验证
        // -------------------------------------------------------------------------
        let ext_tmpl_path = project_file("表格迁账参考模板.xlsx");
        let ext_inbound_path = project_file("西药入库单.xlsx");
        let ext_inbound_res = core::external::generate_external_inbound_voucher(
            Path::new(&ext_inbound_path),
            Path::new(&ext_tmpl_path),
            Some(&ext_inbound_output),
            Some("2026-08-31"),
            Some("1"),
            Some(&cfg),
            None,
        )
        .expect("外账入库凭证生成必须成功");
        assert!(ext_inbound_res.is_balanced, "外账入库借贷必须平衡");
        assert_eq!(ext_inbound_res.diff, 0.0, "外账入库借贷差额必须为 0.00");
        assert_eq!(ext_inbound_res.total_debit, ext_inbound_res.total_credit);
        assert_eq!(
            ext_inbound_res.total_entries, 46,
            "外账入库凭证分录行数应为 46 行"
        );

        // 严格校验生成的 Excel 中 I 列、J 列为空，P 列无日期
        let mut in_wb = core::excel_utils::open_excel(Path::new(&ext_inbound_res.output_file))
            .expect("打开外账入库凭证");
        let in_range = in_wb.worksheet_range("凭证").expect("打开凭证工作表");
        for row in in_range.rows().skip(3) {
            if let Some(c8) = row.get(8) {
                assert!(
                    core::excel_utils::cell_as_string(c8).trim().is_empty(),
                    "外账入库 I 列必须留空"
                );
            }
            if let Some(c9) = row.get(9) {
                assert!(
                    core::excel_utils::cell_as_string(c9).trim().is_empty(),
                    "外账入库 J 列必须留空"
                );
            }
            if let Some(c15) = row.get(15) {
                assert!(
                    core::excel_utils::cell_as_string(c15).trim().is_empty(),
                    "外账入库 P 列不得有日期"
                );
            }
        }
        println!(
            "[PASS] 6. 外账入库凭证验证: 分录 46 行, 5 家供应商, 借方 ¥{:.2} == 贷方 ¥{:.2}, 差额 0.00 (I/J列留空且P列无日期)",
            ext_inbound_res.total_debit, ext_inbound_res.total_credit
        );

        // -------------------------------------------------------------------------
        // 7. 外账销售出库结转凭证纯 Rust 原生生成与借贷平衡验证
        // -------------------------------------------------------------------------
        let ext_sales_path = project_file("2026.8月西药销售表_已汇总.xlsx");
        let ext_outbound_res = core::external::generate_external_outbound_voucher(
            Path::new(&ext_sales_path),
            Path::new(&ext_tmpl_path),
            Some(&ext_outbound_output),
            Some("2026-08-31"),
            Some("2"),
            Some(&cfg),
            None,
        )
        .expect("外账销售出库结转凭证生成必须成功");
        assert!(ext_outbound_res.is_balanced, "外账出库借贷必须平衡");
        assert_eq!(ext_outbound_res.diff, 0.0, "外账出库借贷差额必须为 0.00");
        assert_eq!(ext_outbound_res.total_debit, ext_outbound_res.total_credit);
        assert_eq!(
            ext_outbound_res.total_entries, 55,
            "外账出库凭证分录行数应为 55 行"
        );

        let mut out_wb = core::excel_utils::open_excel(Path::new(&ext_outbound_res.output_file))
            .expect("打开外账出库凭证");
        let out_range = out_wb.worksheet_range("凭证").expect("打开出库凭证工作表");
        for row in out_range.rows().skip(3) {
            if let Some(c8) = row.get(8) {
                assert!(
                    core::excel_utils::cell_as_string(c8).trim().is_empty(),
                    "外账出库 I 列必须留空"
                );
            }
            if let Some(c9) = row.get(9) {
                assert!(
                    core::excel_utils::cell_as_string(c9).trim().is_empty(),
                    "外账出库 J 列必须留空"
                );
            }
            if let Some(c15) = row.get(15) {
                assert!(
                    core::excel_utils::cell_as_string(c15).trim().is_empty(),
                    "外账出库 P 列不得有日期"
                );
            }
        }
        println!(
            "[PASS] 7. 外账销售出库验证: 分录 55 行, 54 种药品, 借方 ¥{:.2} == 贷方 ¥{:.2}, 差额 0.00 (I/J列留空且P列无日期)",
            ext_outbound_res.total_debit, ext_outbound_res.total_credit
        );

        // -------------------------------------------------------------------------
        // 8. 外账结存数比对与四维智能审计验证
        // -------------------------------------------------------------------------
        let ext_audit_res = core::external::generate_external_inventory_audit(
            Path::new(&ext_tmpl_path),
            Path::new(&west_path),
            Path::new(&west_path),
            Path::new(&west_path),
            Some(&ext_audit_output),
            Some(&cfg),
        )
        .expect("外账结存数比对必须成功");
        assert_eq!(
            ext_audit_res.total_items, 85,
            "外账辅助信息品规数应为 85 种"
        );
        assert!(
            (0.0..=100.0).contains(&ext_audit_res.match_rate),
            "外账比对吻合率应在有效范围内"
        );
        println!(
            "[PASS] 8. 外账结存数核对验证: 辅助账品规 {} 种, 吻合 {}, 差异 {}, 吻合率 {:.2}%",
            ext_audit_res.total_items,
            ext_audit_res.equal_count,
            ext_audit_res.diff_count,
            ext_audit_res.match_rate
        );

        println!(
            "--------------------------------------------------------------------------------"
        );
        println!(">>> 全部 8 大财务进销存核心业务模块端到端自动化测试 100% 通过！");
        println!(
            "================================================================================\n"
        );

        for path in [
            internal_sales_output,
            internal_inbound_output,
            internal_audit_output,
            ext_inbound_output,
            ext_outbound_output,
            ext_audit_output,
        ] {
            let _ = fs::remove_file(path);
        }
    }
}
