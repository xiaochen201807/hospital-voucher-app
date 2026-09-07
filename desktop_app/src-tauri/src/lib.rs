use std::path::{Path, PathBuf};
use std::process::Command;
use serde_json::Value;

/// 运行目标枚举：独立二进制或 Python 脚本
enum RunnerTarget {
    Standalone(PathBuf),
    PythonScript(PathBuf),
}

/// 智能寻找 runner（优先寻找独立打包的 runner.exe，其次回退 runner.py）
fn find_runner_target() -> Result<RunnerTarget, String> {
    let exe_name = if cfg!(windows) { "runner.exe" } else { "runner" };

    // 1. 优先检查打包内置的独立可执行文件 runner.exe / runner
    if let Ok(current_exe) = std::env::current_exe() {
        if let Some(exe_dir) = current_exe.parent() {
            let standalone_candidates = [
                exe_dir.join(exe_name),
                exe_dir.join("resources").join(exe_name),
                exe_dir.join("..").join("resources").join(exe_name),
                exe_dir.join("backend").join(exe_name),
            ];
            for candidate in &standalone_candidates {
                if candidate.exists() {
                    return Ok(RunnerTarget::Standalone(candidate.clone()));
                }
            }
        }
    }

    // 检查工作区相对路径下的 runner.exe
    let local_standalone = [
        PathBuf::from(format!("backend/{}", exe_name)),
        PathBuf::from(format!("desktop_app/backend/{}", exe_name)),
        PathBuf::from(format!("src-tauri/resources/{}", exe_name)),
    ];
    for candidate in &local_standalone {
        if candidate.exists() {
            return Ok(RunnerTarget::Standalone(candidate.clone()));
        }
    }

    // 2. 回退寻找 Python 脚本 runner.py
    let script_candidates = [
        PathBuf::from("backend/runner.py"),
        PathBuf::from("desktop_app/backend/runner.py"),
        PathBuf::from("../backend/runner.py"),
        PathBuf::from("/Users/youyou/Downloads/python/desktop_app/backend/runner.py"),
    ];

    for candidate in &script_candidates {
        if candidate.exists() {
            return Ok(RunnerTarget::PythonScript(candidate.canonicalize().unwrap_or(candidate.clone())));
        }
    }

    if let Ok(exe_path) = std::env::current_exe() {
        if let Some(exe_dir) = exe_path.parent() {
            let p1 = exe_dir.join("backend").join("runner.py");
            if p1.exists() {
                return Ok(RunnerTarget::PythonScript(p1));
            }
            let p2 = exe_dir.join("..").join("backend").join("runner.py");
            if p2.exists() {
                return Ok(RunnerTarget::PythonScript(p2));
            }
            let p3 = exe_dir.join("..").join("..").join("backend").join("runner.py");
            if p3.exists() {
                return Ok(RunnerTarget::PythonScript(p3));
            }
        }
    }

    Err("未找到调度引擎（未找到 runner.exe 或 backend/runner.py），请检查安装目录".into())
}

/// 执行调度引擎并解析 JSON 结果
fn execute_runner(args: &[&str]) -> Result<Value, String> {
    let runner_target = find_runner_target()?;
    let mut cmd = match &runner_target {
        RunnerTarget::Standalone(exe_path) => {
            let mut c = Command::new(exe_path);
            for arg in args {
                c.arg(arg);
            }
            if let Some(parent) = exe_path.parent() {
                c.current_dir(parent);
            }
            c
        }
        RunnerTarget::PythonScript(py_script) => {
            let py_cmd = if cfg!(windows) { "python" } else { "python3" };
            let mut c = Command::new(py_cmd);
            c.arg(py_script);
            for arg in args {
                c.arg(arg);
            }
            if let Some(parent) = py_script.parent().and_then(|p| p.parent()) {
                c.current_dir(parent);
            }
            c
        }
    };

    let output = cmd.output().map_err(|e| format!("执行调度引擎失败: {}", e))?;
    let stdout_str = String::from_utf8_lossy(&output.stdout).trim().to_string();
    let stderr_str = String::from_utf8_lossy(&output.stderr).trim().to_string();

    if stdout_str.is_empty() {
        if !output.status.success() || !stderr_str.is_empty() {
            return Err(format!("Python 脚本执行异常: {}", stderr_str));
        }
        return Err("Python 脚本未返回任何输出".into());
    }

    // 从输出中截取最后一行 JSON（防备环境中有前置 warning 输出）
    let json_line = stdout_str
        .lines()
        .rev()
        .find(|line| line.starts_with('{') && line.ends_with('}'))
        .unwrap_or(&stdout_str);

    match serde_json::from_str::<Value>(json_line) {
        Ok(json_val) => Ok(json_val),
        Err(e) => Err(format!("解析 Python 返回 JSON 失败: {}。原始输出: {}", e, stdout_str)),
    }
}

#[tauri::command]
fn scan_files(dir: Option<String>) -> Result<Value, String> {
    let mut args = vec!["scan"];
    let dir_str;
    if let Some(d) = &dir {
        dir_str = d.clone();
        args.push("--dir");
        args.push(&dir_str);
    }
    execute_runner(&args)
}

#[tauri::command]
fn execute_sales_process(file: String, output: Option<String>, sheet_name: Option<String>) -> Result<Value, String> {
    let mut args = vec!["process_sales", "--file", &file];
    let out_str;
    if let Some(o) = &output {
        out_str = o.clone();
        args.push("--output");
        args.push(&out_str);
    }
    let sheet_str;
    if let Some(s) = &sheet_name {
        sheet_str = s.clone();
        args.push("--sheet-name");
        args.push(&sheet_str);
    }
    execute_runner(&args)
}

#[tauri::command]
fn execute_outbound_voucher(
    sales: String,
    ledger: String,
    template: String,
    output: Option<String>,
    date: Option<String>,
    fallback_price: bool,
    config: Option<String>,
) -> Result<Value, String> {
    let mut args = vec![
        "generate_voucher",
        "--sales", &sales,
        "--ledger", &ledger,
        "--template", &template,
    ];
    let out_str;
    if let Some(o) = &output {
        out_str = o.clone();
        args.push("--output");
        args.push(&out_str);
    }
    let date_str;
    if let Some(d) = &date {
        date_str = d.clone();
        args.push("--date");
        args.push(&date_str);
    }
    if fallback_price {
        args.push("--fallback-price");
    }
    let cfg_str;
    if let Some(c) = &config {
        cfg_str = c.clone();
        args.push("--config");
        args.push(&cfg_str);
    }
    execute_runner(&args)
}

#[tauri::command]
fn execute_inbound_voucher(
    inbound: String,
    ledger: String,
    template: String,
    output: Option<String>,
    date: Option<String>,
    voucher_no: Option<String>,
    config: Option<String>,
) -> Result<Value, String> {
    let mut args = vec![
        "generate_inbound_voucher",
        "--inbound", &inbound,
        "--ledger", &ledger,
        "--template", &template,
    ];
    let out_str;
    if let Some(o) = &output {
        out_str = o.clone();
        args.push("--output");
        args.push(&out_str);
    }
    let date_str;
    if let Some(d) = &date {
        date_str = d.clone();
        args.push("--date");
        args.push(&date_str);
    }
    let vno_str;
    if let Some(v) = &voucher_no {
        vno_str = v.clone();
        args.push("--voucher-no");
        args.push(&vno_str);
    }
    let cfg_str;
    if let Some(c) = &config {
        cfg_str = c.clone();
        args.push("--config");
        args.push(&cfg_str);
    }
    execute_runner(&args)
}

#[tauri::command]
fn get_config(config: Option<String>) -> Result<Value, String> {
    let mut args = vec!["get_config"];
    let cfg_str;
    if let Some(c) = &config {
        cfg_str = c.clone();
        args.push("--config");
        args.push(&cfg_str);
    }
    execute_runner(&args)
}

#[tauri::command]
fn save_config(data: String, config: Option<String>) -> Result<Value, String> {
    let mut args = vec!["save_config", "--data", &data];
    let cfg_str;
    if let Some(c) = &config {
        cfg_str = c.clone();
        args.push("--config");
        args.push(&cfg_str);
    }
    execute_runner(&args)
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
            execute_outbound_voucher,
            execute_inbound_voucher,
            get_config,
            save_config,
            open_in_system,
            show_in_folder
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}

