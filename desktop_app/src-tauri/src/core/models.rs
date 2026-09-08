use serde::{Deserialize, Serialize};

/// 财务总账中的存货科目。
///
/// 入库、出库和账实核对都依赖这套统一模型，避免各模块分别解析同一行总账数据。
#[derive(Debug, Clone)]
pub struct LedgerEntry {
    pub code: String,
    pub aux_code: String,
    pub name_full: String,
    pub drug_name: String,
    pub spec: String,
    pub price: f64,
    pub end_qty: f64,
    pub end_amount: f64,
}

/// 供前端人工指定的财务总账候选科目。
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct LedgerCandidateOption {
    pub code: String,
    pub name: String,
    pub spec: String,
    pub qty: f64,
    pub price: f64,
    pub amount: f64,
}

/// 凭证生成时由前端确认的人工科目映射。
///
/// `id` 对应本次入库单或销售汇总表中的有效明细行编号，`ledger_code`
/// 使用财务总账中的完整科目编码（例如 `1201_XY0069`）。
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct ConfirmedLedgerMapping {
    pub id: usize,
    pub ledger_code: String,
}

/// 生成凭证时无法匹配总账科目的明细。
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct UnmatchedDrug {
    pub id: usize,
    pub name: String,
    pub target_name: String,
    pub spec: String,
    pub factory: String,
    pub supplier: String,
    pub qty: f64,
    pub price: f64,
    pub in_price: f64,
    pub amount: f64,
    pub in_amt: f64,
    pub reason: String,
    pub candidates: Vec<LedgerCandidateOption>,
}
