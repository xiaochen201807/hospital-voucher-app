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
}

/// 生成凭证时无法匹配总账科目的明细。
#[derive(Debug, Serialize, Deserialize, Clone)]
pub struct UnmatchedDrug {
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
}
