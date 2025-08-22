# Compounding Planning – Requirements

**Version:** 2025-08-21  
**Scope:** 从现有 VBA 代码与最新业务规则“反推”出的需求说明，用于后续开发、维护与验收。

---

## 1. 背景与目标

我们需要在 Excel（启用宏）中对成型（Compounding）产能进行中短期滚动排程，按周/批（batch）聚合需求，输出两类核心结果：
1) **Company Allocation**：按公司/周/批的用量分配明细，确保总分配量 = 需求量（受 Horizon 限制的范围内）。  
2) **Company Batch Summary**：按批次输出锚定日期、可用能力(EffCap)、实际分配量与剩余等指标，用于可视化与对账。

系统应支持一个**可配置的计划视窗（HorizonDays）**，以及**按批能力与窗口期**的约束（见 §3 参数）。

---

## 2. 输入数据与模板结构

### 2.1 数据来源工作表（示例）
- `CompoundingTab`  
  - **I1 = HorizonDays**（整数天数；最新默认值：60）  
  - 其他用于配置/常量的单元格可后续补充

- `SHEET_SRC`（源数据表，名称以实际为准）
  - **必需列（至少）**  
    - `FG start date` / `Start date`（开始日期，日期类型）  
    - `FG end date` / `End date`（结束日期，日期类型；如无可与 start 同日）  
    - `plan order qty`（计划订单数量，件数或包数等**非吨**单位）  
  - **弃用列**  
    - 旧的 `Derived compounding usage (t)` / `usage (t)` **不再直接使用**，而是由 `plan order qty` 通过新公式换算（§2.2）。

> 代码中已有对列名的 **模糊匹配**：`FindCol(ws, Array("FG start date","Start date"))` 等；若列名变更，应在数组中补充别名，保持兼容性。

### 2.2 需求用量换算（从 plan order qty → usage(t)）
使用公式：
usage(t) = plan_order_qty × 1.07 × 10.4 ÷ 1,000 ÷ 1,000
- 其中 **1.07** 为损耗/系数；**10.4** 为单位换算因子；结果单位为 **吨 (t)**。  
- **精准性**：保留 4 位小数用于中间计算，最终汇总按 3–4 位小数输出（与历史报表一致即可）。

---

## 3. 计划参数与业务规则

- **HorizonDays**（来自 `CompoundingTab!I1`）：仅对**开始日期**在「今天（运行日）」起的 `HorizonDays` 以内的需求进行排程与分配。超窗需求可保留至下次滚动。
- **批能力（EffCap per batch）**：当前运行数据示例为 **37.5 t/批**。此值应为**可配置**（可放在 `CompoundingTab` 或常量模块）。  
- **批锚定（Anchor / First start / Last start）**：  
  - 以**周为节拍**，通常锚定到某个周六/周日或业务定义的“批起始日”；  
  - `Window(d)`：例如 **21 天** 的批窗口（可配置），窗口内集中分配；  
  - 对每个业务周/批，形成一条批记录：`Batch#、Anchor、First start、Last start、EffCap、Allocated、Window(d)、Horizon(d)`。
- **分配一致性**：  
  - `sum(Company Allocation.usage(t) within horizon) == sum(all demands usage(t) within horizon)`  
  - `Company Batch Summary.Allocated(t)` = 该批内 `Company Allocation` 的合计。  
- **溢出/剩余**：当某批 `EffCap` > 实际分配量时，记录 **剩余能力**；反之不得超配（超配为错误）。
- **日期对齐**：所有需求分配到其覆盖窗口内对应的批；如同一天内有多个批锚（不建议），按锚定优先级或时间先后处理。

---

## 4. 计算流程（高层算法）

1. **读参数**：`HorizonDays`、`EffCap per batch`、`Window(d)` 等。  
2. **取数/清洗**：从 `SHEET_SRC` 读取原始订单行，过滤出 `Start date` 在 [运行日, 运行日 + HorizonDays] 的记录。  
3. **换算用量**：按 §2.2 从 `plan order qty` 计算 `usage(t)`。  
4. **批生成**：基于时间轴（按周）和锚定规则生成批清单（含 Anchor/First/Last）。  
5. **装载/分配**：逐批遍历，将窗口内的订单（或聚合至公司/产品维度）装载到批中，直到 `Allocated ≤ EffCap`；不得超过。  
6. **一致性校验**：  
   - 计划窗内的**需求总吨数** = 所有批 `Allocated` 之和；  
   - `Company Allocation` 与 `Company Batch Summary` 交叉核验；  
   - 若发现未分配需求（计划窗内），报告为**错误**（或输出「未分配列表」工作表）。  
7. **输出**：  
   - `Company Allocation`（明细级：公司/产品/日期/批/用量）；  
   - `Company Batch Summary`（批级：EffCap、Allocated、剩余、窗口、Horizon 等）；  
   - 可选文本/日志输出（便于审计）。

---

## 5. 输出规范

### 5.1 Company Allocation（示意列）
- `Company`（或客户/业务单元）
- `Product`（可选）
- `Order ID`（可选）
- `Start date` / `End date`
- `Batch#`
- `usage (t)`（按 §2.2 计算）
- `Anchor` / `First start` / `Last start`（可选，用于追溯）
- `HorizonDays`（用于记录本次运行的入参）

**约束**：合计（计划窗内） = 所有需求（计划窗内）的 `usage(t)` 合计。

### 5.2 Company Batch Summary（示意列）
- `Batch#`
- `Anchor` / `First start` / `Last start`
- `EffCap (t)`
- `Allocated (t)`
- `Remaining (t) = EffCap - Allocated`
- `Window(d)`
- `Horizon(d)`

**约束**：`sum(Allocated)` = `Company Allocation` 的按批合计。

---

## 6. 宏/代码结构与入口

- 主要过程（示例命名，供对照/重构）  
  - `PlanFixed40Core(writeSheets As Boolean, writeText As Boolean)`：主流程  
  - 依赖工具函数：`GetSheetByNameSafe`、`FindCol`、日期/分组/批锚定生成等  
- **错误防护**  
  - 缺列：如 `FG start date` / `Start date` 或 `plan order qty` 缺失 → 抛 `Err.Raise 1001, "Missing required columns."`  
  - **重复声明/类型未定义**：确保 `Type Batch`、`Type FGOrder` 在**单一公共模块**中定义，避免重复。  
  - **未分配检查**：若计划窗内存在未分配吨数，报错或输出未分配清单。  
- **性能**  
  - 读表一次性拉取到数组；  
  - 计算在内存中完成后一次性写回；  
  - 避免在循环中频繁访问单元格。

---

## 7. 可配置项（建议落位）

- `CompoundingTab!I1`: `HorizonDays`（默认 60）  
- `CompoundingTab!<cell>`: `EffCap per batch`（例如 37.5 t）  
- `CompoundingTab!<cell>`: `Window(d)`（例如 21）  
- 损耗/换算系数：1.07、10.4（如需调整，做成常量或配置）

---

## 8. 验收标准（关键检查点）

1. **总量守恒**：计划窗内 `Company Allocation` 合计 = 需求 `usage(t)` 合计。  
2. **不超配**：任一批 `Allocated ≤ EffCap`。  
3. **双表一致**：`Company Batch Summary.Allocated` = `Company Allocation` 按批聚合。  
4. **窗内完整**：无漏分配（或输出未分配明细并标记为异常）。  
5. **参数生效**：修改 `HorizonDays`/`EffCap`/`Window(d)` 后，结果随之正确变化。  
6. **换算正确**：随机抽查 `usage(t)` 按新公式无误差（小数位一致）。

---

## 9. 运行与交付

- **运行环境**：Microsoft Excel（支持宏，`.xlsm`）；启用宏安全信任。  
- **运行入口**：在 `Developer` → `Macros` → 选择入口过程（如 `PlanFixed40Core`）→ `Run`；或在工作表按钮触发。  
- **输出位置**：  
  - `Company Allocation` 工作表（自动创建/覆盖）  
  - `Company Batch Summary` 工作表（自动创建/覆盖）  
  - 可选文本日志（按 `DefaultOutputFolder`）

---

## 10. 变更记录（与代码对应）

- **2025-08-21**  
  - 新增：以 `plan order qty` 计算 `usage(t)`：`qty × 1.07 × 10.4 / 1e6`  
  - 新增：`HorizonDays` 作为外部配置（`CompoundingTab!I1`，默认 60）  
  - 校验：总量守恒 + 不超配 + 双表一致  
  - 建议：`EffCap`、`Window(d)` 外部化为配置项

---
