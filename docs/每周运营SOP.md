# 每周运营 SOP · 脚本出数 + 填表

> **适用**：Youro + RonChamp 双店周分析表  
> **脚本**：`generate_weekly_new_orders.py`  
> **原则**：**跑脚本出 CSV → 核对 → 粘贴 xlsx**；默认不写回 xlsx。

---

## 1. 每周开始前（一次性 / 换周时）

1. **复制上周 xlsx** → 改文件名为当周末日（如 `（6.29—7.5）.xlsx`），Youro / RonChamp 各一份。  
2. 确认源 Excel 已更新到当周末：**A02 / A03 / A04 / A05 / A07 / A060x**。  
3. 编辑 `config.yaml`：
   - `week.begin_date` / `week.end_date` / `period_label`
   - `api.jsessionid`（Cookie 过期时更新）
   - `paths.a03` / `paths.a04`（店铺日汇总，Step6/7 用）

---

## 2. 跑脚本（核心一步）

```bash
pip install pyyaml openpyxl   # 首次
python generate_weekly_new_orders.py
```

**一条命令产出 `review/` 下全部 CSV**（见 [`业务规则汇总.md`](业务规则汇总.md) §1.1）。

可选：

```bash
python generate_weekly_new_orders.py --write-xlsx   # 仅写回 6./4.周新客订单表
python generate_weekly_new_orders.py --no-conversion
```

---

## 3. 核对（5～10 分钟）

| 文件 | 看什么 |
|------|--------|
| `流量交叉核对-汇总.csv` | 有 ⚠ 的店铺/业务员 |
| `渠道未归类.csv` | 需登记 `exceptions.yaml` 或确认渠道 |
| `采购核对.csv` / `品牌复核.csv` | 差异单 |
| `转化表例外应用.csv` | 例外是否预期 |

---

## 4. 粘贴 xlsx（按顺序）

### Youro · `2026年Youro运营数据周分析表（*.xlsx）`

| 顺序 | Sheet | review CSV |
|------|-------|------------|
| 1 | 6.周新客订单表 | `6.周新客订单表-Youro.csv` |
| 2 | 2.新客转化表 | `2.新客转化表.csv`（**整表覆盖**，改标题为 `6.1 - 当周末`） |
| 3 | 4.周流量分析（品牌） | `Step3-Youro-周流量品牌-*.csv` |
| 4 | 7.新流量地区分布 | `Step3-Youro-新流量地区-*.csv` |
| 5 | 5.周数据表（业务流程） | `Step4-Youro-业务流程-*.csv` |
| 6 | 1.店铺汇总 | `Step6-店铺汇总-*.csv` → **Youro 列** |
| 7 | 3.周数据表（总览） | `Step7-Youro-周数据总览-*.csv` |

### RonChamp · `2026年Ronchamp运营数据周分析表（*.xlsx）`

| 顺序 | Sheet | review CSV |
|------|-------|------------|
| 1 | 4.周新客订单表 | `4.周新客订单表-RonChamp.csv` |
| 2 | 2.周流量分析（品牌） | `Step3-RonChamp-周流量品牌-*.csv` |
| 3 | 5.新流量地区分布 | `Step3-RonChamp-新流量地区-*.csv` |
| 4 | 3.周数据表（业务流程） | `Step4-RonChamp-业务流程-*.csv` |
| 5 | 1.周数据表（总览） | `Step7-RonChamp-周数据总览-*.csv` |
| — | 1.周数据表（总览） | **买家周注销账号（个）** → **国际站后台手填**（常为 0） |

> RonChamp **无**「店铺汇总」Sheet；总览在 **Sheet 1**，与 Youro Sheet 3 口径类似。

---

## 5. 填完自洽检查

- Sheet 4 / Step3 **周总流量** ≈ Step6 / Step7 **TM+询盘**（Youro 基础段用 A03，Step3/4/5 用 A05，应逐日一致）  
- Sheet 5 / Step4 **新客订单数、金额** = 周新客 CSV 汇总  
- Sheet 3 = Step7 Youro ≈ Step6 **Youro 列**（基础+广告段）  
- Sheet 2 **C** = 当月首单总数（含未归类，减 exceptions 排除）

---

## 6. 不能脚本化的项

| 项 | 处理 |
|----|------|
| A07 意向/高潜 | **不拆店**；截止意向 / 业务流程意向段 **仅填 Youro** |
| RonChamp 买家周注销账号 | 阿里后台手填 |
| RonChamp Sheet 6/7（工作汇总、上品计划） | 本 SOP 外 |

---

## 7. 下周是否「只跑脚本」？

**是，数据侧可以只跑脚本**，前提是：

1. ✅ 已按 §1 更新 `config.yaml` 日期 + Cookie  
2. ✅ 源 Excel（A02/A03/A04/A05/A07/A060x）已维护到当周末  
3. ⚠️ 仍需 **核对 CSV** + **粘贴两个 xlsx** + RonChamp **注销账号** 手填  

脚本**不会**自动填完整 xlsx（除非对周新客表使用 `--write-xlsx`，且仅 6./4. 订单 Sheet）。

---

*版本：2026-06-30*
