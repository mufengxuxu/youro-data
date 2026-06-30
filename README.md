# Youro 周新客订单表生成

从 ERP API + 本地 Excel 汇总，生成运营周分析表数据，支持 **屿路（Youro）/ 镕川（RonChamp）** 分店铺输出，以及 **2.新客转化表** 月度汇总。

## 快速使用

```bash
pip install pyyaml openpyxl
python generate_weekly_new_orders.py
```

```bash
cp config.example.yaml config.yaml
# 编辑 config.yaml，填入 JSESSIONID 与日期范围
```

## 输出

| 文件 | 说明 |
|------|------|
| `review/6.周新客订单表-Youro.csv` | 屿路当周首单（22 列） |
| `review/4.周新客订单表-RonChamp.csv` | 镕川当周首单（22 列） |
| `review/2.新客转化表.csv` | **当月累计**新客转化（月初 ~ `week.end_date`） |
| `review/转化表例外应用.csv` | 本次从 `exceptions.yaml` 应用的转化表个案 |
| `review/渠道未归类.csv` | 无流量且非明确「其他」的首单，计入 C 但不进 P 列 |
| `review/流量交叉核对-汇总.csv` | A05 vs A060x 新流量数量交叉对比（按店铺/业务员） |
| `review/流量交叉核对-明细.csv` | 两边不一致的客户明细（仅A05/仅A060x/字段差） |
| `review/采购核对.csv` | API vs A02 采购金额对比 |
| `review/品牌复核.csv` | 品牌推断 + 置信度，供人工修正 |
| `review/Step3-{Youro\|RonChamp}-周流量品牌-MMDD-MMDD.csv` | **步骤③** 品牌流量 + **其中变频器** + 产品汇总行「变频器」 |
| `review/Step3-{Youro\|RonChamp}-其它杂类明细-MMDD-MMDD.csv` | 步骤③「其它杂类」逐条明细 |
| `review/Step3-{Youro\|RonChamp}-新流量地区-MMDD-MMDD.csv` | 步骤③ 国家排名 |
| `review/Step4-{Youro\|RonChamp}-业务流程-MMDD-MMDD.csv` | **步骤④** 业务流程单行（TM/L1+/L3+/新客订单；Youro 含 A07 意向/高潜） |
| `review/Step4-Youro-意向订单明细-MMDD-MMDD.csv` | A07 当周意向订单明细（仅 Youro） |
| `review/Step4-Youro-高潜明细-MMDD-MMDD.csv` | A07 当周高潜订单明细（仅 Youro） |
| `review/Step7-{Youro\|RonChamp}-周数据总览-MMDD-MMDD.csv` | **步骤⑦** 总览（Youro→Sheet 3；RonChamp→Sheet 1，均与 Step6 同店同源） |
| `review/Step6-店铺汇总-MMDD-MMDD.csv` | **步骤⑥** Sheet 1 双店指标（A03/A04 基础+运营 + 周新客订单段；**截止意向仅 Youro**） |

**默认只出 CSV**，周分析 xlsx 由人工粘贴更新。需要脚本写回 xlsx 时使用 `--write-xlsx`。

**店铺判定**：以 A02「所属店铺」（屿路 / 镕川）为准，对应 Youro Sheet 6 / RonChamp Sheet 4。

**转化表日期**：自动取 `week.end_date` 所在月的 **1 日 ~ week.end_date**（例：周次 6.22—6.28 → 转化表标题 `6.1 - 6.28`）。

**转化表个案**：见 [`docs/业务规则汇总.md`](docs/业务规则汇总.md) §12。

## 每周流程（下周起）

1. 更新 `config.yaml` 周次 + Cookie；源 Excel 维护到周末  
2. `python generate_weekly_new_orders.py` → 出齐 `review/` 全部 CSV  
3. 核对交叉核对 / 未归类 / 采购 / 品牌  
4. 按 **[`docs/每周运营SOP.md`](docs/每周运营SOP.md)** 粘贴 Youro + RonChamp 两个 xlsx  
5. RonChamp Sheet1 手填 **买家周注销账号**（常为 0）

> **只跑脚本即可出数**；粘贴 xlsx 与注销账号仍需人工。

## 数据流

1. `curl-orders.sh` 同源 API → 上周 `firstOrder=Y` 销售订单
2. 逐单 `curl-purchaser.sh` 同源 API → 采购金额
3. `A02` / `A05` / `A-060x` Excel → 流量、财务校验、店铺归属
4. 映射规则见：
   - **[`docs/业务规则汇总.md`](docs/业务规则汇总.md)** — 已确认业务规则总览（优先阅读）
   - [`docs/周新客订单表-字段映射方案.md`](docs/周新客订单表-字段映射方案.md)
   - [`docs/新客转化表-字段映射方案.md`](docs/新客转化表-字段映射方案.md)

## Cookie 更新

`config.yaml` 中 `api.jsessionid` 过期时，从浏览器复制新值替换。

## 选项

```bash
python generate_weekly_new_orders.py                  # CSV + 转化表（默认）
python generate_weekly_new_orders.py --write-xlsx     # 额外写回周分析 xlsx
python generate_weekly_new_orders.py --no-conversion  # 跳过 2.新客转化表
python generate_weekly_new_orders.py -c other.yaml
```

## 品牌规则

可编辑 `brands.yaml` 追加型号前缀 / 关键词；低置信度订单在 `review/品牌复核.csv` 中人工填写 `brand_final`。
