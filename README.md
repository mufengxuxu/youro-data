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
| `review/采购核对.csv` | API vs A02 采购金额对比 |
| `review/品牌复核.csv` | 品牌推断 + 置信度，供人工修正 |

**默认只出 CSV**，周分析 xlsx 由人工粘贴更新。需要脚本写回 xlsx 时使用 `--write-xlsx`。

**店铺判定**：以 A02「所属店铺」（屿路 / 镕川）为准，对应 Youro Sheet 6 / RonChamp Sheet 4。

**转化表日期**：自动取 `week.end_date` 所在月的 **1 日 ~ week.end_date**（例：周次 6.22—6.28 → 转化表标题 `6.1 - 6.28`）。

## 数据流

1. `curl-orders.sh` 同源 API → 上周 `firstOrder=Y` 销售订单
2. 逐单 `curl-purchaser.sh` 同源 API → 采购金额
3. `A02` / `A05` / `A-060x` Excel → 流量、财务校验、店铺归属
4. 映射规则见：
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
