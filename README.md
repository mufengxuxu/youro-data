# Youro 周新客订单表生成

从 ERP API + 本地 Excel 汇总，生成 `2026年Youro运营数据周分析表` → **「6.周新客订单表」** 数据。

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
| `review/6.周新客订单表.csv` | 22 列周表数据 |
| `review/采购核对.csv` | API vs A02 采购金额对比 |
| `review/品牌复核.csv` | 品牌推断 + 置信度，供人工修正 |
| `2026年Youro运营数据周分析表（*.xlsx` | 自动追加/替换同周期行 |

## 数据流

1. `curl-orders.sh` 同源 API → 上周 `firstOrder=Y` 销售订单
2. 逐单 `curl-purchaser.sh` 同源 API → 采购金额
3. `A02` / `A05` / `A-060x` Excel → 流量、财务校验
4. 映射规则见 [`docs/周新客订单表-字段映射方案.md`](docs/周新客订单表-字段映射方案.md)

## Cookie 更新

`config.yaml` 中 `api.jsessionid` 过期时，从浏览器复制新值替换。

## 选项

```bash
python generate_weekly_new_orders.py --no-xlsx   # 只出 CSV，不写回 xlsx
python generate_weekly_new_orders.py -c other.yaml
```

## 品牌规则

可编辑 `brands.yaml` 追加型号前缀 / 关键词；低置信度订单在 `review/品牌复核.csv` 中人工填写 `brand_final`。
