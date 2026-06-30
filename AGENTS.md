# AGENTS · youro-data

自动化生成 Youro / Ronchamp 运营周分析表 CSV 的脚本仓库。

## 必读文档（按优先级）

1. [`docs/业务规则汇总.md`](docs/业务规则汇总.md) — **已确认业务规则总览**（店铺、流量、渠道、输出）
2. [`docs/周新客订单表-字段映射方案.md`](docs/周新客订单表-字段映射方案.md) — 6./4.周新客订单表 22 列映射
3. [`docs/新客转化表-字段映射方案.md`](docs/新客转化表-字段映射方案.md) — 2.新客转化表 月度汇总
4. [`docs/Youro周分析表-Sheet依赖与填表顺序.md`](docs/Youro周分析表-Sheet依赖与填表顺序.md) — 7 个 Sheet 依赖与填表顺序

## 脚本入口

```bash
python generate_weekly_new_orders.py          # 默认：CSV + 转化表 + 交叉核对 + Step3/4/6
python generate_weekly_new_orders.py --write-xlsx   # 可选写回 xlsx
```

## 关键约定（勿违背）

- **店铺**：以 A02「所属店铺」为准，不用 API `company` 过滤  
- **转化表 D/E**：仅 A05，禁止 A05+A060x 相加  
- **默认不写 xlsx**；转化表日期 = 月初 ~ `week.end_date`  
- **渠道**：无流量 ≠ 自动「其他」；个案见 `exceptions.yaml`（不改源 Excel）  
- **A0604 / A0607**：单表混 Youro+RonChamp，须按行分店铺  

## 配置与密钥

- `config.yaml` 已 gitignore，勿提交 JSESSIONID  
- 复制 `config.example.yaml` 为模板  
- 转化表单次例外：复制 `exceptions.example.yaml` → `exceptions.yaml`，按 `order_no` 追加  

## 输出目录

- `review/` 已 gitignore，为每周人工核对产物  
