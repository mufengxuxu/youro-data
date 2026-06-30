# Changelog

## Unreleased

### Added

- Step7 Youro 周数据总览 CSV（Sheet 3，与 Step6 Youro 同源）
- Step6 店铺汇总 CSV（A03/A04 + 周新客；A07 截止意向仅 Youro）
- `generate_weekly_new_orders.py` 自动生成 **步骤③/④** 核对 CSV（品牌流量、地区分布、业务流程、A07 意向/高潜明细）
- `exceptions.yaml` / `exceptions.example.yaml` — 转化表单次例外（按 order_no，不改源 Excel）
- `review/转化表例外应用.csv` 输出
- `docs/业务规则汇总.md` — 汇总店铺拆分、A05/A060x 交叉核对、渠道分类、Grace/Lily 双店表等已确认规则
- `AGENTS.md` — Agent 维护指引与文档索引

### Changed

- 2026-06 转化表：Luck×2 / Grace Anis / Ennerson Adesola 渠道覆盖；Hema latha-G 转化表排除

- 更新 `docs/新客转化表-字段映射方案.md`、`docs/周新客订单表-字段映射方案.md` 与业务规则对齐
- `README.md` 增加 [`docs/业务规则汇总.md`](docs/业务规则汇总.md) 链接

## 2026-06-30

- 初始脚本：周新客订单表 + 店铺拆分 + 新客转化表 + 流量交叉核对
- 渠道规则：TM / RFQ / 其他 / 未归类；A07 意向、Sally 转介、Alan分配
- A0604/A0607 按 Youro/RonChamp 分行；客户别名 airo→Jairo Mafra
