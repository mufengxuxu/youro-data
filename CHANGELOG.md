# Changelog

## Unreleased

### Changed

- **A07 新客意向拆店**：通过 A05 流量表匹配客户店铺，Step4/Step6 分别输出 Youro / RonChamp；新增 `Step4-A07-店铺推断-*.csv` 核对文件
- **Luck 两单渠道修正**：`LU260605659`→TM、`LU260617825`→RFQ（原误标「其他/转介」）
- **Grace/Lily 双店拆单**：有 A05 流量时以流量表店铺为准（如 Grace · derby kembo → RonChamp）
- **Step3 地区 CSV**：新增 `周总流量合计`、`印孟巴合计`（印度+孟加拉国+巴基斯坦）；国家名合并（印尼/沙特/孟加拉别名）

## 2026-06-30

### Added

- **Step3–7** 周分析 CSV：品牌/地区/业务流程/店铺汇总/总览（Youro + RonChamp）
- [`docs/每周运营SOP.md`](docs/每周运营SOP.md) — 每周跑脚本 + 粘贴 xlsx 流程
- `paths.a03` / `paths.a04` — 店铺日汇总（Step6/7）
- `exceptions.yaml` / Step3 **其中变频器** / A07 意向仅 Youro 规则

### Changed

- 重写 [`docs/Youro周分析表-Sheet依赖与填表顺序.md`](docs/Youro周分析表-Sheet依赖与填表顺序.md) v2.0（脚本全覆盖、RonChamp 对照）
- [`docs/业务规则汇总.md`](docs/业务规则汇总.md) v1.1 — 完整输出清单、注销账号、每周 SOP 引用

### Fixed

- RonChamp **Step7-RonChamp-周数据总览** CSV（Sheet 1 总览）

## 2026-06-30 (earlier)

- 初始脚本：周新客订单表 + 店铺拆分 + 新客转化表 + 流量交叉核对
- 渠道规则：TM / RFQ / 其他 / 未归类；Grace/Lily 双店 A060x
- 2026-06 转化表个案（Luck/Grace/Ennerson/Hema）
