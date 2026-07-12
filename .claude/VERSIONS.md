# 看板版本记录

## v1 — 2026-07-12（当前）
**快照目录：** `.claude/versions/2026-07-12_v1/`

### 变更
- **按分组标签页拆分**：线上组（蓝）和线下门店（绿）分区显示，各区域独立按完成率升序
- JS 选择器全部更新，排序/刷新/导入均跳过 `.section-header` 行
- 分区标题行无 hover 高亮

### 回滚方法
```
cp .claude/versions/2026-07-12_v1/generate_dashboard.py generate_dashboard.py
cp .claude/versions/2026-07-12_v1/南区7月目标看板.html 南区7月目标看板.html
```

---

## 下次变更时在此追加 v2、v3…
