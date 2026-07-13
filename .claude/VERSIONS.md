# 看板版本记录

## 回滚方式（Git 为主）

```bash
# 查看历史
git log --oneline

# 回滚单个文件到某版本
git checkout <commit> -- generate_dashboard.py 南区7月目标看板.html

# 或整个工作区回到某版本
git checkout <commit> -- .
```

> 快照目录 `.claude/versions/` 作为兜底备份，`cp` 覆盖即可。

---

## v1 — 2026-07-12（当前）
**Git commit:** `21f351e`
**快照目录：** `.claude/versions/2026-07-12_v1/`

### 变更
- **按分组标签页拆分**：线上组（蓝）和线下门店（绿）分区显示，各区域独立按完成率升序
- JS 选择器全部更新，排序/刷新/导入均跳过 `.section-header` 行
- 分区标题行无 hover 高亮

---

## 下次变更时提交新 commit，在此追加 v2、v3…
