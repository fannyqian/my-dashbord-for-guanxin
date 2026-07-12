#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""南区7月目标进展看板 — 参考业绩进展.xlsx Sheet1结构，数据从底表实时计算"""

import pandas as pd
import numpy as np

# ==================== 加载数据 ====================
orders = pd.read_excel(r'C:\Users\fanny\Desktop\数据日日\7月收入-青少年体验课订单明细.xlsx')
target_xls = pd.ExcelFile(r'C:\Users\fanny\Desktop\数据日日\南区-目标.xlsx')
target_summary = pd.read_excel(target_xls, sheet_name='Sheet1')
target_online_person = pd.read_excel(target_xls, sheet_name='线上个人')
target_cm = pd.read_excel(target_xls, sheet_name='线下CM')

# parse back_at
def parse_back_at(val):
    try:
        t = pd.to_datetime(val)
        return t if t.year >= 2000 else pd.NaT
    except:
        return pd.NaT

orders['back_dt'] = orders['back_at'].apply(parse_back_at)

# ==================== 映射表 ====================
# 参考7.9 sheet角色定义: 销售/班主任(含观心学院)/门店
SALES_ORGS = ['课程销售1组','课程销售2组','课程销售3组','课程销售4组','课程销售5组']
ADVISOR_ORGS = ['班主任2组','班主任3组']  # 仅班主任2组和3组计入
STORE_MAP = {
    '上海普陀店': '普陀','上海徐汇店': '徐汇','云南分公司': '昆明',
    '南京分公司': '南京','广州观心诊所': '广州','成都分公司': '成都',
    '杭州分公司': '杭州','武汉观心诊所': '武汉','深圳分公司': '深圳',
    '珠海分公司': '珠海','福州观心诊所': '福州',
}
# 组织→分组名 (匹配目标表)
ORG_TO_GROUP = {
    '课程销售1组':'课程1组','课程销售2组':'课程2组','课程销售3组':'课程3组',
    '课程销售4组':'销售4组','课程销售5组':'销售5组',
    '班主任1组':'班主任1组','班主任2组':'班主任2组','班主任3组':'班主任3组',
    '班主任组':'班主任组',
}
ORG_TO_GROUP.update(STORE_MAP)

# 角色分类 (匹配7.9 sheet)
def classify_role(org):
    if org in SALES_ORGS: return '销售'
    if org in ADVISOR_ORGS: return '班主任'
    if org in STORE_MAP: return '门店'
    return '其他'

orders['role'] = orders['组织'].apply(classify_role)
orders['target_group'] = orders['组织'].map(ORG_TO_GROUP)

# ==================== 解析目标表 ====================
# Sheet1 分组目标
target_map = {}
current_section = None
section_tgts = {'线上':0,'线下门店':0}
for _, row in target_summary.iterrows():
    label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
    g = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
    t = row.iloc[3]
    if g in ('分组','nan','') or g == '合计':
        if label in ('线上','线下门店'): current_section = label
        continue
    if label in ('线上','线下门店'): current_section = label
    try:
        val = float(t)
        target_map[g] = val
        if current_section: section_tgts[current_section] += val
    except: continue

# 线上个人目标
online_person_target = {}
for _, row in target_online_person.iterrows():
    online_person_target[(row['分组'], row['销售'])] = float(row['目标'])

# 线下CM目标
cm_list = []
for _, row in target_cm.iterrows():
    store = row.iloc[0]; name = row.iloc[1]; tgt = row.iloc[2]
    done = row.iloc[3]; prev = row.iloc[5]; refund = row.iloc[7]
    if pd.isna(store) or str(store).strip() in ('分院',''): continue
    cm_list.append({
        'store': str(store).strip(), 'name': str(name).strip(),
        'target': float(tgt) if pd.notna(tgt) else 0,
        'done_locked': float(done) if pd.notna(done) else 0,
        'prev_month': float(prev) if pd.notna(prev) else 0,
        'refund_locked': float(refund) if pd.notna(refund) else 0,
    })
cm_df = pd.DataFrame(cm_list)
putuo_locked_done = cm_df[cm_df['store']=='普陀']['done_locked'].sum()

# ==================== 同期筛选 ====================
JULY_START = '2026-07-01'; JULY_END = '2026-07-11'
JUNE_START = '2026-06-01'; JUNE_END = '2026-06-11'

# 收入: 按 pay_at
def period_income(df, start, end):
    return df[(df['pay_at']>=start)&(df['pay_at']<end)]['pay_amount'].sum()

# 退费: 按 back_dt
def period_refund(df, start, end):
    return df[(df['back_dt']>=start)&(df['back_dt']<end)]['pay_amount'].sum()

# ==================== 计算函数 ====================
def safe_div(a, b):
    return a / b if b != 0 else 0

def calc_metrics(grp_df, tgt, is_putuo=False):
    """计算一组指标: 目标/完成/完成率/退费/上月退费/退费环比/上月完成/完成环比/净流水/上月净流水/净流水环比"""
    done_raw = period_income(grp_df, JULY_START, JULY_END)
    prev_raw = period_income(grp_df, JUNE_START, JUNE_END)
    ref_raw = period_refund(grp_df, JULY_START, JULY_END)
    prev_ref_raw = period_refund(grp_df, JUNE_START, JUNE_END)

    if is_putuo:
        done = putuo_locked_done
    else:
        done = done_raw

    rate = safe_div(done, tgt)
    ref_chg = safe_div(ref_raw - prev_ref_raw, prev_ref_raw)
    done_chg = safe_div(done_raw - prev_raw, prev_raw)
    net = done - ref_raw
    prev_net = prev_raw - prev_ref_raw
    net_chg = safe_div(net - prev_net, prev_net)

    return {
        'target': tgt, 'done': done, 'rate': rate,
        'refund': ref_raw, 'prev_refund': prev_ref_raw, 'refund_chg': ref_chg,
        'prev_done': prev_raw, 'done_chg': done_chg,
        'net': net, 'prev_net': prev_net, 'net_chg': net_chg,
    }

# ==================== Level 1: 按角色 ====================
role_metrics = {}
putuo_order_done = period_income(orders[orders['target_group']=='普陀'], JULY_START, JULY_END)
putuo_delta = putuo_locked_done - putuo_order_done  # 普陀锁定增量

for role in ['销售','班主任','门店']:
    grp = orders[orders['role']==role]
    tgt = sum(target_map.get(g,0) for g in grp['target_group'].dropna().unique())
    m = calc_metrics(grp, tgt)
    if role == '门店':
        # 普陀部分用锁定值替换
        m['done'] = m['done'] + putuo_delta
        m['rate'] = safe_div(m['done'], tgt)
        m['net'] = m['done'] - m['refund']
        m['net_chg'] = safe_div(m['net'] - m['prev_net'], m['prev_net'])
    role_metrics[role] = m

# 南区合计
total_tgt = sum(r['target'] for r in role_metrics.values())
total_grp = orders[orders['role'].isin(['销售','班主任','门店'])]
total_metrics = calc_metrics(total_grp, total_tgt)
# 合并普陀锁定
total_metrics['done'] = total_metrics['done'] + putuo_delta
total_metrics['rate'] = safe_div(total_metrics['done'], total_tgt)
total_metrics['net'] = total_metrics['done'] - total_metrics['refund']
total_metrics['net_chg'] = safe_div(total_metrics['net'] - total_metrics['prev_net'], total_metrics['prev_net'])

# ==================== Level 2: 按分组 ====================
# 所有分组 (按完成率升序 = 最差的排前面)
ALL_GROUPS = ['课程1组','课程2组','课程3组','销售4组','销售5组',
              '班主任2组','班主任3组',
              '普陀','昆明','成都','广州','深圳','珠海','福州','杭州','南京','武汉','徐汇']

group_metrics = {}
for g in ALL_GROUPS:
    grp = orders[orders['target_group']==g]
    tgt = target_map.get(g, 0)
    if tgt == 0: continue  # skip groups without target
    is_putuo = (g == '普陀')
    group_metrics[g] = calc_metrics(grp, tgt, is_putuo)

# 拆分线上组 / 线下门店
ONLINE_GROUPS = ['课程1组','课程2组','课程3组','销售4组','销售5组','班主任2组','班主任3组']
OFFLINE_GROUPS = ['普陀','昆明','成都','广州','深圳','珠海','福州','杭州','南京','武汉','徐汇']

# 各自按完成率升序排列（差的在前）
ONLINE_ORDER = sorted([g for g in ONLINE_GROUPS if g in group_metrics], key=lambda g: group_metrics[g]['rate'])
OFFLINE_ORDER = sorted([g for g in OFFLINE_GROUPS if g in group_metrics], key=lambda g: group_metrics[g]['rate'])

# ==================== Level 3a: 线上个人 ====================
person_online_list = []
for (grp_name, person_name), tgt in online_person_target.items():
    p_orders = orders[(orders['seller_name']==person_name)]
    m = calc_metrics(p_orders, tgt)
    person_online_list.append({
        'group': grp_name, 'name': person_name, **m
    })
person_online_df = pd.DataFrame(person_online_list)

# ==================== Level 3b: 线下个案 ====================
person_cm_list = []
for _, cm in cm_df.iterrows():
    store = cm['store']; name = cm['name']; tgt = cm['target']
    is_putuo = (store == '普陀')
    if is_putuo:
        done = cm['done_locked']
        # 普陀退费也要从底表查 (by seller_name)
        p_orders = orders[(orders['seller_name']==name)]
        ref_raw = period_refund(p_orders, JULY_START, JULY_END)
        prev_raw = cm['prev_month']
        prev_ref_raw = cm['refund_locked']
    else:
        p_orders = orders[(orders['seller_name']==name)]
        done = period_income(p_orders, JULY_START, JULY_END)
        ref_raw = period_refund(p_orders, JULY_START, JULY_END)
        prev_raw = period_income(p_orders, JUNE_START, JUNE_END)
        prev_ref_raw = period_refund(p_orders, JUNE_START, JUNE_END)

    rate = safe_div(done, tgt)
    ref_chg = safe_div(ref_raw - prev_ref_raw, prev_ref_raw)
    done_chg = safe_div(done - prev_raw, prev_raw)
    net = done - ref_raw
    prev_net = prev_raw - prev_ref_raw
    net_chg = safe_div(net - prev_net, prev_net)

    person_cm_list.append({
        'store': store, 'name': name, 'target': tgt, 'done': done, 'rate': rate,
        'refund': ref_raw, 'prev_refund': prev_ref_raw, 'refund_chg': ref_chg,
        'prev_done': prev_raw, 'done_chg': done_chg,
        'net': net, 'prev_net': prev_net, 'net_chg': net_chg,
        'is_putuo': is_putuo,
    })
person_cm_df = pd.DataFrame(person_cm_list)

# ==================== HTML 渲染 ====================
def fmt(n):
    if abs(n) >= 1e8: return f'{n/1e8:.2f}亿'
    if abs(n) >= 1e4: return f'{n/1e4:.0f}万'
    return f'{n:,.0f}'

def pct1(n): return f'{n:+.1f}%' if n and n != float('inf') else '—'
def pct2(n): return f'{n*100:.1f}%'
def chg_html(val):
    if val is None or val == float('inf') or val == -float('inf'): return '<span style="color:#94a3b8">—</span>'
    c = '#22c55e' if val >= 0 else '#ef4444'
    arrow = '↑' if val >= 0 else '↓'
    return f'<span style="color:{c};font-weight:600">{arrow}{abs(val)*100:.1f}%</span>'

def rate_bar(val):
    """val is 0-1 ratio. Large prominent progress bar with percentage."""
    p = min(val * 100, 100)
    if p >= 80: color = '#22c55e'; bg = '#dcfce7'
    elif p >= 50: color = '#f59e0b'; bg = '#fef3c7'
    elif p >= 30: color = '#f97316'; bg = '#ffedd5'
    elif p >= 10: color = '#ef4444'; bg = '#fee2e2'
    else: color = '#dc2626'; bg = '#fecaca'
    return f'''<div style="display:flex;align-items:center;gap:8px;min-width:160px">
  <div style="flex:1;height:18px;background:{bg};border-radius:9px;overflow:hidden;position:relative;border:1px solid #e2e8f0">
    <div style="width:{p}%;height:100%;background:{color};border-radius:9px;transition:width 0.3s"></div>
    <div style="position:absolute;top:0;left:0;width:100%;height:100%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#1e293b">{pct2(val)}</div>
  </div>
</div>'''

def role_badge(role):
    colors = {'销售':('#dbeafe','#1e40af'),'班主任':('#fce7f3','#9d174d'),'门店':('#d1fae5','#065f46')}
    bg, fg = colors.get(role,('#f3f4f6','#6b7280'))
    return f'<span style="background:{bg};color:{fg};padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600">{role}</span>'

TIME_PROGRESS = 10/31  # 月度时间进度 (~32.3%)

def trend_label(val):
    """趋势判定"""
    ratio = val / TIME_PROGRESS if TIME_PROGRESS > 0 else 1
    if ratio >= 1.0:
        return 'ontrack', '正常推进'
    elif ratio >= 0.5:
        return 'warning', '可能完不成'
    else:
        return 'critical', '必定完不成'

def refund_td(refund, done, person_type='group'):
    """退费率单元格：纯百分比 + 条件标红
       - 销售/班主任个人: >30% 标红
       - 门店CM个人: >10% 标红
       - 分组级: 不标红"""
    rate = safe_div(refund, done)
    pct_str = pct2(rate)
    if rate == 0 and done == 0:
        return '<td class="num">—</td>'
    # 个人级条件标红
    if person_type == 'sales_advisor' and rate > 0.3:
        return f'<td class="num"><span style="color:#dc2626;font-weight:700">{pct_str}</span></td>'
    if person_type == 'cm' and rate > 0.1:
        return f'<td class="num"><span style="color:#dc2626;font-weight:700">{pct_str}</span></td>'
    return f'<td class="num">{pct_str}</td>'

def rate_bar_with_trend(val):
    """完成率进度条 + 趋势标记"""
    p = min(val * 100, 100)
    trend, tip = trend_label(val)
    if trend == 'ontrack':
        color = '#22c55e'; bg = '#dcfce7'; bar_bg = '#f0fdf4'
        badge = ''
    elif trend == 'warning':
        color = '#f59e0b'; bg = '#fef3c7'; bar_bg = '#fffbeb'
        badge = f'<span style="font-size:10px;background:#fef3c7;color:#b45309;padding:1px 6px;border-radius:8px;font-weight:700;margin-left:4px" title="{tip}">⚠️ 关注</span>'
    else:
        color = '#dc2626'; bg = '#fecaca'; bar_bg = '#fef2f2'
        badge = f'<span style="font-size:10px;background:#fee2e2;color:#b91c1c;padding:1px 6px;border-radius:8px;font-weight:700;margin-left:4px" title="{tip}">🚫 高危</span>'
    return f'''<div style="display:flex;align-items:center;gap:6px;min-width:180px">
  <div style="flex:1;height:20px;background:{bar_bg};border-radius:10px;overflow:hidden;position:relative;border:1px solid #e2e8f0">
    <div style="width:{p}%;height:100%;background:{color};border-radius:10px;transition:width 0.3s"></div>
    <div style="position:absolute;top:0;left:0;width:100%;height:100%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#1e293b;text-shadow:0 0 4px rgba(255,255,255,0.8)">{pct2(val)}</div>
  </div>
  {badge}
</div>'''

def rate_td(val):
    """完成率单元格：进度条+趋势"""
    return f'<td style="min-width:170px">{rate_bar_with_trend(val)}</td>'
    colors = {'销售':('#dbeafe','#1e40af'),'班主任':('#fce7f3','#9d174d'),'门店':('#d1fae5','#065f46')}
    bg, fg = colors.get(role,('#f3f4f6','#6b7280'))
    return f'<span style="background:{bg};color:{fg};padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600">{role}</span>'

# Build role-level rows (Level 1)
def role_row(role, m):
    return f'''<tr>
  <td style="font-size:14px">{role_badge(role)}</td>
  <td class="num">{fmt(m['target'])}</td>
  <td class="num"><b>{fmt(m['done'])}</b></td>
  {rate_td(m['rate'])}
  <td class="num">{fmt(m['refund'])}</td>
  {refund_td(m['refund'], m['done'])}
  <td class="num"><b>{fmt(m['net'])}</b></td>
</tr>'''

# Build group-level rows (Level 2)
def group_row(g, m, role):
    note = ' <span style="font-size:10px;color:#f59e0b">⚠锁定</span>' if g == '普陀' else ''
    return f'''<tr>
  <td><b>{g}</b>{note}</td>
  <td>{role_badge(role)}</td>
  <td class="num">{fmt(m['target'])}</td>
  <td class="num"><b>{fmt(m['done'])}</b></td>
  {rate_td(m['rate'])}
  <td class="num">{chg_html(m['done_chg'])}</td>
  <td class="num">{fmt(m['refund'])}</td>
  {refund_td(m['refund'], m['done'])}
  <td class="num">{chg_html(m['refund_chg'])}</td>
  <td class="num"><b>{fmt(m['net'])}</b></td>
  <td class="num">{fmt(m['prev_net'])}</td>
  <td class="num">{chg_html(m['net_chg'])}</td>
</tr>'''

# Build person rows (Level 3)
def person_row(name, group_label, tgt, m, person_type='sales_advisor'):
    note = ' <span style="font-size:10px;color:#f59e0b">⚠锁定</span>' if m.get('is_putuo') else ''
    return f'''<tr>
  <td>{group_label}</td>
  <td>{name}{note}</td>
  <td class="num">{fmt(tgt)}</td>
  <td class="num"><b>{fmt(m['done'])}</b></td>
  {rate_td(m['rate'])}
  <td class="num">{chg_html(m['done_chg'])}</td>
  <td class="num">{fmt(m['refund'])}</td>
  {refund_td(m['refund'], m['done'], person_type)}
  <td class="num">{chg_html(m['refund_chg'])}</td>
  <td class="num"><b>{fmt(m['net'])}</b></td>
  <td class="num">{fmt(m['prev_net'])}</td>
  <td class="num">{chg_html(m['net_chg'])}</td>
</tr>'''

# ==================== 生成 HTML ====================
html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>南区7月目标进展看板</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;background:#f1f5f9;color:#1e293b;padding:24px}}
.container{{max-width:1500px;margin:0 auto}}
h1{{font-size:22px;margin-bottom:4px}}
.subtitle{{color:#64748b;font-size:13px;margin-bottom:20px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(170px,1fr));gap:12px;margin-bottom:24px}}
.kpi-card{{background:#fff;border-radius:12px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}}
.kpi-card .label{{font-size:11px;color:#64748b;margin-bottom:2px}}
.kpi-card .value{{font-size:24px;font-weight:700}}
.kpi-card .sub{{font-size:11px;color:#94a3b8;margin-top:2px}}
.tabs{{display:flex;gap:4px;margin-bottom:0}}
.tab{{padding:8px 20px;border-radius:8px 8px 0 0;font-size:13px;font-weight:600;cursor:pointer;border:none;background:#e2e8f0;color:#64748b}}
.tab.active{{background:#fff;color:#1e293b}}
.tab-content{{display:none;background:#fff;border-radius:0 12px 12px;padding:20px;box-shadow:0 1px 3px rgba(0,0,0,0.06);overflow-x:auto}}
.tab-content.active{{display:block}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{text-align:left;padding:10px 12px;border-bottom:2px solid #e2e8f0;color:#64748b;font-weight:600;font-size:11px;white-space:nowrap}}
td{{padding:7px 12px;border-bottom:1px solid #f1f5f9}}
tr:hover td{{background:#f8fafc}}
tr.section-header:hover td{{background:inherit!important}}
.num{{text-align:right;font-variant-numeric:tabular-nums}}
.subtotal td{{background:#f8fafc!important;font-weight:700;border-top:1px solid #e2e8f0}}
.footer{{text-align:center;color:#94a3b8;font-size:11px;margin-top:20px}}
.section-title{{font-size:14px;font-weight:700;color:#1e293b;padding:12px 0 8px;border-bottom:1px solid #e2e8f0;margin-bottom:12px}}
@media(max-width:768px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
</style>
</head>
<body>
<div class="container">
<div style="display:flex;justify-content:space-between;align-items:flex-start">
<div><h1>📊 南区7月目标进展看板</h1>
<p class="subtitle">数据截止：2026年7月10日 | 环比基期：6月1-10日同口径 | 普陀业绩锁定目标表</p></div>
<div style="display:flex;gap:8px;align-items:center;margin-top:4px">
<button onclick="exportBaseTable()" style="padding:8px 18px;background:#fff;color:#1e293b;border:1px solid #d1d5db;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">📥 导出底表</button>
<button onclick="snapshot()" style="padding:8px 18px;background:#1e293b;color:#fff;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">📸 快照</button>
<label style="padding:8px 18px;background:#1e293b;color:#fff;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">
  📤 导入底表 <input type="file" onchange="importBaseTable(this)" accept=".xlsx,.xls" style="display:none">
</label>
<button onclick="exportPutuoTemplate()" style="padding:8px 18px;background:#fff;color:#b45309;border:1px solid #f59e0b;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">📥 普陀模板</button>
<label style="padding:8px 18px;background:#fef3c7;color:#b45309;border:1px solid #f59e0b;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">
  ⚠ 导入普陀 <input type="file" onchange="importPutuo(this)" accept=".xlsx,.xls" style="display:none">
</label>
</div>
</div>

<!-- KPI Cards -->
<div class="kpi-grid">
<div class="kpi-card">
  <div class="label">📌 月总目标</div>
  <div class="value">{fmt(total_metrics['target'])}</div>
  <div class="sub">线上 {fmt(section_tgts['线上'])} + 线下 {fmt(section_tgts['线下门店'])}</div>
</div>
<div class="kpi-card">
  <div class="label">✅ 累计完成</div>
  <div class="value">{fmt(total_metrics['done'])}</div>
  {rate_bar_with_trend(total_metrics['rate'])}
  <div class="sub">上月同期 {fmt(total_metrics['prev_done'])} | 环比 {pct1(total_metrics['done_chg'])}</div>
</div>
</div>

<!-- Legend -->
<div style="display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin-bottom:16px;padding:10px 16px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,0.04);font-size:12px">
  <span style="font-weight:700;color:#1e293b">📐 图例：</span>
  <div style="display:flex;align-items:center;gap:6px">
    <span style="display:inline-block;width:20px;height:12px;background:#22c55e;border-radius:6px"></span>
    <span style="color:#166534;font-weight:600">🟢 正常</span>
    <span style="color:#64748b">完成率 ≥ 时间进度({pct2(TIME_PROGRESS)})，节奏正常可达成</span>
  </div>
  <div style="display:flex;align-items:center;gap:6px">
    <span style="display:inline-block;width:20px;height:12px;background:#f59e0b;border-radius:6px"></span>
    <span style="color:#b45309;font-weight:600">⚠️ 关注</span>
    <span style="color:#64748b">完成率 ≥ {pct2(TIME_PROGRESS*0.5)} 但落后时间进度，按此速度可能完不成</span>
  </div>
  <div style="display:flex;align-items:center;gap:6px">
    <span style="display:inline-block;width:20px;height:12px;background:#dc2626;border-radius:6px"></span>
    <span style="color:#b91c1c;font-weight:600">🚫 高危</span>
    <span style="color:#64748b">完成率 &lt; {pct2(TIME_PROGRESS*0.5)}，严重落后，需特别特别关注</span>
  </div>
  <span style="color:#94a3b8;margin-left:auto">|</span>
  <div style="display:flex;align-items:center;gap:6px">
    <span style="color:#dc2626;font-weight:700">退费标红</span>
    <span style="color:#64748b">销售/班主任 &gt;30% | 门店CM &gt;10%</span>
  </div>
</div>

<!-- Tabs -->
<div class="tabs">
<button class="tab active" data-tab="role" onclick="switchTab('role')">🏢 按角色</button>
<button class="tab" data-tab="group" onclick="switchTab('group')">📋 按分组</button>
<button class="tab" data-tab="online_person" onclick="switchTab('online_person')">👤 线上销售</button>
<button class="tab" data-tab="cm_person" onclick="switchTab('cm_person')">🏥 线下个案</button>
</div>

<!-- Tab 1: 按角色 -->
<div id="tab-role" class="tab-content active">
<table>
<thead><tr>
  <th>角色</th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">本月净流水</th>
</tr></thead><tbody>
{''.join(role_row(r, role_metrics[r]) for r in ['销售','班主任','门店'])}
<tr class="subtotal"><td>南区合计</td>
  <td class="num">{fmt(total_metrics['target'])}</td><td class="num">{fmt(total_metrics['done'])}</td>{rate_td(total_metrics['rate'])}
  <td class="num">{fmt(total_metrics['refund'])}</td>{refund_td(total_metrics['refund'], total_metrics['done'])}
  <td class="num">{fmt(total_metrics['net'])}</td>
</tr>
</tbody></table>
</div>

<!-- Tab 2: 按分组 -->
<div id="tab-group" class="tab-content">
<table>
<thead><tr>
  <th>分组</th><th>角色</th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th><th class="num">完成环比</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">退费环比</th>
  <th class="num">本月净流水</th><th class="num">上月同期净流水</th><th class="num">净流水环比</th>
</tr></thead><tbody>
'''

# Group rows — 线上组
html += '<tr class="section-header"><td colspan="12" style="background:#dbeafe;color:#1e40af;font-weight:700;font-size:13px;padding:10px 12px;border-bottom:2px solid #93c5fd">📡 线上组</td></tr>'
for g in ONLINE_ORDER:
    m = group_metrics.get(g)
    if m is None or m['target'] == 0: continue
    role = classify_role(orders[orders['target_group']==g]['组织'].iloc[0]) if len(orders[orders['target_group']==g]) > 0 else '其他'
    html += group_row(g, m, role)

# Group rows — 线下门店
html += '<tr class="section-header"><td colspan="12" style="background:#d1fae5;color:#065f46;font-weight:700;font-size:13px;padding:10px 12px;border-bottom:2px solid #6ee7b7">🏥 线下门店</td></tr>'
for g in OFFLINE_ORDER:
    m = group_metrics.get(g)
    if m is None or m['target'] == 0: continue
    role = classify_role(orders[orders['target_group']==g]['组织'].iloc[0]) if len(orders[orders['target_group']==g]) > 0 else '其他'
    html += group_row(g, m, role)

html += '''</tbody></table></div>'''

# Tab 3: 线上个人
html += '''<div id="tab-online_person" class="tab-content">
<table>
<thead><tr>
  <th>分组</th><th>销售</th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th><th class="num">完成环比</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">退费环比</th>
  <th class="num">本月净流水</th><th class="num">上月同期净流水</th><th class="num">净流水环比</th>
</tr></thead><tbody>
'''

person_online_sorted = person_online_df.sort_values(['group','rate'], ascending=[True,True])
last_group = ''
group_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0}
for _, r in person_online_sorted.iterrows():
    if r['group'] != last_group:
        if last_group and group_sub['target'] > 0:
            sub_done = group_sub['done']; sub_ref = group_sub['refund']
            sub_net = group_sub['net']; sub_prev_net = group_sub['prev_net']
            html += f'''<tr class="subtotal"><td>{last_group}</td><td>小计</td>
  <td class="num">{fmt(group_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,group_sub['target']))}
  <td class="num">{chg_html(safe_div(sub_done-group_sub['prev_done'],group_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{chg_html(safe_div(sub_ref-group_sub['prev_refund'],group_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td></tr>'''
        last_group = r['group']
        group_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0}
    group_sub['target'] += r['target']
    group_sub['done'] += r['done']
    group_sub['refund'] += r['refund']
    group_sub['prev_done'] += r['prev_done']
    group_sub['prev_refund'] += r['prev_refund']
    group_sub['net'] += r['net']
    group_sub['prev_net'] += r['prev_net']
    html += person_row(r['name'], r['group'], r['target'], r, 'sales_advisor')
# last group subtotal
if last_group and group_sub['target'] > 0:
    sub_done = group_sub['done']; sub_ref = group_sub['refund']
    sub_net = group_sub['net']; sub_prev_net = group_sub['prev_net']
    html += f'''<tr class="subtotal"><td>{last_group}</td><td>小计</td>
  <td class="num">{fmt(group_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,group_sub['target']))}</td>
  <td class="num">{chg_html(safe_div(sub_done-group_sub['prev_done'],group_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{chg_html(safe_div(sub_ref-group_sub['prev_refund'],group_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td></tr>'''

html += '''</tbody></table></div>'''

# Tab 4: 线下个案
html += '''<div id="tab-cm_person" class="tab-content">
<table>
<thead><tr>
  <th>分院</th><th>个案经理</th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th><th class="num">完成环比</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">退费环比</th>
  <th class="num">本月净流水</th><th class="num">上月同期净流水</th><th class="num">净流水环比</th>
</tr></thead><tbody>
'''

cm_sorted = person_cm_df.sort_values(['store','rate'], ascending=[True,True])
last_store = ''
store_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0}
for _, r in cm_sorted.iterrows():
    if r['store'] != last_store:
        if last_store and store_sub['target'] > 0:
            sub_done = store_sub['done']; sub_ref = store_sub['refund']
            sub_net = store_sub['net']; sub_prev_net = store_sub['prev_net']
            html += f'''<tr class="subtotal"><td><b>{last_store}</b></td><td>小计</td>
  <td class="num">{fmt(store_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,store_sub['target']))}</td>
  <td class="num">{chg_html(safe_div(sub_done-store_sub['prev_done'],store_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{chg_html(safe_div(sub_ref-store_sub['prev_refund'],store_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td></tr>'''
        last_store = r['store']
        store_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0}
    store_sub['target'] += r['target']
    store_sub['done'] += r['done']
    store_sub['refund'] += r['refund']
    store_sub['prev_done'] += r['prev_done']
    store_sub['prev_refund'] += r['prev_refund']
    store_sub['net'] += r['net']
    store_sub['prev_net'] += r['prev_net']
    html += person_row(r['name'], r['store'], r['target'], r, 'cm')
# last store subtotal
if last_store and store_sub['target'] > 0:
    sub_done = store_sub['done']; sub_ref = store_sub['refund']
    sub_net = store_sub['net']; sub_prev_net = store_sub['prev_net']
    html += f'''<tr class="subtotal"><td><b>{last_store}</b></td><td>小计</td>
  <td class="num">{fmt(store_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,store_sub['target']))}</td>
  <td class="num">{chg_html(safe_div(sub_done-store_sub['prev_done'],store_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{chg_html(safe_div(sub_ref-store_sub['prev_refund'],store_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td></tr>'''

html += '''</tbody></table></div>'''

html += f'''
<div class="footer">
  南区7月目标进展看板 | 底表数据截止 2026.07.10 | 环比 = 7月1-10日 vs 6月1-10日同口径 | 普陀业绩锁定目标表
</div>
</div>
<script>
function switchTab(name){{
  document.querySelectorAll(".tab").forEach(function(t){{t.classList.remove("active")}});
  document.querySelectorAll(".tab-content").forEach(function(t){{t.classList.remove("active")}});
  document.querySelector('.tab[data-tab="'+name+'"]').classList.add("active");
  document.getElementById("tab-"+name).classList.add("active");
}}
</script>
</body></html>'''

# ==================== 导出 Excel ====================
def to_pct(v):
    return f'{v*100:.1f}%' if pd.notna(v) and v != float('inf') else ''

def to_pct_signed(v):
    if pd.isna(v) or v == float('inf') or v == -float('inf'): return ''
    return f'{v*100:+.1f}%'

# 1) 按角色
role_export = []
for role in ['销售','班主任','门店']:
    m = role_metrics[role]
    role_export.append({
        '角色': role, '目标': m['target'], '完成': m['done'],
        '完成率': to_pct(m['rate']),
        '本月退费': m['refund'], '退费率': to_pct(safe_div(m['refund'], m['done'])),
        '上月完成': m['prev_done'], '完成环比': to_pct_signed(m['done_chg']),
        '本月净流水': m['net'], '上月净流水': m['prev_net'], '净流水环比': to_pct_signed(m['net_chg']),
    })
role_export.append({
    '角色': '南区合计', '目标': total_metrics['target'], '完成': total_metrics['done'],
    '完成率': to_pct(total_metrics['rate']),
    '本月退费': total_metrics['refund'], '退费率': to_pct(safe_div(total_metrics['refund'], total_metrics['done'])),
    '上月完成': total_metrics['prev_done'], '完成环比': to_pct_signed(total_metrics['done_chg']),
    '本月净流水': total_metrics['net'], '上月净流水': total_metrics['prev_net'], '净流水环比': to_pct_signed(total_metrics['net_chg']),
})
pd.DataFrame(role_export).to_excel(r'C:\Users\fanny\Desktop\数据日日\导出_按角色.xlsx', index=False)

# 2) 按分组
group_export = []
for g in ONLINE_ORDER + OFFLINE_ORDER:
    m = group_metrics.get(g)
    if m is None: continue
    grp_orders = orders[orders['target_group']==g]
    role = classify_role(grp_orders['组织'].iloc[0]) if len(grp_orders)>0 else ''
    group_export.append({
        '分组': g, '角色': role, '目标': m['target'], '完成': m['done'],
        '完成率': to_pct(m['rate']), '完成环比': to_pct_signed(m['done_chg']),
        '本月退费': m['refund'], '退费率': to_pct(safe_div(m['refund'], m['done'])),
        '退费环比': to_pct_signed(m['refund_chg']),
        '本月净流水': m['net'], '上月净流水': m['prev_net'], '净流水环比': to_pct_signed(m['net_chg']),
    })
pd.DataFrame(group_export).to_excel(r'C:\Users\fanny\Desktop\数据日日\导出_按分组.xlsx', index=False)

# 3) 线上个人
po_export = []
for _, r in person_online_df.sort_values(['group','rate'], ascending=[True,True]).iterrows():
    po_export.append({
        '分组': r['group'], '销售': r['name'], '目标': r['target'], '完成': r['done'],
        '完成率': to_pct(r['rate']), '完成环比': to_pct_signed(r['done_chg']),
        '本月退费': r['refund'], '退费率': to_pct(safe_div(r['refund'], r['done'])),
        '退费环比': to_pct_signed(r['refund_chg']),
        '本月净流水': r['net'], '上月净流水': r['prev_net'], '净流水环比': to_pct_signed(r['net_chg']),
    })
pd.DataFrame(po_export).to_excel(r'C:\Users\fanny\Desktop\数据日日\导出_线上个人.xlsx', index=False)

# 4) 线下个案
cm_export = []
for _, r in person_cm_df.sort_values(['store','rate'], ascending=[True,True]).iterrows():
    cm_export.append({
        '分院': r['store'], '个案经理': r['name'], '目标': r['target'], '完成': r['done'],
        '完成率': to_pct(r['rate']), '完成环比': to_pct_signed(r['done_chg']),
        '本月退费': r['refund'], '退费率': to_pct(safe_div(r['refund'], r['done'])),
        '退费环比': to_pct_signed(r['refund_chg']),
        '本月净流水': r['net'], '上月净流水': r['prev_net'], '净流水环比': to_pct_signed(r['net_chg']),
    })
pd.DataFrame(cm_export).to_excel(r'C:\Users\fanny\Desktop\数据日日\导出_线下个案.xlsx', index=False)

# ==================== 嵌入映射数据 + 个人日报为 JSON ====================
import json

# 聚合个人每日业绩 (用于TOP10日期筛选)
july_orders = orders[(orders['pay_at'] >= JULY_START) & (orders['pay_at'] < JULY_END)]
person_daily = {}  # {name: {day: amount, ...}, ...}
for _, row in july_orders.iterrows():
    name = row['seller_name']
    if pd.isna(name) or not name: continue
    day = row['pay_at'].day
    amt = float(row['pay_amount'])
    if name not in person_daily: person_daily[name] = {}
    person_daily[name][str(day)] = person_daily[name].get(str(day), 0) + amt
    person_daily[name]['_total'] = person_daily[name].get('_total', 0) + amt
    person_daily[name]['_org'] = row['组织']

mapping_json = json.dumps({
    'org_to_group': ORG_TO_GROUP,
    'store_to_group': STORE_MAP,
    'sales_orgs': SALES_ORGS,
    'advisor_orgs': ADVISOR_ORGS,
    'target_map': target_map,
    'putuo_locked': putuo_locked_done,
    'online_person_target': {f'{k[0]}|{k[1]}': v for k, v in online_person_target.items()},
    'cm_list': cm_list,
    'time_progress': TIME_PROGRESS,
    'july_start': JULY_START, 'july_end': JULY_END,
    'june_start': JUNE_START, 'june_end': JUNE_END,
    'person_daily': person_daily,
}, ensure_ascii=False)

# ==================== 交互式 JS 工具条 ====================
import_toolbar = '''
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<script>
window._M = ''' + mapping_json + ''';
window._TP = window._M.time_progress;
</script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
.tab-toolbar { display:flex; justify-content:flex-end; gap:6px; margin-bottom:8px; }
.tab-toolbar button { padding:5px 14px; border-radius:6px; font-size:12px; cursor:pointer; border:1px solid #d1d5db; background:#fff; color:#4b5563; font-weight:500; }
.tab-toolbar button:hover { background:#f3f4f6; }
td[contenteditable="true"] { outline:2px solid transparent; border-radius:4px; cursor:text; }
td[contenteditable="true"]:hover { outline-color:#93c5fd; background:#eff6ff; }
td[contenteditable="true"]:focus { outline-color:#3b82f6; background:#fff; }
</style>
<script>
// Recalculate subtotals & KPIs after edit
function recalcAll() {
  // Role tab
  let roleRows = document.querySelectorAll('#tab-role tbody tr:not(.subtotal)');
  let roleTotal = {target:0, done:0, refund:0, prev_done:0, net:0};
  roleRows.forEach(row => {
    let cells = row.querySelectorAll('td.num');
    roleTotal.target += parseFloat(cells[0]?.textContent.replace(/,/g,'')) || 0;
    roleTotal.done += parseFloat(cells[1]?.textContent.replace(/,/g,'')) || 0;
    roleTotal.refund += parseFloat(cells[3]?.textContent.replace(/,/g,'')) || 0;
    roleTotal.net += parseFloat(cells[5]?.textContent.replace(/,/g,'')) || 0;
  });
  // update role subtotal row
  let st = document.querySelector('#tab-role tr.subtotal');
  if(st) {
    st.cells[1].textContent = fmtNum(roleTotal.target);
    st.cells[2].textContent = fmtNum(roleTotal.done);
    st.cells[3].textContent = fmtNum(roleTotal.refund);
    st.cells[5].textContent = fmtNum(roleTotal.net);
  }
}

function fmtNum(n) { return n >= 1e4 ? Math.round(n/1e4)+'万' : n.toLocaleString('en-US'); }

// Column sorting - click any header to sort asc/desc
(function(){
  function parseVal(td) {
    var t = td.textContent.replace(/[,，↑↓%\\s]/g,'');
    var mul = 1;
    if(t.indexOf('亿')>=0){ mul=100000000; t=t.replace('亿',''); }
    else if(t.indexOf('万')>=0){ mul=10000; t=t.replace('万',''); }
    var n = parseFloat(t);
    return isNaN(n) ? t : n * mul;
  }
  document.querySelectorAll('.tab-content table thead th').forEach(function(th, idx) {
    th.style.cursor = 'pointer';
    th.title = '点击升序/降序';
    th.addEventListener('click', function() {
      var tbody = th.closest('table').querySelector('tbody');
      var rows = Array.from(tbody.querySelectorAll('tr'));
      var dataRows = rows.filter(function(r){return !r.classList.contains('subtotal') && !r.classList.contains('section-header')});
      var subtotalRows = rows.filter(function(r){return r.classList.contains('subtotal')});
      var isAsc = th.getAttribute('data-sort') !== 'asc';
      th.setAttribute('data-sort', isAsc ? 'asc' : 'desc');
      th.closest('table').querySelectorAll('th').forEach(function(h){
        if(h!==th) h.removeAttribute('data-sort');
        h.textContent = h.textContent.replace(/ [↑↓]$/,'');
      });
      dataRows.sort(function(a,b){
        var va = parseVal(a.cells[idx]), vb = parseVal(b.cells[idx]);
        if(typeof va==='number' && typeof vb==='number')
          return isAsc ? va-vb : vb-va;
        return isAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
      });
      dataRows.forEach(function(r){tbody.appendChild(r)});
      subtotalRows.forEach(function(r){tbody.appendChild(r)});
      var arrow = isAsc ? ' ↑' : ' ↓';
      th.textContent = th.textContent.replace(/ [↑↓]$/,'') + arrow;
    });
  });
})();

// Make cells editable on double-click
document.addEventListener('dblclick', function(e) {
  let td = e.target.closest('td.num');
  if(!td || td.isContentEditable) return;
  td.contentEditable = 'true';
  td.focus();
  // Select all text
  let range = document.createRange();
  range.selectNodeContents(td);
  let sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
});
document.addEventListener('blur', function(e) {
  let td = e.target.closest('td.num[contenteditable="true"]');
  if(!td) return;
  let raw = td.textContent.trim().replace(/[,，万]/g,'');
  if(raw.includes('亿')) raw = raw.replace('亿','') * 1e8;
  let val = parseFloat(raw);
  if(!isNaN(val)) {
    td.textContent = fmtNum(val);
    td.setAttribute('data-val', val);
    recalcAll();
  }
  td.contentEditable = 'false';
}, true);

// ====== full refresh ======
window._TP = ''' + str(TIME_PROGRESS) + ''';
function parseNum(td){ if(!td) return 0; var raw=String(td.getAttribute("data-val")||td.textContent).replace(/[,，\\s↑↓%]/g,""); var mul=1; if(raw.indexOf("亿")>=0){ mul=100000000; raw=raw.replace("亿",""); } else if(raw.indexOf("万")>=0){ mul=10000; raw=raw.replace("万",""); } return (parseFloat(raw)||0)*mul; }
function writeNum(td,v){ if(!td) return; td.setAttribute("data-val",v); td.textContent=fmtNum(v); }
function rebuildBar(td,rate){
  if(!td||isNaN(rate)) return; rate=Math.min(1,Math.max(0,rate)); var p=rate*100,ratio=rate/window._TP;
  var color,bg,barBg,badge="";
  if(ratio>=1){color="#22c55e";bg="#dcfce7";barBg="#f0fdf4"}
  else if(ratio>=0.5){color="#f59e0b";bg="#fef3c7";barBg="#fffbeb";badge='<span style="font-size:10px;background:#fef3c7;color:#b45309;padding:1px 6px;border-radius:8px;font-weight:700;margin-left:4px">⚠ 关注</span>'}
  else{color="#dc2626";bg="#fecaca";barBg="#fef2f2";badge='<span style="font-size:10px;background:#fee2e2;color:#b91c1c;padding:1px 6px;border-radius:8px;font-weight:700;margin-left:4px">🚫 高危</span>'}
  td.innerHTML='<div style="display:flex;align-items:center;gap:6px;min-width:180px"><div style="flex:1;height:20px;background:'+barBg+';border-radius:10px;overflow:hidden;position:relative;border:1px solid #e2e8f0"><div style="width:'+p+'%;height:100%;background:'+color+';border-radius:10px"></div><div style="position:absolute;top:0;left:0;width:100%;height:100%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#1e293b">'+(rate*100).toFixed(1)+'%</div></div>'+badge+'</div>';
  td.setAttribute("data-rate",rate);
}
function rebuildRefund(td,ref,done,type){
  if(!td) return; var rate=done>0?ref/done:0, str=done>0?(rate*100).toFixed(1)+"%":"—";
  var red=(type==="sales"&&rate>0.3)||(type==="cm"&&rate>0.1);
  td.className="num"; td.innerHTML=red?'<span style="color:#dc2626;font-weight:700">'+str+'</span>':str;
}
function refreshAll(){
  var tds,rTotals={tgt:0,done:0,ref:0,net:0},row,st,g;
  // role tab
  document.querySelectorAll("#tab-role tbody tr:not(.subtotal)").forEach(function(row){tds=row.querySelectorAll("td");if(tds.length<7)return;var t=parseNum(tds[1]),d=parseNum(tds[2]),r=parseNum(tds[4]);rTotals.tgt+=t;rTotals.done+=d;rTotals.ref+=r;rTotals.net+=d-r;writeNum(tds[5],d-r);rebuildBar(tds[3],d/t);rebuildRefund(tds[4].nextElementSibling,r,d,"group")});
  st=document.querySelector("#tab-role tr.subtotal");if(st){tds=st.querySelectorAll("td");writeNum(tds[1],rTotals.tgt);writeNum(tds[2],rTotals.done);rebuildBar(tds[3],rTotals.done/rTotals.tgt);writeNum(tds[4],rTotals.ref);writeNum(tds[5],rTotals.net);rebuildRefund(tds[4].nextElementSibling,rTotals.ref,rTotals.done,"group")}
  // group tab
  document.querySelectorAll("#tab-group tbody tr:not(.subtotal):not(.section-header)").forEach(function(row){tds=row.querySelectorAll("td");if(tds.length<12)return;var t=parseNum(tds[2]),d=parseNum(tds[3]),r=parseNum(tds[6]);writeNum(tds[9],d-r);rebuildBar(tds[4],d/t);rebuildRefund(tds[7],r,d,"group")});
  // person tabs
  ["online_person","cm_person"].forEach(function(tabId){var tab=document.getElementById("tab-"+tabId);if(!tab)return;var grps={};tab.querySelectorAll("tbody tr:not(.subtotal)").forEach(function(row){tds=row.querySelectorAll("td");if(tds.length<12)return;var grp=tds[0].textContent.trim(),t=parseNum(tds[2]),d=parseNum(tds[3]),r=parseNum(tds[6]);writeNum(tds[10],d-r);rebuildBar(tds[4],d/t);rebuildRefund(tds[7],r,d,tabId==="tab-cm_person"?"cm":"sales");if(!grps[grp])grps[grp]={tgt:0,done:0,ref:0,net:0};grps[grp].tgt+=t;grps[grp].done+=d;grps[grp].ref+=r;grps[grp].net+=d-r});tab.querySelectorAll("tbody tr.subtotal").forEach(function(st){tds=st.querySelectorAll("td");g=grps[tds[0].textContent.trim()];if(!g||tds.length<10)return;writeNum(tds[2],g.tgt);writeNum(tds[3],g.done);rebuildBar(tds[4],g.done/g.tgt);writeNum(tds[6],g.ref);rebuildRefund(tds[7],g.ref,g.done,tabId==="tab-cm_person"?"cm":"sales");writeNum(tds[9],g.net)})});
  // KPIs
  var cards=document.querySelectorAll(".kpi-card .value");if(cards[0])cards[0].textContent=fmtNum(rTotals.tgt);if(cards[1]){cards[1].textContent=fmtNum(rTotals.done);rebuildBar(cards[1].nextElementSibling,rTotals.done/rTotals.tgt)}
}
recalcAll = refreshAll;
document.addEventListener("blur", function(e){var td=e.target.closest('td.num[contenteditable="true"]');if(!td)return;var raw=td.textContent.trim().replace(/[,，万]/g,"");if(raw.indexOf("亿")>=0)raw=raw.replace("亿","")*1e8;var v=parseFloat(raw);if(!isNaN(v)){td.setAttribute("data-val",v);td.textContent=fmtNum(v)}td.contentEditable="false";setTimeout(refreshAll,100)},true);

// ====== TOP10 弹窗排行榜 (含日期筛选) ======
function showTop10(tabId, fromDay, toDay) {
  fromDay = fromDay || 1; toDay = toDay || 10;
  var tab = document.getElementById('tab-'+tabId);
  if(!tab) return;
  var title = tabId === 'online_person' ? '线上销售 TOP10' : '线下个案 TOP10';
  var isOnline = tabId === 'online_person';

  // Compute from embedded person_daily data
  var M = window._M;
  var pd = M.person_daily || {};
  var personDone = {};
  Object.keys(pd).forEach(function(name) {
    var info = pd[name];
    var org = info._org || '';
    var g = M.org_to_group[org] || M.store_to_group[org] || '';
    // Filter: online sales only for online_person, store CM only for cm_person
    if(isOnline && M.store_to_group[org]) return; // skip store people in online view
    if(!isOnline && !M.store_to_group[org]) return; // skip online people in CM view
    // Filter by day range
    var done = 0;
    for(var d=fromDay; d<=toDay; d++) {
      done += (info[String(d)] || 0);
    }
    if(done <= 0) return;
    personDone[name] = {name: name, group: g, done: done, target: 0};
  });
  // For CM view: include 普陀 CMs from table (may not be in person_daily)
  if(!isOnline) {
    var cmTab = document.getElementById('tab-cm_person');
    if(cmTab) {
      cmTab.querySelectorAll('tbody tr:not(.subtotal)').forEach(function(row) {
        var tds = row.querySelectorAll('td');
        if(tds.length < 4) return;
        if(tds[0].textContent.trim() !== '普陀') return;
        var name = tds[1].textContent.replace('⚠锁定','').trim();
        var tableVal = parseNum(tds[3]);
        var tgt = parseNum(tds[2]);
        if(tableVal > 0) {
          if(!personDone[name]) {
            personDone[name] = {name: name, group: '普陀', done: 0, target: 0};
          }
          personDone[name].done = tableVal;
          personDone[name].target = tgt;
        }
      });
    }
  }

  // Match targets
  if(isOnline) {
    Object.keys(M.online_person_target).forEach(function(k) {
      var parts = k.split('|'), n = parts[1];
      if(personDone[n]) personDone[n].target = M.online_person_target[k];
    });
  } else {
    M.cm_list.forEach(function(cm) {
      if(personDone[cm.name]) personDone[cm.name].target = cm.target;
    });
  }

  var data = Object.values(personDone).filter(function(d){ return d.done > 0; });
  data.sort(function(a,b){ return b.done - a.done; });
  var top10 = data.slice(0, 10);

  var html = '<div id="top10-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.4);z-index:9999;display:flex;align-items:center;justify-content:center" onclick="if(event.target===this)closeTop10()">'+
    '<div style="background:#fff;border-radius:16px;padding:24px;width:520px;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,0.3)">'+
    '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px">'+
    '<h2 style="margin:0;font-size:20px">🏆 '+title+'</h2>'+
    '<button onclick="closeTop10()" style="border:none;background:none;font-size:22px;cursor:pointer;color:#94a3b8">&times;</button>'+
    '</div>'+
    '<div style="display:flex;align-items:center;gap:8px;margin-bottom:16px;padding:10px 14px;background:#f8fafc;border-radius:10px;font-size:13px">'+
    '<span style="color:#64748b;font-weight:600">📅 7月</span>'+
    '<input id="top10-from" type="number" min="1" max="31" value="'+fromDay+'" style="width:50px;padding:4px 8px;border:1px solid #d1d5db;border-radius:6px;text-align:center;font-size:13px">'+
    '<span style="color:#64748b">日 —</span>'+
    '<input id="top10-to" type="number" min="1" max="31" value="'+toDay+'" style="width:50px;padding:4px 8px;border:1px solid #d1d5db;border-radius:6px;text-align:center;font-size:13px">'+
    '<span style="color:#64748b">日</span>'+
    '<button id="top10-filter-btn" style="padding:4px 12px;background:#1e293b;color:#fff;border:none;border-radius:6px;cursor:pointer;font-size:13px;font-weight:600">筛选</button>'+
    '</div>'+
    '<div style="display:flex;flex-direction:column;gap:8px">';
  var crowns = ['🥇','🥈','🥉'];
  var colors = ['#fef3c7','#f1f5f9','#ffedd5'];
  top10.forEach(function(p, i) {
    var rank = i<3 ? '<span style="font-size:22px">'+crowns[i]+'</span>' : '<span style="width:28px;text-align:center;color:#94a3b8;font-weight:700">'+(i+1)+'</span>';
    var bg = i<3 ? colors[i] : '#fff';
    var rate = p.target>0 ? (p.done/p.target*100).toFixed(1)+'%' : '—';
    html += '<div style="display:flex;align-items:center;gap:12px;padding:10px 14px;background:'+bg+';border-radius:10px;border:1px solid #e2e8f0">'+
      '<div style="width:32px;text-align:center">'+rank+'</div>'+
      '<div style="flex:1"><div style="font-weight:700;font-size:14px">'+p.name+'</div><div style="font-size:11px;color:#64748b">'+p.group+'</div></div>'+
      '<div style="text-align:right"><div style="font-weight:700;font-size:16px;color:#1e293b">'+fmtNum(p.done)+'</div><div style="font-size:11px;color:#64748b">目标 '+fmtNum(p.target)+' | 完成率 '+rate+'</div></div>'+
    '</div>';
  });
  if(top10.length === 0) html += '<div style="text-align:center;padding:20px;color:#94a3b8">该时间段暂无数据</div>';
  html += '</div></div></div>';
  var old = document.getElementById('top10-overlay');
  if(old) old.remove();
  document.body.insertAdjacentHTML('beforeend', html);
  setTimeout(function() {
    var btn = document.getElementById('top10-filter-btn');
    if(btn) btn.onclick = function() { applyTop10Filter(tabId); };
  }, 50);
}
function applyTop10Filter(tabId) {
  var f = parseInt(document.getElementById('top10-from').value) || 1;
  var t = parseInt(document.getElementById('top10-to').value) || 10;
  closeTop10();
  showTop10(tabId, f, t);
}
function closeTop10() {
  var overlay = document.getElementById('top10-overlay');
  if(overlay) overlay.remove();
}

// ====== 快照 ======
function snapshot() {
  var w = window.open('', '_blank', 'width=800,height=900');
  var kpis = document.querySelector('.kpi-grid').outerHTML;
  var role = document.querySelector('#tab-role').outerHTML.replace('class="tab-content active"','class="tab-content active" style="display:block"');
  var group = document.querySelector('#tab-group').outerHTML.replace('class="tab-content"','class="tab-content" style="display:block"');
  var legend = document.querySelector('.kpi-grid').nextElementSibling ? document.querySelector('.kpi-grid').nextElementSibling.outerHTML : '';
  w.document.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><title>南区7月目标快照</title>'+
    '<style>'+
    '*{margin:0;padding:0;box-sizing:border-box}'+
    'body{font-family:-apple-system,BlinkMacSystemFont,"PingFang SC","Microsoft YaHei",sans-serif;background:#fff;color:#1e293b;padding:24px;max-width:1100px;margin:0 auto}'+
    'h1{font-size:20px;margin-bottom:4px}'+
    '.subtitle{color:#64748b;font-size:12px;margin-bottom:16px}'+
    '.kpi-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px;margin-bottom:20px}'+
    '.kpi-card{background:#f8fafc;border-radius:12px;padding:14px 18px;border:1px solid #e2e8f0}'+
    '.kpi-card .label{font-size:11px;color:#64748b;margin-bottom:2px}'+
    '.kpi-card .value{font-size:22px;font-weight:700}'+
    '.kpi-card .sub{font-size:10px;color:#94a3b8;margin-top:2px}'+
    'table{width:100%;border-collapse:collapse;font-size:12px;margin-bottom:20px}'+
    'th{text-align:left;padding:8px 10px;border-bottom:2px solid #e2e8f0;color:#64748b;font-weight:600;font-size:11px}'+
    'td{padding:6px 10px;border-bottom:1px solid #f1f5f9}'+
    '.num{text-align:right}'+
    '.subtotal td{background:#f8fafc;font-weight:700;border-top:1px solid #e2e8f0}'+
    '.tab-toolbar{display:none}'+
    '.footer{text-align:center;color:#94a3b8;font-size:10px;margin-top:16px}'+
    '@media print{body{padding:0} .footer{position:fixed;bottom:0}}'+
    '</style></head><body>'+
    '<h1>📊 南区7月目标进展快照</h1>'+
    '<p class="subtitle">数据截止：2026年7月10日 | 环比基期：6月1-10日同口径</p>'+
    kpis + legend +
    '<h2 style="font-size:16px;margin:16px 0 8px">🏢 按角色</h2>'+
    role +
    '<h2 style="font-size:16px;margin:16px 0 8px">📋 按分组</h2>'+
    group +
    '<div class="footer">📸 南区7月目标进展快照 | 观心运营助手</div>'+
    '</body></html>');
  w.document.close();
}

// ====== 导出底表 (订单维度) ======
function exportBaseTable() {
  var raw = window._rawOrders;
  if(raw && raw.length > 0) {
    // Export from imported raw data
    var ws = XLSX.utils.json_to_sheet(raw);
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '底表');
    XLSX.writeFile(wb, '7月收入-青少年体验课订单明细.xlsx');
  } else {
    // No raw data imported yet - download the original file
    var a = document.createElement('a');
    a.href = '7月收入-青少年体验课订单明细.xlsx';
    a.download = '7月收入-青少年体验课订单明细.xlsx';
    a.click();
  }
}

// ====== 普陀模板导出 ======
function exportPutuoTemplate() {
  var M = window._M;
  // Gather 普陀 data from CM table
  var cmTab = document.getElementById('tab-cm_person');
  var rows = cmTab.querySelectorAll('tbody tr:not(.subtotal)');
  var templateRows = [['门店业绩','门店退款','个案经理','完成业绩']];
  var totalRef = 0;
  rows.forEach(function(row) {
    var tds = row.querySelectorAll('td');
    if(tds.length < 4 || tds[0].textContent.trim() !== '普陀') return;
    var name = tds[1].textContent.replace('⚠锁定','').trim();
    var done = parseNum(tds[3]);
    var ref = parseNum(tds[6]);
    totalRef += ref;
    templateRows.push(['', '', name, done]);
  });
  // Add summary rows
  var grpRows = document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)');
  var putuoTotal = 0, putuoRefTotal = 0;
  grpRows.forEach(function(row) {
    var tds = row.querySelectorAll('td');
    if(tds.length < 12) return;
    if(tds[0].textContent.replace('⚠锁定','').trim() !== '普陀') return;
    putuoTotal = parseNum(tds[3]);
    putuoRefTotal = parseNum(tds[6]);
  });
  templateRows.push(['', '', '', '']);
  templateRows.push([putuoTotal, putuoRefTotal, '合计', putuoTotal]);
  var ws = XLSX.utils.aoa_to_sheet(templateRows);
  // Set column widths
  ws['!cols'] = [{wch:14},{wch:14},{wch:12},{wch:14}];
  var wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, '普陀');
  XLSX.writeFile(wb, '普陀门店数据模板.xlsx');
}

// ====== 普陀专属导入 ======
function importPutuo(input) {
  var file = input.files[0]; if(!file) return;
  var reader = new FileReader();
  reader.onload = function(e) {
    var wb = XLSX.read(e.target.result, {type:'array'});
    var sheet = wb.Sheets[wb.SheetNames[0]];
    // Try both object mode and array mode
    var data = XLSX.utils.sheet_to_json(sheet);
    console.log('普陀导入: 原始数据行数=' + data.length, data.slice(0,3));
    // Detect columns from first row
    var cols = data.length > 0 ? Object.keys(data[0]) : [];
    console.log('普陀导入: 列名=' + cols.join(', '));
    // Map columns flexibly
    var nameCol = cols.find(function(c){ return c.indexOf('个案')>=0 || c.indexOf('经理')>=0 || c.indexOf('姓名')>=0 || c === 'name'; }) || cols[2] || '';
    var doneCol = cols.find(function(c){ return c.indexOf('完成')>=0 || c.indexOf('业绩')>=0 || c === 'done'; }) || cols[3] || '';
    var revCol = cols.find(function(c){ return c.indexOf('门店业绩')>=0 || c.indexOf('业绩')>=0; }) || cols[0] || '';
    var refCol = cols.find(function(c){ return c.indexOf('退款')>=0 || c.indexOf('退费')>=0; }) || cols[1] || '';
    console.log('普陀导入: nameCol='+nameCol+' doneCol='+doneCol+' revCol='+revCol+' refCol='+refCol);

    var putuoDone = {};  // name -> done
    var totalPutuo = 0, storeRefund = null;
    data.forEach(function(row) {
      var name = (row[nameCol] || '').toString().trim();
      var done = parseFloat(row[doneCol]) || 0;
      var storeRev = parseFloat(row[revCol]) || 0;
      var storeRef = parseFloat(row[refCol]) || 0;
      if(storeRev > 0 && !name) totalPutuo = storeRev;
      if(storeRef > 0 && !name) storeRefund = storeRef;
      if(name && name !== '合计' && done > 0) {
        putuoDone[name] = done;
        console.log('普陀导入: 个案 '+name+' = '+done);
      }
    });
    // Merge summary row values
    if(totalPutuo === 0) {
      Object.values(putuoDone).forEach(function(v) { totalPutuo += v; });
    }
    console.log('普陀导入: totalPutuo='+totalPutuo+' 个案数='+Object.keys(putuoDone).length);
    if(totalPutuo === 0 && Object.keys(putuoDone).length === 0) {
      alert('未识别到普陀数据。列名: '+cols.join(',')+' -- 请确保表格包含个案经理和完成业绩列'); return;
    }

    // Update CM person table - 普陀 rows
    var cmTab = document.getElementById('tab-cm_person');
    if(!cmTab) { console.log('普陀导入: CM tab not found!'); return; }
    var cmRows = cmTab.querySelectorAll('tbody tr:not(.subtotal)');
    console.log('普陀导入: CM表行数=' + cmRows.length);
    var cmGroupSub = {tgt:0, done:0, ref:0, net:0};
    var updatedCount = 0;
    cmRows.forEach(function(row) {
      var tds = row.querySelectorAll('td');
      if(tds.length < 4) return;
      var store = tds[0].textContent.trim();
      if(store !== '普陀') return;
      var name = tds[1].textContent.replace('⚠锁定','').trim();
      var newDone = putuoDone[name];
      if(newDone !== undefined) {
        console.log('普陀导入: 更新CM '+name+' 从 '+parseNum(tds[3])+' → '+newDone);
        writeNum(tds[3], newDone);
        tds[1].innerHTML = name + ' <span style="font-size:10px;color:#f59e0b">⚠锁定</span>';
        var tgt = parseNum(tds[2]);
        rebuildBar(tds[4], newDone/tgt);
        var ref = parseNum(tds[6]);
        writeNum(tds[10], newDone - ref);
        rebuildRefund(tds[7], ref, newDone, 'cm');
        updatedCount++;
      }
      cmGroupSub.tgt += parseNum(tds[2]);
      cmGroupSub.done += parseNum(tds[3]);
      cmGroupSub.ref += parseNum(tds[6]);
      cmGroupSub.net += parseNum(tds[10]);
    });
    console.log('普陀导入: 更新了'+updatedCount+'条CM记录');

    // Update CM subtotal for 普陀
    cmTab.querySelectorAll('tbody tr.subtotal').forEach(function(st) {
      var tds = st.querySelectorAll('td');
      if(tds[0].textContent.trim() !== '普陀') return;
      writeNum(tds[2], cmGroupSub.tgt);
      writeNum(tds[3], cmGroupSub.done);
      rebuildBar(tds[4], cmGroupSub.done/cmGroupSub.tgt);
      writeNum(tds[6], cmGroupSub.ref);
      rebuildRefund(tds[7], cmGroupSub.ref, cmGroupSub.done, 'cm');
      writeNum(tds[9], cmGroupSub.net);
    });

    // Update group table - 普陀 row
    var grpRows = document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)');
    grpRows.forEach(function(row) {
      var tds = row.querySelectorAll('td');
      if(tds.length < 12) return;
      if(tds[0].textContent.replace('⚠锁定','').trim() !== '普陀') return;
      writeNum(tds[3], totalPutuo);
      var tgt = parseNum(tds[2]);
      rebuildBar(tds[4], totalPutuo/tgt);
      if(storeRefund !== null) { writeNum(tds[6], storeRefund); }
      var ref = parseNum(tds[6]);
      writeNum(tds[9], totalPutuo - ref);
      rebuildRefund(tds[7], ref, totalPutuo, 'group');
    });

    // Update window._M.putuo_locked for底表导入 consistency
    window._M.putuo_locked = totalPutuo;

    // Full refresh cascades to role table, KPIs, etc
    refreshAll();
    alert('✅ 普陀数据已更新！总业绩: '+fmtNum(totalPutuo)+'，个案数: '+Object.keys(putuoDone).length);
  };
  reader.readAsArrayBuffer(file);
}

// ====== 底表导入 & 全量重算 ======
function importBaseTable(input) {
  var file = input.files[0]; if(!file) return;
  var reader = new FileReader();
  reader.onload = function(e) {
    var wb = XLSX.read(e.target.result, {type:'array'});
    var sheet = wb.Sheets[wb.SheetNames[0]];
    var data = XLSX.utils.sheet_to_json(sheet);
    var M = window._M;
    var julyS = M.july_start, julyE = M.july_end, juneS = M.june_start, juneE = M.june_end;

    // Store full raw orders for export
    window._rawOrders = data;

    // Rebuild person_daily from imported raw data
    var newPD = {};
    data.forEach(function(row) {
      var payAt = row['pay_at'] ? new Date(row['pay_at']) : null;
      if(!payAt || payAt.getMonth()+1 !== 7) return;
      var name = row['seller_name'] || '';
      if(!name) return;
      var day = payAt.getDate();
      var amt = parseFloat(row['pay_amount']) || 0;
      if(!newPD[name]) { newPD[name] = {}; newPD[name]._total = 0; newPD[name]._org = row['组织'] || ''; }
      newPD[name][String(day)] = (newPD[name][String(day)] || 0) + amt;
      newPD[name]._total += amt;
    });
    window._M.person_daily = newPD;

    // Build group-level metrics from raw data
    var groupCalc = {};
    data.forEach(function(row) {
      var org = row['组织'] || '';
      var g = M.org_to_group[org] || M.store_to_group[org] || null;
      if(!g) return;
      if(!groupCalc[g]) groupCalc[g] = {j7_inc:0, j6_inc:0, j7_ref:0, j6_ref:0};
      var pay = parseFloat(row['pay_amount']) || 0;
      var payAt = row['pay_at'] ? new Date(row['pay_at']) : null;
      var backAt = row['back_at'] ? new Date(row['back_at']) : null;
      if(backAt && backAt.getFullYear() < 2000) backAt = null;
      // Income: by pay_at
      if(payAt && payAt >= new Date(julyS) && payAt < new Date(julyE)) groupCalc[g].j7_inc += pay;
      if(payAt && payAt >= new Date(juneS) && payAt < new Date(juneE)) groupCalc[g].j6_inc += pay;
      // Refund: by back_at
      if(backAt && backAt >= new Date(julyS) && backAt < new Date(julyE)) groupCalc[g].j7_ref += pay;
      if(backAt && backAt >= new Date(juneS) && backAt < new Date(juneE)) groupCalc[g].j6_ref += pay;
    });

    // Update group table
    var grpRows = document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)');
    grpRows.forEach(function(row) {
      var tds = row.querySelectorAll('td');
      if(tds.length < 12) return;
      var gName = tds[0].textContent.replace('⚠锁定','').trim();
      var gc = groupCalc[gName];
      if(!gc) return;
      var tgt = parseNum(tds[2]);
      var done = gName === '普陀' ? M.putuo_locked : gc.j7_inc;
      writeNum(tds[3], done);
      rebuildBar(tds[4], done/tgt);
      // chg
      var chg = gc.j6_inc > 0 ? ((gc.j7_inc - gc.j6_inc)/gc.j6_inc*100) : 0;
      var arrow = chg>=0 ? '↑' : '↓', color = chg>=0 ? '#22c55e' : '#ef4444';
      tds[5].innerHTML = '<span style="color:'+color+';font-weight:600">'+arrow+Math.abs(chg).toFixed(1)+'%</span>';
      writeNum(tds[6], gc.j7_ref);
      rebuildRefund(tds[7], gc.j7_ref, done, 'group');
      tds[9].textContent = fmtNum(done - gc.j7_ref);
    });

    // Update role table from group data
    var roleCalc = {sales:{tgt:0,done:0,ref:0,net:0}, advisor:{tgt:0,done:0,ref:0,net:0}, store:{tgt:0,done:0,ref:0,net:0}};
    var roleMap = {};
    M.sales_orgs.forEach(function(o){ roleMap[M.org_to_group[o]] = 'sales' });
    M.advisor_orgs.forEach(function(o){ roleMap[M.org_to_group[o]] = 'advisor' });
    Object.keys(M.store_to_group).forEach(function(o){ roleMap[M.store_to_group[o]] = 'store' });

    Object.keys(groupCalc).forEach(function(g) {
      var role = roleMap[g] || 'store';
      if(role === 'store' && ['班主任1组','班主任组','线上课程销售组'].indexOf(g) >= 0) role = 'ignore';
      if(role === 'ignore') return;
      var gc = groupCalc[g], tgt = M.target_map[g] || 0;
      var done = g === '普陀' ? M.putuo_locked : gc.j7_inc;
      roleCalc[role].tgt += tgt;
      roleCalc[role].done += done;
      roleCalc[role].ref += gc.j7_ref;
      roleCalc[role].net += done - gc.j7_ref;
    });

    var roleRows = document.querySelectorAll('#tab-role tbody tr:not(.subtotal)');
    var roleOrder = ['sales','advisor','store'];
    var roleNames = {'sales':'销售','advisor':'班主任','store':'门店'};
    roleRows.forEach(function(row, i) {
      var tds = row.querySelectorAll('td');
      var rkey = roleOrder[i];
      if(!rkey || !roleCalc[rkey]) return;
      var rc = roleCalc[rkey];
      writeNum(tds[1], rc.tgt);
      writeNum(tds[2], rc.done);
      rebuildBar(tds[3], rc.done/rc.tgt);
      writeNum(tds[4], rc.ref);
      writeNum(tds[5], rc.net);
      rebuildRefund(tds[4].nextElementSibling, rc.ref, rc.done, 'group');
    });

    // Full refresh for subtotals, KPIs, etc
    refreshAll();
    alert('✅ 底表导入完成！所有数据已刷新');
  };
  reader.readAsArrayBuffer(file);
}
</script>
'''

# Insert toolbar into each tab content div
html = html.replace('<div id="tab-role" class="tab-content active">',
  '<div id="tab-role" class="tab-content active">')
html = html.replace('<div id="tab-group" class="tab-content">',
  '<div id="tab-group" class="tab-content">')
html = html.replace('<div id="tab-online_person" class="tab-content">',
  '<div id="tab-online_person" class="tab-content">\n<div class="tab-toolbar"><button onclick="showTop10(\'online_person\')">🏆 TOP10</button></div>')
html = html.replace('<div id="tab-cm_person" class="tab-content">',
  '<div id="tab-cm_person" class="tab-content">\n<div class="tab-toolbar"><button onclick="showTop10(\'cm_person\')">🏆 TOP10</button></div>')

# Insert JS after styles, before </head>
html = html.replace('</style>\n</head>', '</style>\n' + import_toolbar + '\n</head>')

# Write
out = r'C:\Users\fanny\Desktop\数据日日\南区7月目标看板.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)

print(f'✅ 看板已生成: {out}')
print(f'✅ 导出文件: 导出_按角色.xlsx / 导出_按分组.xlsx / 导出_线上个人.xlsx / 导出_线下个案.xlsx')
print(f'总目标: {fmt(total_metrics["target"])}')
print(f'完成: {fmt(total_metrics["done"])} ({pct2(total_metrics["rate"])})')
print(f'退费: {fmt(total_metrics["refund"])} (退费率 {pct2(safe_div(total_metrics["refund"],total_metrics["done"]))})')
print(f'净流水: {fmt(total_metrics["net"])} (环比 {pct1(total_metrics["net_chg"])})')
