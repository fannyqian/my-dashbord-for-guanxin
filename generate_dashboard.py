#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""南区7月目标进展看板 — 参考业绩进展.xlsx Sheet1结构，数据从底表实时计算"""

import pandas as pd
import numpy as np
from collections import defaultdict

# ==================== 加载数据 ====================
import os, glob

# 自动找每日日文件夹里最新的数据文件（支持 xlsx/csv，排除模板和临时文件）
data_dir = r'C:\Users\fanny\Desktop\数据日日\每日日'
exclude_patterns = ['普陀', '康博嘉', '模板']
candidates = []
for pat in ['*.xlsx', '*.csv']:
    for f in glob.glob(os.path.join(data_dir, pat)):
        bn = os.path.basename(f)
        if bn.startswith('~$'):
            continue
        if any(p in bn for p in exclude_patterns):
            continue
        candidates.append(f)

# 筛选包含订单底表特征列 (order_id) 的文件
data_files = []
for f in candidates:
    try:
        if f.endswith('.csv'):
            sample = pd.read_csv(f, nrows=0, encoding='utf-8')
        else:
            sample = pd.read_excel(f, nrows=0)
        if 'order_id' in sample.columns:
            data_files.append(f)
    except Exception:
        pass
if not data_files:
    raise FileNotFoundError(f'{data_dir} 中没有找到订单底表文件')
latest_file = max(data_files, key=os.path.getmtime)
print(f'📂 数据源: {os.path.basename(latest_file)}')
if latest_file.endswith('.csv'):
    orders = pd.read_csv(latest_file, encoding='utf-8')
else:
    orders = pd.read_excel(latest_file)
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
# 确保日期列是 datetime 类型
orders['pay_at'] = pd.to_datetime(orders['pay_at'])
orders['back_at'] = pd.to_datetime(orders['back_at'], errors='coerce')

# ==================== 从 beisen_org_full_name 提取组织名 ====================
def extract_org(beisen_path):
    """从北森组织全路径提取原组织名"""
    if pd.isna(beisen_path):
        return None
    parts = str(beisen_path).split('_')
    last = parts[-1]
    # 课程销售组/班主任组：末段即为组织名
    if last in ('课程销售1组','课程销售2组','课程销售3组','课程销售4组','课程销售5组','课程销售6组',
                '班主任1组','班主任2组','班主任3组','班主任组','线上课程销售组','销售管理部'):
        return last
    # 门店：在路径中匹配已知门店名
    for part in parts:
        if part in ('上海普陀','上海普陀店'): return '上海普陀店'
        if part == '上海徐汇店': return '上海徐汇店'
        if part == '云南分公司': return '云南分公司'
        if part == '南京分公司': return '南京分公司'
        if part == '广州观心诊所': return '广州观心诊所'
        if part == '成都分公司': return '成都分公司'
        if part in ('杭州分公司','杭州观心诊所'): return '杭州分公司'
        if part == '武汉观心诊所': return '武汉观心诊所'
        if part == '深圳分公司': return '深圳分公司'
        if part == '珠海分公司': return '珠海分公司'
        if part == '福州观心诊所': return '福州观心诊所'
    return None

orders['组织'] = orders['beisen_org_full_name'].apply(extract_org)
# 过滤掉未能识别组织的行
orders = orders.dropna(subset=['组织'])

# === 特殊归属规则：指定销售的底表收入改派到其他门店 ===
# 余佼组织架构在普陀，但底表业绩计入徐汇店（2026-07-17 确认）
SELLER_ORG_OVERRIDE = {'余佼': '上海徐汇店'}
for _nm, _org in SELLER_ORG_OVERRIDE.items():
    orders.loc[orders['seller_name'] == _nm, '组织'] = _org

# ==================== 映射表 ====================
# 参考7.9 sheet角色定义: 销售/班主任(含观心学院)/门店
SALES_ORGS = ['课程销售1组','课程销售2组','课程销售3组','课程销售4组','课程销售5组','课程销售6组']
ADVISOR_ORGS = ['班主任1组','班主任2组','班主任3组','班主任组']  # 全部班主任计入
STORE_MAP = {
    '上海普陀店': '普陀','上海徐汇店': '徐汇','云南分公司': '昆明',
    '南京分公司': '南京','广州观心诊所': '广州','成都分公司': '成都',
    '杭州分公司': '杭州','武汉观心诊所': '武汉','深圳分公司': '深圳',
    '珠海分公司': '珠海','福州观心诊所': '福州',
}
# 线下个案三组划分
STORE_TO_TRIO = {
    '普陀': '一组','徐汇': '一组','南京': '一组','广州': '一组',
    '深圳': '二组','杭州': '二组','武汉': '二组',
    '昆明': '三组','珠海': '三组','成都': '三组','福州': '三组',
}
# 组织→分组名 (匹配目标表)
ORG_TO_GROUP = {
    '课程销售1组':'课程1组','课程销售2组':'课程2组','课程销售3组':'课程3组',
    '课程销售4组':'课程4组','课程销售5组':'课程5组','课程销售6组':'课程6组',
    '班主任1组':'班主任','班主任2组':'班主任','班主任3组':'班主任',
    '班主任组':'班主任',
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
    t = row.iloc[2]  # 7月目标/w
    if g in ('分组','nan','') or g == '合计':
        if label in ('线上','线下门店'): current_section = label
        continue
    if label in ('线上','线下门店'): current_section = label
    try:
        val = float(t)
        target_map[g] = val
        if current_section: section_tgts[current_section] += val
    except: continue

# 对齐分组名：Sheet1 用 销售N组，改为 课程N组
for i in range(1, 7):
    old_k = f'销售{i}组'
    new_k = f'课程{i}组'
    if old_k in target_map:
        target_map[new_k] = target_map.pop(old_k)
# 线上班主任 → 班主任
if '线上班主任' in target_map:
    target_map['班主任'] = target_map.pop('线上班主任')

# === 用户指定目标（覆盖 Sheet1） ===
# 重算分节汇总
section_tgts = {'线上': 0, '线下门店': 0}
ONLINE_TGT_GROUPS = ['课程1组','课程2组','课程3组','课程4组','课程5组','课程6组','班主任']
OFFLINE_TGT_GROUPS = ['普陀','昆明','成都','广州','深圳','珠海','福州','杭州','南京','武汉','徐汇']
for g in ONLINE_TGT_GROUPS:
    section_tgts['线上'] += target_map.get(g, 0)
for g in OFFLINE_TGT_GROUPS:
    section_tgts['线下门店'] += target_map.get(g, 0)
# KPI卡片展示用：线上固定，线下凑整到1510万
section_tgts_display = {
    '线上': section_tgts['线上'],
    '线下门店': 1510_0000 - section_tgts['线上']
}

# 线上个人目标（合并班主任 → 班主任）
online_person_target = {}
for _, row in target_online_person.iterrows():
    grp = str(row['分组']).strip()
    if grp in ('班主任2组', '班主任3组'):
        grp = '班主任'
    online_person_target[(grp, row['销售'])] = float(row['目标'])

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

# ==================== 普陀业绩单独更新（从每日明细自动汇总） ====================
putuo_files = [f for f in
    glob.glob(os.path.join(data_dir, '*普陀*明细*.xlsx')) +
    glob.glob(os.path.join(data_dir, '*普陀*明细*.xls')) +
    glob.glob(os.path.join(data_dir, 'HG0019*.xlsx')) +
    glob.glob(os.path.join(data_dir, 'HG0019*.xls'))
    if not os.path.basename(f).startswith('~$')]
if not putuo_files:
    # 回退：旧版普陀模板文件
    putuo_files = [f for f in glob.glob(os.path.join(data_dir, '*普陀*.xlsx'))
                   if not os.path.basename(f).startswith('~$')]
putuo_file = max(putuo_files, key=os.path.getmtime) if putuo_files else None

putuo_cm_done = {}  # name -> done
putuo_store_perf = 0
putuo_store_refund = 0

if putuo_file:
    # 1. 尝试读取，自动修复 $A$1 坏单元格引用
    try:
        putuo_df = pd.read_excel(putuo_file)
    except Exception as e:
        if '$' in str(e):
            print(f'🔧 普陀文件损坏($引用)，自动修复中...')
            import zipfile, re, io
            fixed_path = putuo_file.replace('.xlsx', '_fixed.xlsx').replace('.xls', '_fixed.xls')
            zin = zipfile.ZipFile(putuo_file)
            zout = zipfile.ZipFile(fixed_path, 'w', zipfile.ZIP_DEFLATED)
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename.startswith('xl/worksheets/') and item.filename.endswith('.xml'):
                    data = re.sub(rb'r="\$([A-Z]+)\$(\d+)"', rb'r="\1\2"', data)
                    data = re.sub(rb'ref="\$([A-Z]+)\$(\d+)', rb'ref="\1\2', data)
                    data = re.sub(rb':\$([A-Z]+)\$(\d+)"', rb':\1\2"', data)
                zout.writestr(item, data)
            zout.close()
            zin.close()
            putuo_df = pd.read_excel(fixed_path)
            os.replace(fixed_path, putuo_file)  # 替换原文件，下次不用再修
            print(f'🔧 普陀文件已修复并覆盖原文件')
        else:
            raise

    # 2. 自动检测表头行：扫描前5行找到含"CM姓名"/"收款金额"的真表头
    if 'CM姓名' not in putuo_df.columns or '收款金额' not in putuo_df.columns:
        raw = pd.read_excel(putuo_file, header=None)
        header_row = None
        for i in range(min(5, len(raw))):
            row_vals = [str(v) for v in raw.iloc[i].values]
            row_text = ' '.join(row_vals)
            if ('CM' in row_text and ('姓名' in row_text or '经理' in row_text)) or '收款金额' in row_text:
                header_row = i
                break
        if header_row is not None:
            putuo_df = raw.iloc[header_row+1:].copy()
            putuo_df.columns = [str(v).strip() for v in raw.iloc[header_row].values]
            putuo_df = putuo_df.reset_index(drop=True)
            # 过滤空行（患者编号为空）
            if '患者编号' in putuo_df.columns:
                putuo_df = putuo_df[putuo_df['患者编号'].notna() & (putuo_df['患者编号'].astype(str).str.strip() != '')]
            # 另存规范版覆盖原文件
            putuo_df.to_excel(putuo_file, index=False)
            print(f'🔧 普陀表头自动检测: 第{header_row+1}行为真表头，已规范化覆盖')

    # 自动识别格式：明细格式有 CM姓名/收款金额 列，模板格式有 个案经理/门店业绩 列
    if 'CM姓名' in putuo_df.columns and '收款金额' in putuo_df.columns:
        # 新格式：原始交易明细 → 按CM汇总
        print(f'📂 普陀明细: {os.path.basename(putuo_file)}')
        for _, row in putuo_df.iterrows():
            cm_name = str(row['CM姓名']).strip() if pd.notna(row['CM姓名']) else ''
            amt = float(row['收款金额']) if pd.notna(row['收款金额']) else 0
            if cm_name and cm_name != 'nan':
                putuo_cm_done[cm_name] = putuo_cm_done.get(cm_name, 0) + amt
        putuo_store_perf = putuo_df['收款金额'].sum()  # 全量（含空CM的医生业绩）
        # 退费从底表单独统计（退款金额列如有则用）
        if '退款金额' in putuo_df.columns:
            putuo_store_refund = putuo_df['退款金额'].sum()
    else:
        # 旧格式：预汇总模板（个案经理/门店业绩/门店退款）
        print(f'📂 普陀模板: {os.path.basename(putuo_file)}')
        for _, row in putuo_df.iterrows():
            name = str(row['个案经理']).strip() if pd.notna(row['个案经理']) else ''
            if name == '合计':
                putuo_store_perf = float(row['门店业绩']) if pd.notna(row['门店业绩']) else 0
                putuo_store_refund = float(row['门店退款']) if pd.notna(row['门店退款']) else 0
            elif name and name != 'nan':
                putuo_cm_done[name] = float(row['完成业绩']) if pd.notna(row['完成业绩']) else 0

# Override 普陀 CM 数据（明细汇总覆盖 done_locked，后面会 + 底表）
for cm in cm_list:
    if cm['store'] == '普陀' and cm['name'] in putuo_cm_done:
        cm['done_locked'] = putuo_cm_done[cm['name']]
putuo_extra = putuo_store_perf  # 明细业绩，稍后叠加到底表
# 重建 cm_df 以应用普陀覆写
cm_df = pd.DataFrame(cm_list)

# ==================== 康博嘉业绩导入 ====================
kangbojia_store_extra = {}  # store_name -> extra_amount
kangbojia_cm_extra = {}     # cm_name -> extra_amount
kangbojia_store_daily = defaultdict(lambda: defaultdict(float))  # store -> day -> amt
kbj_store_map = {
    '观心正德诊所': '广州',
    '武汉观心精神专科门诊部': '武汉',
    '深圳观心正德综合门诊部': '深圳',
    '南京建邺观心综合门诊部': '南京',
    '杭州观心正念诊所有限公司': '杭州',
}
kbj_files = [f for f in
    glob.glob(os.path.join(data_dir, '*康博嘉*.xlsx')) +
    glob.glob(os.path.join(data_dir, '*康博嘉*.xls'))
    if not os.path.basename(f).startswith('~$')]
if kbj_files:
    kbj_file = max(kbj_files, key=os.path.getmtime)
    # 自动选数据最多的 sheet（兼容带标题页的新格式）
    kbj_xls = pd.ExcelFile(kbj_file)
    if len(kbj_xls.sheet_names) > 1:
        best_sheet = max(kbj_xls.sheet_names, key=lambda s: pd.read_excel(kbj_xls, sheet_name=s, nrows=0).shape[1])
    else:
        best_sheet = kbj_xls.sheet_names[0]
    kbj_df = pd.read_excel(kbj_file, sheet_name=best_sheet)
    print(f'📂 康博嘉源: {os.path.basename(kbj_file)} [{best_sheet}]')
    # 检测是否为原始交易数据（列数>10 = 原始明细，<=10 = 预汇总）
    if len(kbj_df.columns) > 10:
        # 原始交易数据：按所在诊所 + 客户经理分别汇总实际收入
        kbj_time_col = kbj_df.columns[0]    # 收费时间
        kbj_store_col = kbj_df.columns[47]  # 所在诊所
        kbj_cm_col = kbj_df.columns[11]     # 客户经理
        kbj_amt_col = kbj_df.columns[33]    # 实际收入
        for _, row in kbj_df.iterrows():
            raw_store = str(row[kbj_store_col]).strip()
            # 只保留南区5家门店（广州/武汉/深圳/南京/杭州）
            if raw_store not in kbj_store_map:
                continue
            cm = str(row[kbj_cm_col]).strip()
            amt = float(row[kbj_amt_col]) if pd.notna(row[kbj_amt_col]) else 0
            store = kbj_store_map[raw_store]
            kangbojia_store_extra[store] = kangbojia_store_extra.get(store, 0) + amt
            # 每日汇总
            try:
                kbj_day = str(pd.to_datetime(row[kbj_time_col]).day)
                kangbojia_store_daily[store][kbj_day] += amt
            except: pass
            if cm and cm != 'nan' and cm != '测试医生':
                kangbojia_cm_extra[cm] = kangbojia_cm_extra.get(cm, 0) + amt
    else:
        # 兼容旧版预汇总数据：col[0]=门店, col[1]=实际收入
        store_col = kbj_df.columns[0]
        amt_col = kbj_df.columns[1]
        for _, row in kbj_df.iterrows():
            store = str(row[store_col]).strip()
            amt = float(row[amt_col]) if pd.notna(row[amt_col]) else 0
            if store and store != 'nan':
                kangbojia_store_extra[store] = amt
    print(f'📂 康博嘉数据(门店): {kangbojia_store_extra}')
    print(f'📂 康博嘉数据(CM): {kangbojia_cm_extra}')
kangbojia_total = sum(kangbojia_store_extra.values())

# ==================== 同期筛选 ====================
# 自动取当天为截止日，环比取上月同天（period_income 用 < end，所以 end = cutoff+1）
from datetime import date as _date, timedelta as _td
_today = _date.today()
JULY_CUTOFF = _today.strftime('%Y-%m-%d')
JULY_END = (_today + _td(days=1)).strftime('%Y-%m-%d')
JULY_START = f'{_today.year}-07-01'
# 上月同天（处理月底溢出：min(当天, 上月最后一天)）
_june_last = (_today.replace(day=1) - _td(days=1)).day
_june_day = min(_today.day, _june_last)
_june_cutoff = _today.replace(month=_today.month-1, day=_june_day) if _today.month > 1 else _today.replace(year=_today.year-1, month=12, day=_june_day)
JUNE_CUTOFF = _june_cutoff.strftime('%Y-%m-%d')
JUNE_END = (_june_cutoff + _td(days=1)).strftime('%Y-%m-%d')
JUNE_START = f'{_june_cutoff.year}-06-01'
# 显示用日期
_DATE_CN = f'{_today.month}月{_today.day}日'
_DATE_DOT = f'{_today.year%100:02d}.{_today.month:02d}.{_today.day:02d}'
_DATE_DAY = str(_today.day)
_DATE_JUNE_CN = f'{_june_cutoff.month}月{_june_cutoff.day}日'

# 周度定义：第一周=1日到第一个周日，之后周一到周日，不跨月
import calendar as _cal
_wk_last = _cal.monthrange(_today.year, _today.month)[1]
_wk_first_wday = _date(_today.year, _today.month, 1).weekday()  # 0=Mon
_wk1_end = 7 - _wk_first_wday if _wk_first_wday != 6 else 1
_week_defs = []  # [(label, start_day, end_day)]
_ws = 1
while _ws <= _wk_last:
    _we = min(_wk1_end if _ws == 1 else _ws + 6, _wk_last)
    _sd = _date(_today.year, _today.month, _ws)
    _ed = _date(_today.year, _today.month, _we)
    _week_defs.append((f'W{len(_week_defs)+1} ({_sd.month}/{_sd.day}-{_ed.month}/{_ed.day})', _ws, _we))
    _ws = _we + 1

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
putuo_locked_done = putuo_order_done + putuo_extra  # 普陀总业绩 = 底表 + 明细
putuo_delta = putuo_extra  # 明细作为额外增量叠加

for role in ['销售','班主任','门店']:
    grp = orders[orders['role']==role]
    tgt = sum(target_map.get(g,0) for g in grp['target_group'].dropna().unique())
    m = calc_metrics(grp, tgt)
    if role == '门店':
        # 普陀部分用锁定值替换
        m['done'] = m['done'] + putuo_delta
        # 康博嘉额外业绩
        m['done'] = m['done'] + kangbojia_total
        m['rate'] = safe_div(m['done'], tgt)
        m['net'] = m['done'] - m['refund']
        m['net_chg'] = safe_div(m['net'] - m['prev_net'], m['prev_net'])
    role_metrics[role] = m

# 南区合计
total_tgt = sum(r['target'] for r in role_metrics.values())
total_tgt = 1510_0000  # 总目标锁定 1510 万
total_grp = orders[orders['role'].isin(['销售','班主任','门店'])]
total_metrics = calc_metrics(total_grp, total_tgt)
# 合并普陀锁定
total_metrics['done'] = total_metrics['done'] + putuo_delta
# 合并康博嘉
total_metrics['done'] = total_metrics['done'] + kangbojia_total
total_metrics['target'] = 1510_0000
total_metrics['rate'] = safe_div(total_metrics['done'], 1510_0000)
total_metrics['net'] = total_metrics['done'] - total_metrics['refund']
total_metrics['net_chg'] = safe_div(total_metrics['net'] - total_metrics['prev_net'], total_metrics['prev_net'])

# ==================== Level 2: 按分组 ====================
# 所有分组 (按完成率升序 = 最差的排前面)
ALL_GROUPS = ['课程1组','课程2组','课程3组','课程4组','课程5组','课程6组',
              '班主任',
              '普陀','昆明','成都','广州','深圳','珠海','福州','杭州','南京','武汉','徐汇']

group_metrics = {}
for g in ALL_GROUPS:
    grp = orders[orders['target_group']==g]
    tgt = target_map.get(g, 0)
    if tgt == 0: continue  # skip groups without target
    is_putuo = (g == '普陀')
    group_metrics[g] = calc_metrics(grp, tgt, is_putuo)

# 注入康博嘉额外业绩到对应门店
for store, extra in kangbojia_store_extra.items():
    if store in group_metrics:
        gm = group_metrics[store]
        gm['done'] = gm['done'] + extra
        gm['rate'] = safe_div(gm['done'], gm['target'])
        gm['net'] = gm['done'] - gm['refund']
        gm['net_chg'] = safe_div(gm['net'] - gm['prev_net'], gm['prev_net'])

# 拆分线上组 / 线下门店
ONLINE_GROUPS = ['课程1组','课程2组','课程3组','课程4组','课程5组','课程6组','班主任']
OFFLINE_GROUPS = ['普陀','昆明','成都','广州','深圳','珠海','福州','杭州','南京','武汉','徐汇']

# 各自按完成率升序排列（差的在前）
ONLINE_ORDER = sorted([g for g in ONLINE_GROUPS if g in group_metrics], key=lambda g: group_metrics[g]['rate'])
OFFLINE_ORDER = sorted([g for g in OFFLINE_GROUPS if g in group_metrics], key=lambda g: group_metrics[g]['rate'])

# ==================== Level 3a: 线上个人 ====================
# ==================== 电话数据加载 (按销售姓名汇总通话人数) ====================
phone_calls = {}  # name -> 通话人数
PHONE_NAME_FIX = {'李钰坤': '李玉坤', '李家慧': '李家惠'}  # 错别字修正
try:
    phone_files = [f for f in glob.glob(os.path.join(data_dir, '电话', '*.xlsx'))
                   if not os.path.basename(f).startswith('~$')]
    if phone_files:
        phone_file = max(phone_files, key=os.path.getmtime)
        phone_df = pd.read_excel(phone_file)
        name_col = next((c for c in phone_df.columns if '姓名' in c or '销售' in c), phone_df.columns[0])
        call_col = next((c for c in phone_df.columns if '通话人数' in c or '通话' in c), None)
        if call_col:
            # 去重：防止同名同次数重复行导致翻倍
            dedup_cols = [name_col, call_col]
            org_col = next((c for c in phone_df.columns if '组织' in c or '部门' in c), None)
            if org_col:
                dedup_cols.insert(1, org_col)
            before = len(phone_df)
            phone_df = phone_df.drop_duplicates(subset=dedup_cols)
            if len(phone_df) < before:
                print(f'📞 电话去重: {before} → {len(phone_df)} 行')
            for _, prow in phone_df.iterrows():
                nm = str(prow[name_col]).strip()
                if not nm or nm == 'nan': continue
                # 错别字修正
                nm = PHONE_NAME_FIX.get(nm, nm)
                cv = prow[call_col]
                cv = int(cv) if pd.notna(cv) else 0
                phone_calls[nm] = phone_calls.get(nm, 0) + cv
        print(f'📞 电话数据: {os.path.basename(phone_file)} — {len(phone_calls)}人 共{sum(phone_calls.values())}通话')
except Exception as e:
    print(f'⚠️ 电话数据加载失败: {e}')

person_online_list = []
for (grp_name, person_name), tgt in online_person_target.items():
    p_orders = orders[(orders['seller_name']==person_name)]
    m = calc_metrics(p_orders, tgt)
    person_online_list.append({
        'group': grp_name, 'name': person_name, 'calls': phone_calls.get(person_name, 0), **m
    })
person_online_df = pd.DataFrame(person_online_list)

# ==================== Level 3b: 线下个案 ====================
person_cm_list = []
for _, cm in cm_df.iterrows():
    store = cm['store']; name = cm['name']; tgt = cm['target']
    is_putuo = (store == '普陀')
    p_orders = orders[(orders['seller_name']==name)]
    # 特殊归属销售（如余佼）多店挂名时：底表订单归入 override 目标门店行，
    # 原门店行（普陀）只保留 HG0019 明细，不叠加底表订单
    if name in SELLER_ORG_OVERRIDE:
        override_target = STORE_MAP.get(SELLER_ORG_OVERRIDE[name], '')
        if store != override_target:
            # 非目标门店行（如余佼的普陀行）：只保留 HG0019 明细，排除底表订单
            p_orders = p_orders.iloc[0:0]
    if is_putuo:
        # 普陀个案业绩 = 明细汇总 + 底表订单
        done = cm['done_locked'] + period_income(p_orders, JULY_START, JULY_END)
        ref_raw = period_refund(p_orders, JULY_START, JULY_END)
        prev_raw = cm['prev_month']
        prev_ref_raw = cm['refund_locked']
    else:
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
        'is_putuo': is_putuo, 'calls': phone_calls.get(name, 0),
    })
person_cm_df = pd.DataFrame(person_cm_list)

# 注入康博嘉额外业绩到对应CM个人
for cm_name, extra in kangbojia_cm_extra.items():
    mask = person_cm_df['name'] == cm_name
    if mask.any():
        idx = person_cm_df[mask].index[0]
        person_cm_df.at[idx, 'done'] = person_cm_df.at[idx, 'done'] + extra
        person_cm_df.at[idx, 'rate'] = safe_div(person_cm_df.at[idx, 'done'], person_cm_df.at[idx, 'target'])
        person_cm_df.at[idx, 'net'] = person_cm_df.at[idx, 'done'] - person_cm_df.at[idx, 'refund']
        person_cm_df.at[idx, 'net_chg'] = safe_div(person_cm_df.at[idx, 'net'] - person_cm_df.at[idx, 'prev_net'], person_cm_df.at[idx, 'prev_net'])

# ==================== 电话数汇总 (分组/角色) ====================
group_calls = {}  # group -> 通话总数
for _, r in person_online_df.iterrows():
    group_calls[r['group']] = group_calls.get(r['group'], 0) + int(r.get('calls', 0))
for _, r in person_cm_df.iterrows():
    group_calls[r['store']] = group_calls.get(r['store'], 0) + int(r.get('calls', 0))

group_count = {}  # group -> 人数
for _, r in person_online_df.iterrows():
    group_count[r['group']] = group_count.get(r['group'], 0) + 1
for _, r in person_cm_df.iterrows():
    group_count[r['store']] = group_count.get(r['store'], 0) + 1

role_calls = {'销售': 0, '班主任': 0, '门店': 0}
role_count = {'销售': 0, '班主任': 0, '门店': 0}
for _, r in person_online_df.iterrows():
    if str(r['group']).startswith('课程'):
        role_calls['销售'] += int(r.get('calls', 0)); role_count['销售'] += 1
    elif '班主任' in str(r['group']):
        role_calls['班主任'] += int(r.get('calls', 0)); role_count['班主任'] += 1
for _, r in person_cm_df.iterrows():
    role_calls['门店'] += int(r.get('calls', 0)); role_count['门店'] += 1

def avg_calls(total, cnt):
    return f'{total/cnt:.1f}' if cnt > 0 else '—'

# ==================== HTML 渲染 ====================
def fmt(n):
    if abs(n) >= 1e8: return f'{n/1e8:.1f}亿'
    if abs(n) >= 1e4: return f'{n/1e4:.1f}万'
    return f'{n:,.1f}'

def pct1(n): return f'{n:+.1f}%' if n and n != float('inf') else '—'
def pct2(n): return f'{n*100:.1f}%'
def chg_html(val):
    if val is None or val == float('inf') or val == -float('inf'): return '<span style="color:#94a3b8">—</span>'
    c = '#22c55e' if val >= 0 else '#ef4444'
    arrow = '↑' if val >= 0 else '↓'
    return f'<span style="color:{c};font-weight:600">{arrow}{abs(val)*100:.1f}%</span>'

def refund_chg_html(val):
    """退费环比：下降=绿，上升=红"""
    if val is None or val == float('inf') or val == -float('inf'): return '<span style="color:#94a3b8">—</span>'
    c = '#ef4444' if val > 0 else '#22c55e' if val < 0 else '#94a3b8'
    arrow = '↑' if val > 0 else '↓' if val < 0 else '→'
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

from datetime import date
_today = date.today()
TIME_PROGRESS = min(_today.day, 31) / 31  # 月度时间进度，动态当天日期

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
  <td class="num" style="color:#0891b2;font-weight:600">{role_calls.get(role, 0)}</td>
  <td class="num col-avgcall" style="color:#0e7490;font-weight:600">{avg_calls(role_calls.get(role,0), role_count.get(role,0))}</td>
</tr>'''

# Build group-level rows (Level 2)
def group_row(g, m, role):
    note = ' <span style="font-size:10px;color:#f59e0b">⚠锁定</span>' if g == '普陀' else ''
    return f'''<tr>
  <td><b>{g}</b>{note}</td>
  <td><div style="display:flex;align-items:center;gap:8px">{role_badge(role)}<canvas class="spark" data-group="{g}" width="80" height="26" style="cursor:pointer;border-radius:4px" title="悬停查看每日趋势"></canvas></div></td>
  <td class="num">{fmt(m['target'])}</td>
  <td class="num"><b>{fmt(m['done'])}</b></td>
  {rate_td(m['rate'])}
  <td class="num">{chg_html(m['done_chg'])}</td>
  <td class="num">{fmt(m['refund'])}</td>
  {refund_td(m['refund'], m['done'])}
  <td class="num">{refund_chg_html(m['refund_chg'])}</td>
  <td class="num"><b>{fmt(m['net'])}</b></td>
  <td class="num">{fmt(m['prev_net'])}</td>
  <td class="num">{chg_html(m['net_chg'])}</td>
  <td class="num" style="color:#0891b2;font-weight:600">{group_calls.get(g, 0)}</td>
  <td class="num col-avgcall" style="color:#0e7490;font-weight:600">{avg_calls(group_calls.get(g,0), group_count.get(g,0))}</td>
</tr>
<tr class="weekly-row" data-weekly-group="{g}"><td colspan="2" style="text-align:right;font-size:10px;color:#94a3b8;padding-right:4px">📅周</td><td colspan="12"><div class="weekly-bar-row" data-group="{g}" style="display:flex;gap:12px;align-items:flex-end;justify-content:center;padding:4px 0"></div></td></tr>'''

# Build person rows (Level 3)
def person_row(name, group_label, tgt, m, person_type='sales_advisor'):
    note = ' <span style="font-size:10px;color:#f59e0b">⚠锁定</span>' if m.get('is_putuo') else ''
    calls = int(m.get('calls', 0)) if hasattr(m, 'get') else 0
    return f'''<tr>
  <td>{group_label}</td>
  <td>{name}{note}</td>
  <td class="num">{fmt(tgt)}</td>
  <td class="num"><b>{fmt(m['done'])}</b></td>
  {rate_td(m['rate'])}
  <td class="num">{chg_html(m['done_chg'])}</td>
  <td class="num">{fmt(m['refund'])}</td>
  {refund_td(m['refund'], m['done'], person_type)}
  <td class="num">{refund_chg_html(m['refund_chg'])}</td>
  <td class="num"><b>{fmt(m['net'])}</b></td>
  <td class="num">{fmt(m['prev_net'])}</td>
  <td class="num">{chg_html(m['net_chg'])}</td>
  <td class="num" style="color:#0891b2;font-weight:600">{calls}</td>
</tr>'''

# ==================== 生成 HTML ====================
html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
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
<p class="subtitle">数据截止：2026年{_DATE_CN} | 环比基期：6月1-{_DATE_DAY}日同口径 | 普陀业绩从明细自动汇总</p></div>
<div class="top-btns" style="display:flex;gap:8px;align-items:center;margin-top:4px">
<button onclick="exportBaseTable()" style="padding:8px 18px;background:#fff;color:#1e293b;border:1px solid #d1d5db;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">📥 导出底表</button>
<button onclick="snapshot()" style="padding:8px 18px;background:#1e293b;color:#fff;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">📸 快照</button>
<label style="padding:8px 18px;background:#1e293b;color:#fff;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">
  📤 导入底表 <input type="file" onchange="importBaseTable(this)" accept=".xlsx,.xls,.csv" style="display:none">
</label>
<label style="padding:8px 18px;background:#ecfdf5;color:#065f46;border:1px solid #6ee7b7;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">
  📤 导入康博嘉 <input type="file" onchange="importKangbojia(this)" accept=".xlsx,.xls" style="display:none">
</label>
<label style="padding:8px 18px;background:#fef3c7;color:#b45309;border:1px solid #f59e0b;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;white-space:nowrap">
  📤 导入普陀明细 <input type="file" onchange="importPutuoDetail(this)" accept=".xlsx,.xls" style="display:none">
</label>
</div>
</div>

<!-- Timeline Progress Bar -->
<div style="margin-bottom:16px;padding:14px 20px;background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,0.04)">
  <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:6px">
    <span style="font-size:13px;font-weight:700;color:#1e293b">📅 7月进度</span>
    <span style="font-size:13px;color:#64748b">已过 <b style="color:#1e293b">{int(TIME_PROGRESS*31)}</b>/31 天 · <b style="color:#6366f1">{pct2(TIME_PROGRESS)}</b></span>
  </div>
  <div style="position:relative;height:8px;background:#f1f5f9;border-radius:4px;overflow:hidden">
    <div style="width:{TIME_PROGRESS*100:.1f}%;height:100%;background:linear-gradient(90deg,#6366f1,#8b5cf6);border-radius:4px;transition:width 0.6s ease"></div>
  </div>
  <div style="display:flex;justify-content:space-between;margin-top:4px;font-size:11px;color:#94a3b8">
    <span>7/1</span><span>7/10</span><span>7/20</span><span>7/31</span>
  </div>
</div>

<!-- KPI Cards -->
<div class="kpi-grid">
<div class="kpi-card">
  <div class="label">📌 月总目标</div>
  <div class="value">{fmt(total_metrics['target'])}</div>
  <div class="sub">线上 {fmt(section_tgts_display['线上'])} + 线下 {fmt(section_tgts_display['线下门店'])}</div>
</div>
<div class="kpi-card">
  <div class="label">✅ 累计完成</div>
  <div class="value">{fmt(total_metrics['done'])}</div>
  {rate_bar_with_trend(total_metrics['rate'])}
  <div class="sub">上月同期 {fmt(total_metrics['prev_done'])} | 环比 {pct1(total_metrics['done_chg'])}</div>
</div>
<div class="kpi-card">
  <div class="label">📅 剩余天数</div>
  <div class="value">{31 - int(TIME_PROGRESS*31)} 天</div>
  <div class="sub">7月{int(TIME_PROGRESS*31)+1}日 - 7月31日</div>
</div>
<div class="kpi-card">
  <div class="label">🎯 日均冲刺</div>
  <div class="value">{fmt(safe_div(total_metrics['target'] - total_metrics['done'], 31 - int(TIME_PROGRESS*31)))}</div>
  <div class="sub">剩余 {fmt(total_metrics['target'] - total_metrics['done'])} ÷ {31 - int(TIME_PROGRESS*31)}天</div>
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
<button class="tab" data-tab="channel" onclick="switchTab('channel');loadChannelData()">📢 渠道业绩</button>
</div>

<!-- Tab 1: 按角色 -->
<div id="tab-role" class="tab-content active">
<table>
<thead><tr>
  <th>角色</th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">本月净流水</th><th class="num">📞电话数</th><th class="num col-avgcall">人均电话 <span onclick="toggleAvgCall()" style="cursor:pointer;user-select:none" title="点击隐藏/显示">👁️</span></th>
</tr></thead><tbody>
{''.join(role_row(r, role_metrics[r]) for r in ['销售','班主任','门店'])}
<tr class="subtotal"><td>南区合计</td>
  <td class="num">{fmt(total_metrics['target'])}</td><td class="num">{fmt(total_metrics['done'])}</td>{rate_td(total_metrics['rate'])}
  <td class="num">{fmt(total_metrics['refund'])}</td>{refund_td(total_metrics['refund'], total_metrics['done'])}
  <td class="num">{fmt(total_metrics['net'])}</td>
  <td class="num" style="color:#0891b2;font-weight:700">{sum(role_calls.values())}</td>
  <td class="num col-avgcall" style="color:#0e7490;font-weight:700">{avg_calls(sum(role_calls.values()), sum(role_count.values()))}</td>
</tr>
</tbody></table>
<div style="margin-top:20px">
  <h4 style="margin:0 0 14px;font-size:15px;color:#1e293b;font-weight:700">📈 每日业绩趋势 <span style="font-size:12px;color:#94a3b8;font-weight:400">· 悬停查看详情</span></h4>
  <div style="display:grid;grid-template-columns:1fr;gap:16px">
    <div style="background:linear-gradient(135deg,#eef2ff,#fff);padding:18px 20px;border-radius:14px;box-shadow:0 2px 12px rgba(99,102,241,0.08);border:1px solid #e0e7ff">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px"><span style="width:10px;height:10px;background:#6366f1;border-radius:3px"></span><span style="font-size:13px;color:#4338ca;font-weight:700">销售</span></div>
      <div style="height:200px"><canvas id="chart_sales"></canvas></div>
    </div>
    <div style="background:linear-gradient(135deg,#ecfdf5,#fff);padding:18px 20px;border-radius:14px;box-shadow:0 2px 12px rgba(16,185,129,0.08);border:1px solid #d1fae5">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px"><span style="width:10px;height:10px;background:#10b981;border-radius:3px"></span><span style="font-size:13px;color:#047857;font-weight:700">班主任</span></div>
      <div style="height:200px"><canvas id="chart_advisor"></canvas></div>
    </div>
    <div style="background:linear-gradient(135deg,#fffbeb,#fff);padding:18px 20px;border-radius:14px;box-shadow:0 2px 12px rgba(245,158,11,0.08);border:1px solid #fef3c7">
      <div style="display:flex;align-items:center;gap:8px;margin-bottom:10px"><span style="width:10px;height:10px;background:#f59e0b;border-radius:3px"></span><span style="font-size:13px;color:#b45309;font-weight:700">门店</span></div>
      <div style="height:200px"><canvas id="chart_store"></canvas></div>
    </div>
  </div>
</div>
<div style="margin-top:24px">
  <h4 style="margin:0 0 14px;font-size:15px;color:#1e293b;font-weight:700">📅 每周业绩趋势 <span style="font-size:12px;color:#94a3b8;font-weight:400">· 周度对比</span></h4>
  <div id="weekly-role-container" style="display:grid;grid-template-columns:repeat(3,1fr);gap:12px"></div>
</div>
</div>

<!-- Tab 2: 按分组 -->
<div id="tab-group" class="tab-content">
<table>
<thead><tr>
  <th>分组</th><th>角色 <span style="font-weight:400;color:#94a3b8;font-size:10px">/趋势</span></th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th><th class="num">完成环比</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">退费环比</th>
  <th class="num">本月净流水</th><th class="num">上月同期净流水</th><th class="num">净流水环比</th><th class="num">📞电话数</th><th class="num col-avgcall">人均电话 <span onclick="toggleAvgCall()" style="cursor:pointer;user-select:none" title="点击隐藏/显示">👁️</span></th>
</tr></thead><tbody>
'''

# Group rows — 线上组
html += '<tr class="section-header"><td colspan="14" style="background:#dbeafe;color:#1e40af;font-weight:700;font-size:13px;padding:10px 12px;border-bottom:2px solid #93c5fd">📡 线上组</td></tr>'
for g in ONLINE_ORDER:
    m = group_metrics.get(g)
    if m is None or m['target'] == 0: continue
    role = classify_role(orders[orders['target_group']==g]['组织'].iloc[0]) if len(orders[orders['target_group']==g]) > 0 else '其他'
    html += group_row(g, m, role)

# Group rows — 线下门店
html += '<tr class="section-header"><td colspan="14" style="background:#d1fae5;color:#065f46;font-weight:700;font-size:13px;padding:10px 12px;border-bottom:2px solid #6ee7b7">🏥 线下门店</td></tr>'
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
  <th class="num">本月净流水</th><th class="num">上月同期净流水</th><th class="num">净流水环比</th><th class="num">📞电话数</th>
</tr></thead><tbody>
'''

person_online_sorted = person_online_df.sort_values(['group','rate'], ascending=[True,True])
last_group = ''
group_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0,'calls':0}
for _, r in person_online_sorted.iterrows():
    if r['group'] != last_group:
        if last_group and group_sub['target'] > 0:
            sub_done = group_sub['done']; sub_ref = group_sub['refund']
            sub_net = group_sub['net']; sub_prev_net = group_sub['prev_net']
            html += f'''<tr class="subtotal"><td>{last_group}</td><td>小计</td>
  <td class="num">{fmt(group_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,group_sub['target']))}
  <td class="num">{chg_html(safe_div(sub_done-group_sub['prev_done'],group_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{refund_chg_html(safe_div(sub_ref-group_sub['prev_refund'],group_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td><td class="num" style="color:#0891b2;font-weight:700">{group_sub['calls']}</td></tr>'''
        last_group = r['group']
        group_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0,'calls':0}
    group_sub['target'] += r['target']
    group_sub['done'] += r['done']
    group_sub['refund'] += r['refund']
    group_sub['prev_done'] += r['prev_done']
    group_sub['prev_refund'] += r['prev_refund']
    group_sub['net'] += r['net']
    group_sub['prev_net'] += r['prev_net']
    group_sub['calls'] += int(r.get('calls', 0))
    html += person_row(r['name'], r['group'], r['target'], r, 'sales_advisor')
# last group subtotal
if last_group and group_sub['target'] > 0:
    sub_done = group_sub['done']; sub_ref = group_sub['refund']
    sub_net = group_sub['net']; sub_prev_net = group_sub['prev_net']
    html += f'''<tr class="subtotal"><td>{last_group}</td><td>小计</td>
  <td class="num">{fmt(group_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,group_sub['target']))}</td>
  <td class="num">{chg_html(safe_div(sub_done-group_sub['prev_done'],group_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{refund_chg_html(safe_div(sub_ref-group_sub['prev_refund'],group_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td><td class="num" style="color:#0891b2;font-weight:700">{group_sub['calls']}</td></tr>'''

html += '''</tbody></table></div>'''

# Tab 4: 线下个案
html += '''<div id="tab-cm_person" class="tab-content">
<table>
<thead><tr>
  <th>分院</th><th>个案经理</th><th class="num">目标</th><th class="num">完成</th><th style="min-width:170px">完成率</th><th class="num">完成环比</th>
  <th class="num">本月退费</th><th class="num">退费率</th><th class="num">退费环比</th>
  <th class="num">本月净流水</th><th class="num">上月同期净流水</th><th class="num">净流水环比</th><th class="num">📞电话数</th>
</tr></thead><tbody>
'''

cm_sorted = person_cm_df.sort_values(['store','rate'], ascending=[True,True])
last_store = ''
store_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0,'calls':0}
for _, r in cm_sorted.iterrows():
    if r['store'] != last_store:
        if last_store and store_sub['target'] > 0:
            sub_done = store_sub['done']; sub_ref = store_sub['refund']
            sub_net = store_sub['net']; sub_prev_net = store_sub['prev_net']
            html += f'''<tr class="subtotal"><td><b>{last_store}</b></td><td>小计</td>
  <td class="num">{fmt(store_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,store_sub['target']))}</td>
  <td class="num">{chg_html(safe_div(sub_done-store_sub['prev_done'],store_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{refund_chg_html(safe_div(sub_ref-store_sub['prev_refund'],store_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td><td class="num" style="color:#0891b2;font-weight:700">{store_sub['calls']}</td></tr>'''
        last_store = r['store']
        store_sub = {'target':0,'done':0,'refund':0,'prev_done':0,'prev_refund':0,'net':0,'prev_net':0,'calls':0}
    store_sub['target'] += r['target']
    store_sub['done'] += r['done']
    store_sub['refund'] += r['refund']
    store_sub['prev_done'] += r['prev_done']
    store_sub['prev_refund'] += r['prev_refund']
    store_sub['net'] += r['net']
    store_sub['prev_net'] += r['prev_net']
    store_sub['calls'] += int(r.get('calls', 0))
    html += person_row(r['name'], r['store'], r['target'], r, 'cm')
# last store subtotal
if last_store and store_sub['target'] > 0:
    sub_done = store_sub['done']; sub_ref = store_sub['refund']
    sub_net = store_sub['net']; sub_prev_net = store_sub['prev_net']
    html += f'''<tr class="subtotal"><td><b>{last_store}</b></td><td>小计</td>
  <td class="num">{fmt(store_sub['target'])}</td><td class="num">{fmt(sub_done)}</td>{rate_td(safe_div(sub_done,store_sub['target']))}</td>
  <td class="num">{chg_html(safe_div(sub_done-store_sub['prev_done'],store_sub['prev_done']))}</td>
  <td class="num">{fmt(sub_ref)}</td>{refund_td(sub_ref, sub_done)}
  <td class="num">{refund_chg_html(safe_div(sub_ref-store_sub['prev_refund'],store_sub['prev_refund']))}</td>
  <td class="num">{fmt(sub_net)}</td><td class="num">{fmt(sub_prev_net)}</td>
  <td class="num">{chg_html(safe_div(sub_net-sub_prev_net,sub_prev_net))}</td><td class="num" style="color:#0891b2;font-weight:700">{store_sub['calls']}</td></tr>'''

html += '''</tbody></table></div>'''

html += '''<div id="tab-channel" class="tab-content">
<div style="text-align:center;padding:40px;color:#94a3b8">👆 点击「渠道业绩」标签加载本地生活数据</div>
</div>'''

html += f'''
<div class="footer">
  南区7月目标进展看板 | 底表数据截止 20{_DATE_DOT} | 环比 = 7月1-{_DATE_DAY}日 vs 6月1-{_DATE_DAY}日同口径 | 普陀业绩从明细自动汇总
</div>
</div>
<script>
function switchTab(name){{
  document.querySelectorAll(".tab").forEach(function(t){{t.classList.remove("active")}});
  document.querySelectorAll(".tab-content").forEach(function(t){{t.classList.remove("active")}});
  document.querySelector('.tab[data-tab="'+name+'"]').classList.add("active");
  document.getElementById("tab-"+name).classList.add("active");
  if(name === 'role') setTimeout(drawTrendChart, 100);
}}
var _avgCallHidden = false;
function toggleAvgCall(){{
  _avgCallHidden = !_avgCallHidden;
  document.querySelectorAll(".col-avgcall").forEach(function(el){{
    // 保留列头的眼睛可点，隐藏数值单元格；列头文字变淡
    if(el.tagName === 'TD'){{ el.style.display = _avgCallHidden ? 'none' : ''; }}
    else {{
      var txt = el.childNodes[0];
      if(txt) txt.textContent = _avgCallHidden ? '' : '人均电话 ';
    }}
  }});
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

# 菌群业绩 person_daily (仅肠道菌群移植)
jq_orders = july_orders[july_orders['service_type_name'].str.contains('菌群', na=False)]
person_daily_junqun = {}
for _, row in jq_orders.iterrows():
    name = row['seller_name']
    if pd.isna(name) or not name: continue
    day = row['pay_at'].day
    amt = float(row['pay_amount'])
    if name not in person_daily_junqun: person_daily_junqun[name] = {}
    person_daily_junqun[name][str(day)] = person_daily_junqun[name].get(str(day), 0) + amt
    person_daily_junqun[name]['_total'] = person_daily_junqun[name].get('_total', 0) + amt
    person_daily_junqun[name]['_org'] = row['组织']

# 每日业绩汇总 (趋势图用 — 按角色)
role_daily = {'销售': defaultdict(float), '班主任': defaultdict(float), '门店': defaultdict(float)}
group_daily = defaultdict(lambda: defaultdict(float))  # 按分组每日
for _, row in july_orders.iterrows():
    org = row['组织']
    role = classify_role(org)
    day = str(row['pay_at'].day)
    amt = float(row['pay_amount'])
    if role in role_daily:
        role_daily[role][day] += amt
    grp = row.get('target_group')
    if pd.notna(grp):
        group_daily[grp][day] += amt
# 转普通 dict
for r in role_daily:
    role_daily[r] = dict(role_daily[r])
group_daily = {g: dict(v) for g, v in group_daily.items()}

# 趋势图截止日 = 底表最新支付日（避免普陀/康博嘉出现半天数据点）
# 如果最新日=今天（数据不完整），回退到昨天
trend_max_day = int(july_orders['pay_at'].max().day) if len(july_orders) else 31
if trend_max_day >= _today.day:
    trend_max_day = _today.day - 1

# 普陀明细每日叠加到门店和普陀分组（只到趋势截止日）
if putuo_file and 'CM姓名' in putuo_df.columns and '收款金额' in putuo_df.columns:
    for _, prow in putuo_df.iterrows():
        pd_day = pd.to_datetime(prow['日期']).day
        if pd_day > trend_max_day: continue
        d = str(pd_day)
        amt = float(prow['收款金额']) if pd.notna(prow['收款金额']) else 0
        role_daily['门店'][d] = role_daily['门店'].get(d, 0) + amt
        group_daily['普陀'][d] = group_daily['普陀'].get(d, 0) + amt

# 康博嘉每日叠加到门店和对应分组（仅南区5家，只到趋势截止日）
for store, dvals in kangbojia_store_daily.items():
    for d, amt in dvals.items():
        if int(d) > trend_max_day: continue
        role_daily['门店'][d] = role_daily['门店'].get(d, 0) + amt
        group_daily[store][d] = group_daily[store].get(d, 0) + amt

# 截断趋势图：排除 trend_max_day 之后的数据（当天数据不完整）
for r in role_daily:
    role_daily[r] = {d: v for d, v in role_daily[r].items() if int(d) <= trend_max_day}
group_daily = {g: {d: v for d, v in dv.items() if int(d) <= trend_max_day} for g, dv in group_daily.items()}

# ====== 预加载电商本地生活数据 ======
print('🔄 拉取本地生活数据...', end=' ')
eco_data_raw = []
try:
    import urllib.request
    req = urllib.request.Request('http://192.168.110.116:8899/api/sheet/ecommerce')
    with urllib.request.urlopen(req, timeout=30) as resp:
        eco_api = json.loads(resp.read().decode('utf-8'))
    eco_sheet = eco_api.get('sheets', {}).get('7月电商总表', {})
    eco_data_raw = eco_sheet.get('rows', [])
    print(f'✅ {len(eco_data_raw)} 条')
except Exception as e:
    print(f'⚠️ 失败: {e}')

mapping_json = json.dumps({
    'org_to_group': ORG_TO_GROUP,
    'seller_org_override': SELLER_ORG_OVERRIDE,
    'store_to_group': STORE_MAP,
    'store_to_trio': STORE_TO_TRIO,
    'sales_orgs': SALES_ORGS,
    'advisor_orgs': ADVISOR_ORGS,
    'target_map': target_map,
    'putuo_locked': putuo_locked_done,
    'putuo_detail_total': putuo_extra,
    'putuo_base_orders': putuo_order_done,
    'online_person_target': {f'{k[0]}|{k[1]}': v for k, v in online_person_target.items()},
    'cm_list': cm_list,
    'time_progress': TIME_PROGRESS,
    'july_start': JULY_START, 'july_end': JULY_END,
    'june_start': JUNE_START, 'june_end': JUNE_END,
    'person_daily': person_daily,
    'person_daily_junqun': person_daily_junqun,
    'role_daily': role_daily,
    'group_daily': group_daily,
    'week_defs': _week_defs,
}, ensure_ascii=False)

# ==================== 交互式 JS 工具条 ====================
import_toolbar = '''
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<script>
window._M = ''' + mapping_json + ''';
window._TP = window._M.time_progress;
window._ECO = ''' + json.dumps(eco_data_raw, ensure_ascii=False) + ''';
// 初始化普陀CM业绩基线（用于导入时回退，防止重复叠加）
window._M._prev_putuo_cm = {};
(window._M.cm_list || []).forEach(function(cm) {
  if (cm.store === '普陀' && cm.done_locked > 0) {
    window._M._prev_putuo_cm[cm.name] = cm.done_locked;
  }
});
</script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
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

function fmtNum(n) { return n >= 1e8 ? (n/1e8).toFixed(1)+'亿' : n >= 1e4 ? (n/1e4).toFixed(1)+'万' : n.toLocaleString('en-US',{minimumFractionDigits:1,maximumFractionDigits:1}); }
function fmtInt(n) { return n >= 1e8 ? (n/1e8).toFixed(1)+'亿' : n >= 1e4 ? (n/1e4).toFixed(1)+'万' : Math.round(n).toLocaleString('en-US'); }

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
  // 先按分组行同步角色行（销售=课程组/班主任/门店），保证导入·编辑后按角色表和KPI一致
  var _gAgg={"销售":{done:0,ref:0},"班主任":{done:0,ref:0},"门店":{done:0,ref:0}};
  document.querySelectorAll("#tab-group tbody tr:not(.subtotal):not(.section-header)").forEach(function(row){
    var t=row.querySelectorAll("td"); if(t.length<12) return;
    var gn=t[0].textContent.replace("⚠锁定","").trim();
    var rk=gn.indexOf("课程")===0?"销售":(gn==="班主任"?"班主任":"门店");
    _gAgg[rk].done+=parseNum(t[3]); _gAgg[rk].ref+=parseNum(t[6]);
  });
  var _rNames=["销售","班主任","门店"];
  document.querySelectorAll("#tab-role tbody tr:not(.subtotal)").forEach(function(row,i){
    var a=_gAgg[_rNames[i]]; if(!a) return; var t=row.querySelectorAll("td"); if(t.length<7) return;
    writeNum(t[2],a.done); writeNum(t[4],a.ref);
  });
  // role tab
  document.querySelectorAll("#tab-role tbody tr:not(.subtotal)").forEach(function(row){tds=row.querySelectorAll("td");if(tds.length<7)return;var t=parseNum(tds[1]),d=parseNum(tds[2]),r=parseNum(tds[4]);rTotals.tgt+=t;rTotals.done+=d;rTotals.ref+=r;rTotals.net+=d-r;writeNum(tds[5],d-r);rebuildBar(tds[3],d/t);rebuildRefund(tds[4].nextElementSibling,r,d,"group")});
  st=document.querySelector("#tab-role tr.subtotal");if(st){tds=st.querySelectorAll("td");writeNum(tds[1],rTotals.tgt);writeNum(tds[2],rTotals.done);rebuildBar(tds[3],rTotals.done/rTotals.tgt);writeNum(tds[4],rTotals.ref);writeNum(tds[5],rTotals.net);rebuildRefund(tds[4].nextElementSibling,rTotals.ref,rTotals.done,"group")}
  // group tab
  document.querySelectorAll("#tab-group tbody tr:not(.subtotal):not(.section-header)").forEach(function(row){tds=row.querySelectorAll("td");if(tds.length<12)return;var t=parseNum(tds[2]),d=parseNum(tds[3]),r=parseNum(tds[6]);writeNum(tds[9],d-r);rebuildBar(tds[4],d/t);rebuildRefund(tds[7],r,d,"group")});
  // person tabs
  ["online_person","cm_person"].forEach(function(tabId){var tab=document.getElementById("tab-"+tabId);if(!tab)return;var grps={};tab.querySelectorAll("tbody tr:not(.subtotal)").forEach(function(row){tds=row.querySelectorAll("td");if(tds.length<12)return;var grp=tds[0].textContent.trim(),t=parseNum(tds[2]),d=parseNum(tds[3]),r=parseNum(tds[6]);writeNum(tds[10],d-r);rebuildBar(tds[4],d/t);rebuildRefund(tds[7],r,d,tabId==="tab-cm_person"?"cm":"sales");if(!grps[grp])grps[grp]={tgt:0,done:0,ref:0,net:0};grps[grp].tgt+=t;grps[grp].done+=d;grps[grp].ref+=r;grps[grp].net+=d-r});tab.querySelectorAll("tbody tr.subtotal").forEach(function(st){tds=st.querySelectorAll("td");g=grps[tds[0].textContent.trim()];if(!g||tds.length<10)return;writeNum(tds[2],g.tgt);writeNum(tds[3],g.done);rebuildBar(tds[4],g.done/g.tgt);writeNum(tds[6],g.ref);rebuildRefund(tds[7],g.ref,g.done,tabId==="tab-cm_person"?"cm":"sales");writeNum(tds[9],g.net)})});
  // KPIs
  var cards=document.querySelectorAll(".kpi-card .value"),tp=window._TP,rem=31-Math.round(tp*31);
  if(cards[0])cards[0].textContent=fmtNum(rTotals.tgt);
  if(cards[1]){cards[1].textContent=fmtNum(rTotals.done);rebuildBar(cards[1].nextElementSibling,rTotals.done/rTotals.tgt)}
  if(cards[4]){var sprint=rem>0?(rTotals.tgt-rTotals.done)/rem:0;cards[4].textContent=fmtNum(sprint);}
}
recalcAll = refreshAll;
document.addEventListener("blur", function(e){var td=e.target.closest('td.num[contenteditable="true"]');if(!td)return;var raw=td.textContent.trim().replace(/[,，万]/g,"");if(raw.indexOf("亿")>=0)raw=raw.replace("亿","")*1e8;var v=parseFloat(raw);if(!isNaN(v)){td.setAttribute("data-val",v);td.textContent=fmtNum(v)}td.contentEditable="false";setTimeout(refreshAll,100)},true);

// ====== TOP10 弹窗排行榜 (含日期筛选) ======
function buildRanking(M, pd, fromDay, toDay, isOnline, skipLocked, trioGroup) {
  var personDone = {};
  Object.keys(pd).forEach(function(name) {
    var info = pd[name];
    var org = info._org || '';
    var g = M.org_to_group[org] || M.store_to_group[org] || '';
    // 始终按角色过滤
    if(isOnline && M.store_to_group[org]) return;
    if(!isOnline && !M.store_to_group[org]) return;
    // 三组拆分过滤
    if(trioGroup) {
      var store = M.store_to_group[org] || '';
      if(M.store_to_trio[store] !== trioGroup) return;
    }
    var done = 0;
    for(var d=fromDay; d<=toDay; d++) { done += (info[String(d)] || 0); }
    if(done <= 0) return;
    personDone[name] = {name: name, group: g, done: done, target: 0};
  });
  // 普陀锁定数据（仅业绩排名，菌群不混入；三组模式下仅一组包含普陀）
  if(!skipLocked && !isOnline && (!trioGroup || M.store_to_trio['普陀'] === trioGroup)) {
    var cmTab = document.getElementById('tab-cm_person');
    if(cmTab) {
      cmTab.querySelectorAll('tbody tr:not(.subtotal)').forEach(function(row) {
        var tds = row.querySelectorAll('td');
        if(tds.length < 4) return;
        if(tds[0].textContent.trim() !== '普陀') return;
        var nm = tds[1].textContent.replace('⚠锁定','').trim();
        var tv = parseNum(tds[3]), tgt = parseNum(tds[2]);
        if(tv > 0) {
          if(!personDone[nm]) personDone[nm] = {name: nm, group: '普陀', done: 0, target: 0};
          personDone[nm].done = tv; personDone[nm].target = tgt;
        }
      });
    }
  }
  // 目标匹配（仅业绩排名）
  if(!skipLocked) {
    if(isOnline) {
      Object.keys(M.online_person_target || {}).forEach(function(k) {
        var n = k.split('|')[1];
        if(personDone[n]) personDone[n].target = M.online_person_target[k];
      });
    } else {
      (M.cm_list || []).forEach(function(cm) {
        if(personDone[cm.name]) personDone[cm.name].target = cm.target;
      });
    }
  }
  var data = Object.values(personDone).filter(function(d){ return d.done > 0; });
  data.sort(function(a,b){ return b.done - a.done; });
  return data.slice(0, 10);
}

function renderRankList(data, crowns, colors) {
  var h = '';
  data.forEach(function(p, i) {
    var rank = i<3 ? '<span style="font-size:22px">'+crowns[i]+'</span>' : '<span style="width:28px;text-align:center;color:#94a3b8;font-weight:700">'+(i+1)+'</span>';
    var bg = i<3 ? colors[i] : '#fff';
    var rate = p.target>0 ? (p.done/p.target*100).toFixed(1)+'%' : '—';
    h += '<div style="display:flex;align-items:center;gap:12px;padding:10px 14px;background:'+bg+';border-radius:10px;border:1px solid #e2e8f0">'+
      '<div style="width:32px;text-align:center">'+rank+'</div>'+
      '<div style="flex:1"><div style="font-weight:700;font-size:14px">'+p.name+'</div><div style="font-size:11px;color:#64748b">'+p.group+'</div></div>'+
      '<div style="text-align:right"><div style="font-weight:700;font-size:16px;color:#1e293b">'+fmtNum(p.done)+'</div><div style="font-size:11px;color:#64748b">目标 '+fmtNum(p.target)+' | 完成率 '+rate+'</div></div>'+
    '</div>';
  });
  if(data.length === 0) h += '<div style="text-align:center;padding:20px;color:#94a3b8">该时间段暂无数据</div>';
  return h;
}

function showTop10(tabId, fromDay, toDay) {
  fromDay = fromDay || 1; toDay = toDay || 12;
  var tab = document.getElementById('tab-'+tabId);
  if(!tab) return;
  var title = tabId === 'online_person' ? '线上销售 TOP10' : '线下个案 TOP10';
  var isOnline = tabId === 'online_person';
  var M = window._M;
  var crowns = ['🥇','🥈','🥉'];
  var colors = ['#fef3c7','#f1f5f9','#ffedd5'];

  var top10 = buildRanking(M, M.person_daily || {}, fromDay, toDay, isOnline);
  var html = '<div id="top10-overlay" style="position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.4);z-index:9999;display:flex;align-items:center;justify-content:center" onclick="if(event.target===this)closeTop10()">'+
    '<div style="background:#fff;border-radius:16px;padding:24px;width:'+(isOnline?'520px':'min(95vw,1400px)')+';max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,0.3)">'+
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
    '</div>';

  if(!isOnline) {
    var top10G1 = buildRanking(M, M.person_daily || {}, fromDay, toDay, isOnline, false, '一组');
    var top10G2 = buildRanking(M, M.person_daily || {}, fromDay, toDay, isOnline, false, '二组');
    var top10G3 = buildRanking(M, M.person_daily || {}, fromDay, toDay, isOnline, false, '三组');
    var top10Jq = buildRanking(M, M.person_daily_junqun || {}, fromDay, toDay, isOnline, true);
    html += '<div style="display:flex;gap:10px">'+
      '<div style="flex:1;min-width:0"><h3 style="font-size:13px;color:#6366f1;margin:0 0 6px">📊 一组 <span style="font-size:11px;color:#94a3b8">普陀·徐汇·南京·广州</span></h3><div style="display:flex;flex-direction:column;gap:4px">'+renderRankList(top10G1, crowns, colors)+'</div></div>'+
      '<div style="flex:1;min-width:0"><h3 style="font-size:13px;color:#8b5cf6;margin:0 0 6px">📊 二组 <span style="font-size:11px;color:#94a3b8">深圳·杭州·武汉</span></h3><div style="display:flex;flex-direction:column;gap:4px">'+renderRankList(top10G2, crowns, colors)+'</div></div>'+
      '<div style="flex:1;min-width:0"><h3 style="font-size:13px;color:#ec4899;margin:0 0 6px">📊 三组 <span style="font-size:11px;color:#94a3b8">昆明·珠海·成都·福州</span></h3><div style="display:flex;flex-direction:column;gap:4px">'+renderRankList(top10G3, crowns, colors)+'</div></div>'+
      '<div style="flex:1;min-width:0"><h3 style="font-size:13px;color:#64748b;margin:0 0 6px">🧬 菌群业绩排名</h3><div style="display:flex;flex-direction:column;gap:4px">'+renderRankList(top10Jq, crowns, colors)+'</div></div>'+
      '</div>';
  } else {
    html += '<div style="display:flex;flex-direction:column;gap:8px">'+renderRankList(top10, crowns, colors)+'</div>';
  }

  html += '</div></div>';
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
  // 保存并隐藏按钮栏和工具栏
  var hiddenEls = document.querySelectorAll('.top-btns, .tab-toolbar');
  var saved = [];
  hiddenEls.forEach(function(el){ saved.push(el.style.display); el.style.display = 'none'; });
  // 展开按角色和按分组两个tab
  var role = document.querySelector('#tab-role');
  var group = document.querySelector('#tab-group');
  var tabs = document.querySelectorAll('.tab-content');
  tabs.forEach(function(t){ t.style.display = 'none'; });
  if (role) role.style.display = 'block';
  if (group) group.style.display = 'block';

  var target = document.querySelector('.container');
  html2canvas(target, {
    scale: 3,
    useCORS: true,
    backgroundColor: '#ffffff',
    logging: false
  }).then(function(canvas) {
    // 恢复
    hiddenEls.forEach(function(el, i){ el.style.display = saved[i]; });
    if (role) role.style.display = '';
    if (group) group.style.display = '';
    canvas.toBlob(function(blob) {
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = '南区7月目标快照_' + new Date().toISOString().slice(0,10) + '.png';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 'image/png');
  }).catch(function(err) {
    hiddenEls.forEach(function(el, i){ el.style.display = saved[i]; });
    if (role) role.style.display = '';
    if (group) group.style.display = '';
    alert('截图失败: ' + err.message);
  });
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
// ====== 康博嘉云数据导入 ======
// ====== 导入进度条（成功绿/失败红） ======
function ensureImportBar() {
  var box = document.getElementById('import-progress');
  if (box) return box;
  box = document.createElement('div');
  box.id = 'import-progress';
  box.style.cssText = 'position:fixed;top:16px;right:16px;z-index:9999;width:340px;background:#fff;border:1px solid #e2e8f0;border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.18);padding:12px 16px;font-size:13px;display:none';
  box.innerHTML = '<div id="imp-label" style="margin-bottom:8px;color:#334155;font-weight:600"></div>'+
    '<div style="background:#f1f5f9;border-radius:6px;height:8px;overflow:hidden"><div id="imp-bar" style="height:100%;width:0%;border-radius:6px;background:#3b82f6;transition:width .3s"></div></div>'+
    '<div id="imp-msg" style="margin-top:8px;color:#64748b;font-size:12px;white-space:pre-line"></div>';
  document.body.appendChild(box);
  return box;
}
var _impTimer = null;
function importProgress(label, pct) {
  var box = ensureImportBar();
  if (_impTimer) { clearTimeout(_impTimer); _impTimer = null; }
  box.style.display = 'block';
  document.getElementById('imp-label').textContent = '⏳ ' + label;
  var bar = document.getElementById('imp-bar');
  bar.style.background = '#3b82f6';
  bar.style.width = (pct || 0) + '%';
  document.getElementById('imp-msg').textContent = '';
}
function importDone(ok, label, msg) {
  var box = ensureImportBar();
  box.style.display = 'block';
  var bar = document.getElementById('imp-bar');
  bar.style.width = '100%';
  bar.style.background = ok ? '#22c55e' : '#ef4444';
  document.getElementById('imp-label').textContent = (ok ? '✅ ' : '❌ ') + label;
  document.getElementById('imp-msg').textContent = msg || '';
  if (_impTimer) clearTimeout(_impTimer);
  _impTimer = setTimeout(function(){ box.style.display = 'none'; }, ok ? 6000 : 15000);
}

var KBJ_STORE_MAP = {
  '观心正德诊所': '广州', '武汉观心精神专科门诊部': '武汉',
  '深圳观心正德综合门诊部': '深圳', '南京建邺观心综合门诊部': '南京',
  '杭州观心正念诊所有限公司': '杭州'
};

function importKangbojia(input) {
  var file = input.files[0]; if(!file) return;
  importProgress('导入康博嘉：'+file.name+'（读取文件…）', 15);
  var reader = new FileReader();
  reader.onerror = function(){ importDone(false, '康博嘉导入失败', '文件读取失败，请重试'); };
  reader.onload = function(e) {
    importProgress('导入康博嘉：'+file.name+'（解析中…）', 50);
    setTimeout(function(){
    var wb;
    try { wb = XLSX.read(e.target.result, {type:'array', cellDates: true}); }
    catch(err) { importDone(false, '康博嘉导入失败', '文件解析失败：'+err.message+'\\n请用 Excel 打开该文件另存为 .xlsx 后重新导入'); return; }
    // 自动选列数最多的 sheet（兼容带标题页的格式）
    var bestSheet = wb.SheetNames[0], bestCols = 0;
    wb.SheetNames.forEach(function(sn) {
      var s = wb.Sheets[sn];
      if(!s['!ref']) return;
      var range = XLSX.utils.decode_range(s['!ref']);
      if(range.e.c > bestCols) { bestCols = range.e.c; bestSheet = sn; }
    });
    var data = XLSX.utils.sheet_to_json(wb.Sheets[bestSheet]);
    var cols = data.length > 0 ? Object.keys(data[0]) : [];
    console.log('康博嘉导入: sheet='+bestSheet+' 行数='+data.length+' 列='+cols.join(','));

    // 查找列：名称精确匹配优先，位置索引兜底（防止首行空单元格导致列序漂移）
    var storeCol = (cols.indexOf('所在诊所')>=0 ? '所在诊所' : null) || cols.find(function(c){ return c.indexOf('诊所')>=0 || c.indexOf('机构')>=0 || c.indexOf('所在')>=0; }) || cols[47] || '';
    var cmCol = (cols.indexOf('客户经理')>=0 ? '客户经理' : null) || cols.find(function(c){ return c.indexOf('客户经理')>=0 || c.indexOf('经理')>=0; }) || cols[11] || '';
    var amtCol = (cols.indexOf('实际收入')>=0 ? '实际收入' : null) || cols.find(function(c){ return c.indexOf('实际收入')>=0 || c.indexOf('收入')>=0 || c.indexOf('金额')>=0; }) || cols[33] || '';
    if(!storeCol || !cmCol || !amtCol) {
      importDone(false, '康博嘉导入失败', '未识别到数据列。请确认文件包含：所在诊所、客户经理、实际收入'); return;
    }
    console.log('康博嘉导入: storeCol='+storeCol+' cmCol='+cmCol+' amtCol='+amtCol);

    // 按门店+客户经理汇总（仅南区5家）
    var storeTotal = {}, cmTotal = {};
    var matched = 0;
    data.forEach(function(row) {
      var rawStore = (row[storeCol] || '').toString().trim();
      var mapped = KBJ_STORE_MAP[rawStore];
      if(!mapped) return;
      var cm = (row[cmCol] || '').toString().trim();
      var amt = parseFloat(row[amtCol]) || 0;
      if(!cm || cm === 'nan' || cm === '测试医生') return;
      storeTotal[mapped] = (storeTotal[mapped] || 0) + amt;
      cmTotal[cm] = (cmTotal[cm] || 0) + amt;
      matched++;
    });
    console.log('康博嘉导入: 匹配'+matched+'条, 门店:', storeTotal, 'CM:', cmTotal);
    if(matched === 0) { importDone(false, '康博嘉导入失败', '未匹配到南区5家门店数据。请确认所在诊所列包含：\\n'+Object.keys(KBJ_STORE_MAP).join('、')); return; }

    // 存储到 _M
    window._M.kbj_store = storeTotal;
    window._M.kbj_cm = cmTotal;

    // 先回退上次导入的康博嘉增量（防止重复叠加）
    if(window._M._kbj_prev_store) {
      document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)').forEach(function(row) {
        var tds = row.querySelectorAll('td');
        if(tds.length < 12) return;
        var storeName = tds[0].textContent.replace('⚠锁定','').trim();
        var prev = window._M._kbj_prev_store[storeName] || 0;
        if(prev > 0) writeNum(tds[3], parseNum(tds[3]) - prev);
      });
      var cmTab2 = document.getElementById('tab-cm_person');
      if(cmTab2) {
        cmTab2.querySelectorAll('tbody tr:not(.subtotal)').forEach(function(row) {
          var tds = row.querySelectorAll('td');
          if(tds.length < 4) return;
          var name = tds[1].textContent.replace(/<[^>]+>/g,'').replace('🔗KBJ','').replace('⚠锁定','').trim();
          var prev = window._M._kbj_prev_cm[name] || 0;
          if(prev > 0) writeNum(tds[3], parseNum(tds[3]) - prev);
        });
      }
    }
    window._M._kbj_prev_store = storeTotal;
    window._M._kbj_prev_cm = cmTotal;

    // 更新 CM 表（叠加康博嘉业绩到对应CM）
    var cmTab = document.getElementById('tab-cm_person');
    if(cmTab) {
      cmTab.querySelectorAll('tbody tr:not(.subtotal)').forEach(function(row) {
        var tds = row.querySelectorAll('td');
        if(tds.length < 4) return;
        var name = tds[1].textContent.replace(/<[^>]+>/g,'').replace('🔗KBJ','').replace('⚠锁定','').trim();
        var extra = cmTotal[name] || 0;
        if(extra > 0) {
          writeNum(tds[3], parseNum(tds[3]) + extra);
          tds[1].innerHTML = name + ' <span style="font-size:10px;color:#059669">🔗KBJ</span>';
        }
      });
    }

    // 更新门店行（叠加康博嘉到对应门店）
    document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)').forEach(function(row) {
      var tds = row.querySelectorAll('td');
      if(tds.length < 12) return;
      var storeName = tds[0].textContent.replace('⚠锁定','').trim();
      var extra = storeTotal[storeName] || 0;
      if(extra > 0) writeNum(tds[3], parseNum(tds[3]) + extra);
    });

    // 全量刷新
    refreshAll();
    importDone(true, '康博嘉导入完成', '匹配 '+matched+' 条 · '+Object.keys(storeTotal).length+' 家门店 · '+Object.keys(cmTotal).length+' 位CM');
    }, 30);
  };
  reader.readAsArrayBuffer(file);
}

// ====== 普陀明细导入（康博嘉同格式） ======
function importPutuoDetail(input) {
  var file = input.files[0]; if(!file) return;
  importProgress('导入普陀明细：'+file.name+'（读取文件…）', 15);
  var reader = new FileReader();
  reader.onerror = function(){ importDone(false, '普陀明细导入失败', '文件读取失败，请重试'); };
  reader.onload = function(e) {
    importProgress('导入普陀明细：'+file.name+'（解析中…）', 50);
    setTimeout(function(){
    var wb;
    try { wb = XLSX.read(e.target.result, {type:'array', cellDates: true}); }
    catch(err) { importDone(false, '普陀明细导入失败', '文件解析失败：'+err.message+'\\n请用 Excel 打开该文件另存为 .xlsx 后重新导入'); return; }
    // 自动选列数最多的 sheet
    var bestSheet = wb.SheetNames[0], bestCols = 0;
    wb.SheetNames.forEach(function(sn) {
      var s = wb.Sheets[sn];
      if(!s['!ref']) return;
      var range = XLSX.utils.decode_range(s['!ref']);
      if(range.e.c > bestCols) { bestCols = range.e.c; bestSheet = sn; }
    });
    var data = XLSX.utils.sheet_to_json(wb.Sheets[bestSheet]);
    var cols = data.length > 0 ? Object.keys(data[0]) : [];
    // 表头不在第一行时（文件带标题行），自动向下查找含 CM姓名+收款金额 的表头行
    if(cols.indexOf('CM姓名')<0 || cols.indexOf('收款金额')<0) {
      var aoa = XLSX.utils.sheet_to_json(wb.Sheets[bestSheet], {header:1});
      for(var hi=0; hi<Math.min(10, aoa.length); hi++) {
        var vals = (aoa[hi]||[]).map(String);
        if(vals.indexOf('CM姓名')>=0 && vals.indexOf('收款金额')>=0) {
          data = XLSX.utils.sheet_to_json(wb.Sheets[bestSheet], {range: hi});
          cols = data.length > 0 ? Object.keys(data[0]) : [];
          console.log('普陀明细导入: 表头在第'+(hi+1)+'行');
          break;
        }
      }
    }
    console.log('普陀明细导入: sheet='+bestSheet+' 行数='+data.length);

    // 查找列：CM姓名(或客户经理)、收款金额(或实际收入)
    var cmCol = cols.find(function(c){ return c.indexOf('CM')>=0 || c.indexOf('客户经理')>=0 || c.indexOf('经理')>=0; }) || '';
    var amtCol = cols.find(function(c){ return c.indexOf('收款金额')>=0 || c.indexOf('实际收入')>=0 || c.indexOf('金额')>=0; }) || '';
    if(!cmCol || !amtCol) {
      importDone(false, '普陀明细导入失败', '未识别到数据列。请确认文件包含：CM姓名/客户经理、收款金额/实际收入'); return;
    }
    console.log('普陀明细导入: cmCol='+cmCol+' amtCol='+amtCol);

    // 先回退上次导入的普陀CM业绩（防止重复叠加）
    if(window._M._prev_putuo_cm) {
      var cmTabPre = document.getElementById('tab-cm_person');
      if(cmTabPre) {
        cmTabPre.querySelectorAll('tbody tr:not(.subtotal)').forEach(function(row) {
          var tds = row.querySelectorAll('td');
          if(tds.length < 4) return;
          if(tds[0].textContent.trim() !== '普陀') return;
          var pname = tds[1].textContent.replace('⚠锁定','').replace('🔗KBJ','').trim();
          var prev = window._M._prev_putuo_cm[pname] || 0;
          if(prev > 0) writeNum(tds[3], parseNum(tds[3]) - prev);
        });
      }
      // 回退分组表普陀行的上次明细汇总部分
      document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)').forEach(function(row) {
        var tds = row.querySelectorAll('td');
        if(tds.length < 12) return;
        if(tds[0].textContent.replace('⚠锁定','').trim() !== '普陀') return;
        writeNum(tds[3], parseNum(tds[3]) - (window._M._prev_putuo_total || 0));
      });
    }

    // 按CM汇总
    var cmDone = {}, total = 0;
    data.forEach(function(row) {
      var cm = (row[cmCol] || '').toString().trim();
      var amt = parseFloat(row[amtCol]) || 0;
      if(!cm || cm === 'nan') { total += amt; return; }  // 空CM归入门店总业绩
      cmDone[cm] = (cmDone[cm] || 0) + amt;
      total += amt;
    });
    console.log('普陀明细导入: 总='+total+' CM数='+Object.keys(cmDone).length, cmDone);

    // 存储到 _M（普陀总业绩 = 门诊明细 + 底表普陀订单）
    window._M.putuo_extra = cmDone;
    window._M._prev_putuo_cm = cmDone;
    window._M._prev_putuo_total = total;
    window._M.putuo_detail_total = total;
    window._M.putuo_locked = total + (window._M.putuo_base_orders || 0);

    // 更新 CM 表中普陀行（叠加门诊明细，保留底表订单部分）
    var cmTab = document.getElementById('tab-cm_person');
    if(cmTab) {
      cmTab.querySelectorAll('tbody tr:not(.subtotal)').forEach(function(row) {
        var tds = row.querySelectorAll('td');
        if(tds.length < 4) return;
        if(tds[0].textContent.trim() !== '普陀') return;
        var name = tds[1].textContent.replace('⚠锁定','').replace('🔗KBJ','').trim();
        var addDone = cmDone[name];
        if(addDone !== undefined) {
          writeNum(tds[3], parseNum(tds[3]) + addDone);
          tds[1].innerHTML = name + ' <span style="font-size:10px;color:#f59e0b">⚠锁定</span>';
        }
      });
    }

    // 更新分组表普陀行（门店总业绩联动）
    document.querySelectorAll('#tab-group tbody tr:not(.subtotal):not(.section-header)').forEach(function(row) {
      var tds = row.querySelectorAll('td');
      if(tds.length < 12) return;
      if(tds[0].textContent.replace('⚠锁定','').trim() !== '普陀') return;
      writeNum(tds[3], parseNum(tds[3]) + total);
      rebuildBar(tds[4], parseNum(tds[3]) / parseNum(tds[2]));
    });

    // 全量刷新
    refreshAll();
    importDone(true, '普陀明细导入完成', '总业绩 '+fmtNum(total)+' · CM数 '+Object.keys(cmDone).length+'\\n普陀门店行已联动更新（含底表订单部分）');
    }, 30);
  };
  reader.readAsArrayBuffer(file);
}

// ====== 组织名提取 (兼容 beisen_org_full_name) ======
function extractOrgFromBeisen(path) {
  if(!path) return '';
  var parts = String(path).split('_');
  var last = parts[parts.length-1];
  var direct = ['课程销售1组','课程销售2组','课程销售3组','课程销售4组','课程销售5组','课程销售6组',
    '班主任1组','班主任2组','班主任3组','班主任组','线上课程销售组','销售管理部'];
  if(direct.indexOf(last) >= 0) return last;
  var storeMap = {'上海普陀':'上海普陀店','上海普陀店':'上海普陀店','上海徐汇店':'上海徐汇店',
    '云南分公司':'云南分公司','南京分公司':'南京分公司','广州观心诊所':'广州观心诊所',
    '成都分公司':'成都分公司','杭州分公司':'杭州分公司','杭州观心诊所':'杭州分公司',
    '武汉观心诊所':'武汉观心诊所','深圳分公司':'深圳分公司','珠海分公司':'珠海分公司',
    '福州观心诊所':'福州观心诊所'};
  for(var i=0; i<parts.length; i++) {
    if(storeMap[parts[i]]) return storeMap[parts[i]];
  }
  return '';
}

// ====== 自动识别字段名 ======
function detectFields(cols) {
  var lower = cols.map(function(c){ return c.toLowerCase(); });
  function find(candidates) {
    for(var i=0; i<candidates.length; i++) {
      var idx = lower.indexOf(candidates[i].toLowerCase());
      if(idx >= 0) return cols[idx];
    }
    for(var i=0; i<candidates.length; i++) {
      for(var j=0; j<lower.length; j++) {
        if(lower[j].indexOf(candidates[i].toLowerCase()) >= 0) return cols[j];
      }
    }
    return null;
  }
  return {
    pay_at:   find(['pay_at','支付时间','付款时间','付款日期','订单日期']),
    pay_amt:  find(['pay_amount','支付金额','金额','实付金额','订单金额']),
    seller:   find(['seller_name','销售','销售姓名','销售员','负责人']),
    org:      find(['组织','beisen_org_full_name','beisen','部门','org_name','org']),
    back_at:  find(['back_at','退费时间','退款时间','退费日期','退款日期']),
    order_id: find(['order_id','订单号','订单ID'])
  };
}

// ====== 底表导入 & 全量重算 ======
function importBaseTable(input) {
  var file = input.files[0]; if(!file) return;
  var isCsv = /\.csv$/i.test(file.name);
  importProgress('导入底表：'+file.name+'（读取文件…）', 15);
  var reader = new FileReader();
  reader.onerror = function(){ importDone(false, '底表导入失败', '文件读取失败，请重试'); };
  reader.onload = function(e) {
    importProgress('导入底表：'+file.name+'（解析重算中…）', 50);
    setTimeout(function(){
    var wb;
    try {
      // CSV 用文本模式读取，避免中文乱码
      wb = XLSX.read(e.target.result, isCsv ? {type:'string'} : {type:'array', cellDates: true});
    } catch(err) {
      importDone(false, '底表导入失败', '文件解析失败：'+err.message+'\\n若文件损坏，请用 Excel 打开另存为 .xlsx 后重新导入'); return;
    }
    var sheet = wb.Sheets[wb.SheetNames[0]];
    var data = XLSX.utils.sheet_to_json(sheet);
    if(!data.length) { importDone(false, '底表导入失败', '未读到数据行，请检查文件内容'); return; }
    var M = window._M;
    var julyS = M.july_start, julyE = M.july_end, juneS = M.june_start, juneE = M.june_end;

    // 自动识别字段名
    var cols = data.length > 0 ? Object.keys(data[0]) : [];
    var F = detectFields(cols);
    console.log('导入: 识别字段', F);

    // 解析组织名：兼容 组织 和 beisen_org_full_name
    var sellerOvr = M.seller_org_override || {};
    data.forEach(function(row) {
      var rawOrg = row[F.org] || '';
      row._org = extractOrgFromBeisen(rawOrg);
      // 特殊归属规则（如余佼→徐汇）
      var snm = String(row[F.seller] || '').trim();
      if(sellerOvr[snm]) row._org = sellerOvr[snm];
    });

    // Store full raw orders for export
    window._rawOrders = data;

    // Rebuild person_daily from imported raw data
    var newPD = {};
    data.forEach(function(row) {
      var payAt = row[F.pay_at] ? new Date(row[F.pay_at]) : null;
      if(!payAt || payAt.getMonth()+1 !== 7) return;
      var name = row[F.seller] || '';
      if(!name) return;
      var day = payAt.getDate();
      var amt = parseFloat(row[F.pay_amt]) || 0;
      if(!newPD[name]) { newPD[name] = {}; newPD[name]._total = 0; newPD[name]._org = row._org || ''; }
      newPD[name][String(day)] = (newPD[name][String(day)] || 0) + amt;
      newPD[name]._total += amt;
    });
    window._M.person_daily = newPD;

    // Build group-level metrics from raw data
    var groupCalc = {};
    data.forEach(function(row) {
      var org = row._org || '';
      var g = M.org_to_group[org] || M.store_to_group[org] || null;
      if(!g) return;
      if(!groupCalc[g]) groupCalc[g] = {j7_inc:0, j6_inc:0, j7_ref:0, j6_ref:0};
      var pay = parseFloat(row[F.pay_amt]) || 0;
      var payAt = row[F.pay_at] ? new Date(row[F.pay_at]) : null;
      var backAt = row[F.back_at] ? new Date(row[F.back_at]) : null;
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
      var done = gc.j7_inc;
      if(gName === '普陀') {
        // 普陀 = 本次底表订单 + 门诊明细（明细沿用最近一次导入/生成值）
        M.putuo_base_orders = gc.j7_inc;
        done = gc.j7_inc + (M.putuo_detail_total || 0);
        M.putuo_locked = done;
      }
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

    // 底表重算已重写门店行，清除康博嘉叠加基线（防止旧基线导致重复回退）
    window._M._kbj_prev_store = null;
    window._M._kbj_prev_cm = null;

    // Full refresh for subtotals, KPIs, etc
    refreshAll();
    importDone(true, '底表导入完成', data.length+' 行订单已全量重算\\n康博嘉增量已重置，请接着导入最新康博嘉文件');
    }, 30);
  };
  if(isCsv) reader.readAsText(file, 'utf-8'); else reader.readAsArrayBuffer(file);
}

// ====== 渠道业绩 - 本地生活（大众点评推广通） ======
var ECO_LOADED = false;

function ecoMapStore(fullName) {
  var map = [
    ['上海旗舰','普陀'], ['徐汇','徐汇'], ['中环大厦','普陀'],
    ['南京','南京'], ['珠江新城','广州'], ['深圳福田','深圳'],
    ['杭州','杭州'], ['成都','成都'],
    ['云南','昆明'], ['珠海','珠海'], ['福州','福州']
  ];
  for (var i = 0; i < map.length; i++) {
    if (fullName.indexOf(map[i][0]) >= 0) return map[i][1];
  }
  return '';
}

function loadChannelData(force) {
  if (ECO_LOADED && !force) return;
  var tab = document.getElementById('tab-channel');
  if (!tab) return;

  var rows = window._ECO;
  if (!rows || rows.length === 0) {
    tab.innerHTML = '<div style="text-align:center;padding:40px;color:#ef4444">❌ 暂无本地生活数据<br><small style="color:#94a3b8">生成看板时 API 拉取失败，请重新运行 generate_dashboard.py</small></div>';
    return;
  }

  tab.innerHTML = '<div style="text-align:center;padding:40px;color:#64748b">⏳ 正在处理本地生活数据...</div>';
  // 用 setTimeout 让 loading 渲染出来
  setTimeout(function() {
    var storeData = {};
    rows.forEach(function(row) {
      var nm = ecoMapStore(row['门店'] || '');
      if (!nm) return;
      if (!storeData[nm]) storeData[nm] = {consume:0, consult:0, visit:0, deal:0, revenue:0, target:0};
      storeData[nm].consume += parseFloat(row['推广通消耗']) || 0;
      storeData[nm].consult += parseInt(row['咨询总数']) || 0;
      storeData[nm].visit   += parseInt(row['纯新到店人数']) || 0;
      storeData[nm].deal    += parseInt(row['成交人数']) || 0;
      storeData[nm].revenue += parseFloat(row['新客业绩']) || 0;
      storeData[nm].target  += parseFloat(row['目标']) || 0;
    });

    var M = window._M;
    var h = '<table><thead><tr>'+
      '<th>门店</th>'+
      '<th class="num">推广消耗</th><th class="num">咨询总数</th>'+
      '<th class="num">到店人数</th><th class="num">到店率</th>'+
      '<th class="num">成交人数</th><th class="num">新客业绩</th>'+
      '<th class="num">客单价</th><th class="num">ROI</th>'+
      '<th class="num">目标</th><th class="num">完成率</th>'+
      '</tr></thead><tbody>';

    var grand = {consume:0, consult:0, visit:0, deal:0, revenue:0, target:0};
    var hasAny = false;

    // 按原顺序展示所有南区门店
    var allStores = ['普陀','徐汇','南京','广州','深圳','杭州','武汉','昆明','珠海','成都','福州'];
    allStores.forEach(function(store) {
      var d = storeData[store];
      if (!d) return;
      hasAny = true;
      var vr = d.consult>0 ? (d.visit/d.consult*100).toFixed(1)+'%' : '—';
      var roi = d.consume>0 ? (d.revenue/d.consume).toFixed(2) : '—';
      var rate = d.target>0 ? (d.revenue/d.target*100).toFixed(1)+'%' : '—';
      var aov = d.deal>0 ? fmtNum(d.revenue/d.deal) : '—';

      h += '<tr>'+
        '<td>'+store+'</td>'+
        '<td class="num">'+fmtNum(d.consume)+'</td><td class="num">'+fmtInt(d.consult)+'</td>'+
        '<td class="num">'+fmtInt(d.visit)+'</td><td class="num">'+vr+'</td>'+
        '<td class="num">'+fmtInt(d.deal)+'</td><td class="num">'+fmtNum(d.revenue)+'</td>'+
        '<td class="num">'+aov+'</td><td class="num">'+roi+'</td>'+
        '<td class="num">'+fmtNum(d.target)+'</td><td class="num">'+rate+'</td>'+
        '</tr>';
      for (var k in grand) grand[k] += d[k];
    });

    // 顶部汇总卡片
    var summaryHtml = '';
    if (hasAny) {
      var gvr = grand.consult>0 ? (grand.visit/grand.consult*100).toFixed(1)+'%' : '—';
      var groi = grand.consume>0 ? (grand.revenue/grand.consume).toFixed(2) : '—';
      var grate = grand.target>0 ? (grand.revenue/grand.target*100).toFixed(1)+'%' : '—';
      var gaov = grand.deal>0 ? fmtNum(grand.revenue/grand.deal) : '—';
      var visitRate = grand.consult>0 ? (grand.visit/grand.consult*100).toFixed(1)+'%' : '—';
      var dealRate = grand.visit>0 ? (grand.deal/grand.visit*100).toFixed(1)+'%' : '—';

      summaryHtml = '<div style="display:flex;gap:12px;margin-bottom:16px;flex-wrap:wrap">'+
        '<div class="kpi-card" style="flex:1;min-width:120px"><div class="label">💰 推广消耗</div><div class="value">'+fmtNum(grand.consume)+'</div></div>'+
        '<div class="kpi-card" style="flex:1;min-width:120px"><div class="label">📞 咨询总数</div><div class="value">'+fmtInt(grand.consult)+'</div></div>'+
        '<div class="kpi-card" style="flex:1;min-width:120px"><div class="label">🚶 到店人数</div><div class="value">'+fmtInt(grand.visit)+'<div class="sub">到店率 '+visitRate+'</div></div></div>'+
        '<div class="kpi-card" style="flex:1;min-width:120px"><div class="label">💳 成交人数</div><div class="value">'+fmtInt(grand.deal)+'<div class="sub">转化率 '+dealRate+'</div></div></div>'+
        '<div class="kpi-card" style="flex:1;min-width:120px"><div class="label">📊 新客业绩</div><div class="value">'+fmtNum(grand.revenue)+'</div></div>'+
        '<div class="kpi-card" style="flex:1;min-width:120px"><div class="label">🎯 目标完成</div><div class="value">'+grate+'<div class="sub">ROI '+groi+' | 客单 '+gaov+'</div></div></div>'+
        '</div>';

      h += '<tr class="subtotal" style="background:#1e293b;color:#fff;font-weight:700">'+
        '<td>南区合计</td>'+
        '<td class="num">'+fmtNum(grand.consume)+'</td><td class="num">'+fmtInt(grand.consult)+'</td>'+
        '<td class="num">'+fmtInt(grand.visit)+'</td><td class="num">'+gvr+'</td>'+
        '<td class="num">'+fmtInt(grand.deal)+'</td><td class="num">'+fmtNum(grand.revenue)+'</td>'+
        '<td class="num">'+gaov+'</td><td class="num">'+groi+'</td>'+
        '<td class="num">'+fmtNum(grand.target)+'</td><td class="num">'+grate+'</td>'+
        '</tr>';
    } else {
      h += '<tr><td colspan="11" style="text-align:center;padding:20px;color:#94a3b8">暂无南区门店本地生活数据</td></tr>';
    }

    h += '</tbody></table>';
    h += '<div style="text-align:center;padding:8px;font-size:11px;color:#94a3b8;margin-top:8px">数据来源: 电商部本地生活 · 大众点评推广通 | 看板每2小时自动更新</div>';
    h = summaryHtml + h;
    tab.innerHTML = h;
    ECO_LOADED = true;
  }, 10);
}

// ====== 每日趋势图（Chart.js 版，带 hover 动效） ======
var _charts = {};
function drawOneChart(canvasId, roleData, rgb) {
  var canvas = document.getElementById(canvasId);
  if(!canvas || typeof Chart === 'undefined') return;

  var maxDay = 0;
  Object.keys(roleData||{}).map(Number).forEach(function(d){ if(d>maxDay) maxDay=d; });
  if(maxDay < 1) maxDay = 15;
  var labels = [], vals = [];
  for(var d=1; d<=maxDay; d++) { labels.push(d+'日'); vals.push(Math.round((roleData||{})[String(d)]||0)); }

  if(_charts[canvasId]) { _charts[canvasId].destroy(); }
  var ctx = canvas.getContext('2d');

  // 柱顶数值标签插件（快照可见）
  var labelPlugin = {
    id: 'barLabels',
    afterDatasetsDraw: function(chart) {
      var c = chart.ctx;
      var meta = chart.getDatasetMeta(1); // bar dataset
      if(!meta || meta.hidden) return;
      c.save();
      c.font = 'bold 9px sans-serif';
      c.fillStyle = 'rgb('+rgb+')';
      c.textAlign = 'center';
      meta.data.forEach(function(bar, i) {
        var v = vals[i];
        if(v <= 0) return;
        c.fillText((v/10000).toFixed(1), bar.x, bar.y - 6);
      });
      c.restore();
    }
  };

  // 渐变填充
  var grad = ctx.createLinearGradient(0, 0, 0, 200);
  grad.addColorStop(0, 'rgba('+rgb+',0.35)');
  grad.addColorStop(1, 'rgba('+rgb+',0.02)');

  _charts[canvasId] = new Chart(ctx, {
    type: 'bar',
    plugins: [labelPlugin],
    data: {
      labels: labels,
      datasets: [
        {
          type: 'line',
          label: '趋势',
          data: vals,
          borderColor: 'rgb('+rgb+')',
          backgroundColor: grad,
          borderWidth: 2.5,
          fill: true,
          tension: 0.4,
          pointRadius: 3,
          pointHoverRadius: 7,
          pointBackgroundColor: '#fff',
          pointBorderColor: 'rgb('+rgb+')',
          pointBorderWidth: 2,
          order: 0
        },
        {
          type: 'bar',
          label: '日业绩',
          data: vals,
          backgroundColor: 'rgba('+rgb+',0.55)',
          hoverBackgroundColor: 'rgb('+rgb+')',
          borderRadius: 6,
          barPercentage: 0.55,
          order: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      animation: { duration: 700, easing: 'easeOutQuart' },
      plugins: {
        legend: { display: false },
        tooltip: {
          backgroundColor: 'rgba(15,23,42,0.92)',
          titleFont: { size: 13, weight: 'bold' },
          bodyFont: { size: 13 },
          padding: 12,
          cornerRadius: 8,
          displayColors: false,
          callbacks: {
            title: function(items) { return items[0].label; },
            label: function(item) {
              var v = item.raw;
              return '业绩：' + (v/10000).toFixed(2) + ' 万';
            }
          }
        }
      },
      scales: {
        y: {
          beginAtZero: true,
          grace: '12%',
          grid: { color: '#f1f5f9', drawBorder: false },
          ticks: {
            color: '#94a3b8', font: { size: 10 },
            callback: function(v) { return (v/10000).toFixed(0)+'W'; }
          },
          title: { display: true, text: '业绩(W)', color: '#cbd5e1', font: { size: 10 } }
        },
        x: {
          grid: { display: false },
          ticks: { color: '#94a3b8', font: { size: 10 }, maxRotation: 0, autoSkip: false }
        }
      }
    }
  });
}

function drawTrendChart() {
  var M = window._M, rd = M.role_daily || {};
  drawOneChart('chart_sales', rd['销售'], '99,102,241');
  drawOneChart('chart_advisor', rd['班主任'], '16,185,129');
  drawOneChart('chart_store', rd['门店'], '245,158,11');
}

setTimeout(function() { drawTrendChart(); }, 400);

// ====== 分组迷你趋势线 (sparkline) + hover 大图 ======
function drawSparklines() {
  var M = window._M, gd = M.group_daily || {};
  document.querySelectorAll('canvas.spark').forEach(function(cv) {
    var g = cv.getAttribute('data-group');
    var data = gd[g] || {};
    var maxDay = 0;
    Object.keys(data).map(Number).forEach(function(d){ if(d>maxDay) maxDay=d; });
    if(maxDay < 1) maxDay = 15;
    var vals = [];
    for(var d=1; d<=maxDay; d++) vals.push(data[String(d)] || 0);
    var maxV = Math.max.apply(null, vals.concat([1]));
    var dpr = window.devicePixelRatio || 1;
    var W = 80, H = 26;
    cv.width = W*dpr; cv.height = H*dpr;
    var ctx = cv.getContext('2d');
    ctx.setTransform(dpr,0,0,dpr,0,0);
    ctx.clearRect(0,0,W,H);
    // 迷你面积+线
    var grad = ctx.createLinearGradient(0,0,0,H);
    grad.addColorStop(0,'rgba(99,102,241,0.3)'); grad.addColorStop(1,'rgba(99,102,241,0.02)');
    function px(i){ return 2 + i*((W-4)/(vals.length-1||1)); }
    function py(v){ return H-3 - (v/maxV)*(H-6); }
    // 面积
    ctx.beginPath(); ctx.moveTo(px(0), H-2);
    vals.forEach(function(v,i){ ctx.lineTo(px(i), py(v)); });
    ctx.lineTo(px(vals.length-1), H-2); ctx.closePath();
    ctx.fillStyle = grad; ctx.fill();
    // 线
    ctx.beginPath();
    vals.forEach(function(v,i){ i===0 ? ctx.moveTo(px(i),py(v)) : ctx.lineTo(px(i),py(v)); });
    ctx.strokeStyle = '#6366f1'; ctx.lineWidth = 1.5; ctx.stroke();
    // 末点
    var li = vals.length-1;
    ctx.beginPath(); ctx.arc(px(li), py(vals[li]), 2, 0, 2*Math.PI);
    ctx.fillStyle = '#6366f1'; ctx.fill();

    // hover 事件
    cv.onmouseenter = function(e){ showSparkPopup(g, data, maxDay, e); };
    cv.onmousemove = function(e){ moveSparkPopup(e); };
    cv.onmouseleave = function(){ hideSparkPopup(); };
  });
}

var _sparkPopup = null, _sparkChart = null;
function showSparkPopup(g, data, maxDay, e) {
  hideSparkPopup();
  _sparkPopup = document.createElement('div');
  _sparkPopup.style.cssText = 'position:fixed;z-index:9999;background:#fff;border:1px solid #e2e8f0;border-radius:12px;box-shadow:0 8px 32px rgba(0,0,0,0.16);padding:14px 16px;width:360px;pointer-events:none';
  _sparkPopup.innerHTML = '<div style="font-size:13px;font-weight:700;color:#1e293b;margin-bottom:8px">📈 '+g+' · 每日业绩趋势</div><div style="height:180px"><canvas id="_sparkBig"></canvas></div>';
  document.body.appendChild(_sparkPopup);
  moveSparkPopup(e);

  var labels = [], vals = [];
  for(var d=1; d<=maxDay; d++){ labels.push(d+'日'); vals.push(Math.round(data[String(d)]||0)); }
  var ctx = document.getElementById('_sparkBig').getContext('2d');
  var grad = ctx.createLinearGradient(0,0,0,180);
  grad.addColorStop(0,'rgba(99,102,241,0.35)'); grad.addColorStop(1,'rgba(99,102,241,0.02)');
  _sparkChart = new Chart(ctx, {
    type:'bar',
    plugins:[{id:'bl',afterDatasetsDraw:function(c){var cc=c.ctx,m=c.getDatasetMeta(1);if(!m)return;cc.save();cc.font='bold 9px sans-serif';cc.fillStyle='#6366f1';cc.textAlign='center';m.data.forEach(function(b,i){if(vals[i]>0)cc.fillText((vals[i]/10000).toFixed(1),b.x,b.y-5)});cc.restore();}}],
    data:{labels:labels,datasets:[
      {type:'line',data:vals,borderColor:'rgb(99,102,241)',backgroundColor:grad,borderWidth:2.5,fill:true,tension:0.4,pointRadius:2,order:0},
      {type:'bar',data:vals,backgroundColor:'rgba(99,102,241,0.5)',borderRadius:5,barPercentage:0.55,order:1}
    ]},
    options:{responsive:true,maintainAspectRatio:false,animation:{duration:300},
      plugins:{legend:{display:false},tooltip:{enabled:false}},
      scales:{y:{beginAtZero:true,grace:'14%',grid:{color:'#f1f5f9'},ticks:{color:'#94a3b8',font:{size:9},callback:function(v){return (v/10000).toFixed(0)+'W';}}},
        x:{grid:{display:false},ticks:{color:'#94a3b8',font:{size:9},maxRotation:0,autoSkip:false}}}}
  });
}
function moveSparkPopup(e) {
  if(!_sparkPopup) return;
  var x = e.clientX + 16, y = e.clientY + 16;
  if(x + 380 > window.innerWidth) x = e.clientX - 376;
  if(y + 230 > window.innerHeight) y = e.clientY - 226;
  _sparkPopup.style.left = x+'px'; _sparkPopup.style.top = y+'px';
}
function hideSparkPopup() {
  if(_sparkChart){ _sparkChart.destroy(); _sparkChart = null; }
  if(_sparkPopup){ _sparkPopup.remove(); _sparkPopup = null; }
}
setTimeout(drawSparklines, 500);

// ====== 周度业绩渲染 ======
function drawWeeklyRoles() {
  var M = window._M, wdefs = M.week_defs || [], rd = M.role_daily || {};
  var container = document.getElementById('weekly-role-container');
  if(!container || !wdefs.length) return;
  var roles = [
    {key:'销售', label:'销售', color:'#6366f1', bg:'#eef2ff'},
    {key:'班主任', label:'班主任', color:'#10b981', bg:'#ecfdf5'},
    {key:'门店', label:'门店', color:'#f59e0b', bg:'#fffbeb'}
  ];
  var html = '';
  roles.forEach(function(r){
    var data = rd[r.key] || {};
    var weekly = [];
    var maxV = 0;
    wdefs.forEach(function(w){
      var v = 0;
      for(var d=w[1]; d<=w[2]; d++) v += (data[String(d)] || 0);
      weekly.push({label: w[0].split(' ')[0], val: v});
      if(v > maxV) maxV = v;
    });
    if(maxV < 1) maxV = 100000;
    var bars = '';
    weekly.forEach(function(w){
      var pct = Math.min(w.val/maxV*100, 100);
      bars += '<div style="display:flex;flex-direction:column;align-items:center;gap:3px;flex:1;min-width:50px">'+
        '<span style="font-size:10px;color:#475569;font-weight:600">'+(w.val/10000).toFixed(1)+'万</span>'+
        '<div style="width:100%;height:48px;background:#f1f5f9;border-radius:6px;overflow:hidden;position:relative">'+
        '<div style="position:absolute;bottom:0;width:100%;height:'+pct+'%;background:'+r.color+';border-radius:6px;transition:height 0.4s"></div>'+
        '</div><span style="font-size:9px;color:#94a3b8;white-space:nowrap">'+w.label+'</span></div>';
    });
    html += '<div style="background:linear-gradient(135deg,'+r.bg+',#fff);padding:14px 16px;border-radius:12px;box-shadow:0 2px 8px rgba(0,0,0,0.04);border:1px solid #e2e8f0">'+
      '<div style="font-size:13px;font-weight:700;color:#1e293b;margin-bottom:10px">'+r.label+'</div>'+
      '<div style="display:flex;gap:8px;align-items:flex-end">'+bars+'</div></div>';
  });
  container.innerHTML = html;
}

function drawWeeklyGroups() {
  var M = window._M, wdefs = M.week_defs || [], gd = M.group_daily || {};
  document.querySelectorAll('.weekly-bar-row').forEach(function(div){
    var g = div.getAttribute('data-group');
    var data = gd[g] || {};
    var weekly = [];
    var maxV = 0;
    wdefs.forEach(function(w){
      var v = 0;
      for(var d=w[1]; d<=w[2]; d++) v += (data[String(d)] || 0);
      weekly.push({label: w[0].split(' ')[0], val: v});
      if(v > maxV) maxV = v;
    });
    if(maxV < 1) maxV = 50000;
    var bars = '';
    weekly.forEach(function(w){
      var pct = Math.min(w.val/maxV*100, 100);
      var barColor = w.val > 0 ? '#6366f1' : '#cbd5e1';
      bars += '<div style="display:flex;flex-direction:column;align-items:center;gap:2px;min-width:70px">'+
        '<span style="font-size:10px;color:#1e293b;font-weight:600">'+(w.val/10000).toFixed(1)+'万</span>'+
        '<div style="width:100%;height:14px;background:#f1f5f9;border-radius:7px;overflow:hidden;border:1px solid #e2e8f0">'+
        '<div style="width:'+pct+'%;height:100%;background:'+barColor+';border-radius:7px;transition:width 0.4s"></div>'+
        '</div><span style="font-size:9px;color:#94a3b8;white-space:nowrap">'+w.label+'</span></div>';
    });
    div.innerHTML = bars;
  });
}
setTimeout(function(){ drawWeeklyRoles(); drawWeeklyGroups(); }, 700);
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

# 分享版（iframe自动刷新）
share_html = '''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>南区7月目标进展看板</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: -apple-system, BlinkMacSystemFont, "PingFang SC", "Microsoft YaHei", sans-serif; background: #f1f5f9; }
.header { background: #1e293b; color: #fff; padding: 12px 24px; display: flex; justify-content: space-between; align-items: center; }
.header h1 { font-size: 16px; font-weight: 600; }
.header .info { font-size: 12px; color: #94a3b8; display: flex; align-items: center; gap: 8px; }
.header .dot { width: 8px; height: 8px; background: #22c55e; border-radius: 50%; animation: pulse 2s infinite; }
@keyframes pulse { 0%,100% { opacity: 1; } 50% { opacity: 0.3; } }
iframe { width: 100%; height: calc(100vh - 48px); border: none; display: block; }
</style>
</head>
<body>
<div class="header">
  <h1>📊 南区7月目标进展看板</h1>
  <div class="info"><span class="dot"></span> 实时看板 · 每12小时自动同步最新数据</div>
</div>
<iframe id="db" src="南区7月目标看板.html"></iframe>
<script>
var REFRESH=43200;
setInterval(function(){document.getElementById('db').src='南区7月目标看板.html?t='+Date.now()},REFRESH*1000);
</script>
</body>
</html>'''
with open(r'C:\Users\fanny\Desktop\数据日日\分享_南区看板.html', 'w', encoding='utf-8') as f:
    f.write(share_html)

# 同步到 Vercel 部署目录（自包含版：分享头 + 看板内容 + 自动刷新）
import shutil
deploy_dir = r'C:\Users\fanny\Desktop\数据日日\deploy'
os.makedirs(deploy_dir, exist_ok=True)

# 防缓存：在 <head> 末尾插入 meta 标签
_cache_buster = pd.Timestamp.now().strftime('%Y%m%d%H%M%S')
_cache_meta = f'<meta http-equiv="Cache-Control" content="no-cache, no-store, must-revalidate"><meta http-equiv="Pragma" content="no-cache"><meta http-equiv="Expires" content="0"><meta name="version" content="{_cache_buster}">'
# 自动刷新脚本（每2小时 + 首次加载强制拉最新）
_refresh_script = f'<script>var _v="{_cache_buster}";if(localStorage.getItem("_dv")!==_v){{localStorage.setItem("_dv",_v);location.reload(true)}}setInterval(function(){{location.reload(true)}},7200000)</script>'

# 在 head 末尾插入缓存控制
deploy_html = html.replace('</head>', _cache_meta + '</head>')
# 在 body 后插入 deploy 顶栏
deploy_header = '''<div style="background:#1e293b;color:#fff;padding:12px 24px;display:flex;justify-content:space-between;align-items:center;position:sticky;top:0;z-index:1000">
<h1 style="font-size:16px;font-weight:600;margin:0">📊 南区7月目标进展看板</h1>
<div style="font-size:12px;color:#94a3b8;display:flex;align-items:center;gap:8px"><span style="width:8px;height:8px;background:#22c55e;border-radius:50%;animation:pulse 2s infinite;display:inline-block"></span> 实时看板 · 每2小时自动刷新</div>
</div>
<style>.deploy-dot{animation:pulse 2s infinite}@keyframes pulse{0%,100%{opacity:1}50%{opacity:.3}}</style>
'''
deploy_html = deploy_html.replace('<body>', '<body>'+deploy_header).replace('</body>', _refresh_script+'</body>')
with open(os.path.join(deploy_dir, 'index.html'), 'w', encoding='utf-8') as f:
    f.write(deploy_html)

# 自动推送到 GitHub Pages
import subprocess
try:
    subprocess.run(['git', 'add', 'index.html'], cwd=deploy_dir, check=True, capture_output=True)
    subprocess.run(['git', 'commit', '-m', f'🔄 更新看板 {pd.Timestamp.now().strftime("%m/%d %H:%M")}'],
                   cwd=deploy_dir, capture_output=True)
    subprocess.run(['git', 'push', 'origin', 'main'], cwd=deploy_dir, check=True, capture_output=True)
    print('🚀 已推送 GitHub Pages: https://fannyqian.github.io/my-dashbord-for-guanxin/')
except Exception as e:
    print(f'⚠️ GitHub推送跳过: {e}')

print(f'✅ 看板已生成: {out}')
print(f'✅ 分享版: 分享_南区看板.html')
print(f'✅ 导出文件: 导出_按角色.xlsx / 导出_按分组.xlsx / 导出_线上个人.xlsx / 导出_线下个案.xlsx')
print(f'总目标: {fmt(total_metrics["target"])}')
print(f'完成: {fmt(total_metrics["done"])} ({pct2(total_metrics["rate"])})')
print(f'退费: {fmt(total_metrics["refund"])} (退费率 {pct2(safe_div(total_metrics["refund"],total_metrics["done"]))})')
print(f'净流水: {fmt(total_metrics["net"])} (环比 {pct1(total_metrics["net_chg"])})')
