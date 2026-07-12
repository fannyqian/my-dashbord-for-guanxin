#!/usr/bin/env python3
"""生成可分享的快照HTML"""
import pandas as pd, numpy as np

orders = pd.read_excel(r'C:\Users\fanny\Desktop\数据日日\7月收入-青少年体验课订单明细.xlsx')
target_summary = pd.read_excel(r'C:\Users\fanny\Desktop\数据日日\南区-目标.xlsx', sheet_name='Sheet1')

ORG_TO_GROUP = {
    '课程销售1组':'课程1组','课程销售2组':'课程2组','课程销售3组':'课程3组',
    '课程销售4组':'销售4组','课程销售5组':'销售5组',
    '班主任2组':'班主任2组','班主任3组':'班主任3组',
}
STORE_MAP = {
    '上海普陀店':'普陀','上海徐汇店':'徐汇','云南分公司':'昆明',
    '南京分公司':'南京','广州观心诊所':'广州','成都分公司':'成都',
    '杭州分公司':'杭州','武汉观心诊所':'武汉','深圳分公司':'深圳',
    '珠海分公司':'珠海','福州观心诊所':'福州',
}
SALES_ORGS = ['课程销售1组','课程销售2组','课程销售3组','课程销售4组','课程销售5组']
ADVISOR_ORGS = ['班主任2组','班主任3组']

def classify_role(org):
    if org in SALES_ORGS: return '销售'
    if org in ADVISOR_ORGS: return '班主任'
    if org in STORE_MAP: return '门店'
    return '其他'

orders['role'] = orders['组织'].apply(classify_role)
orders['target_group'] = orders['组织'].map(ORG_TO_GROUP).fillna(orders['组织'].map(STORE_MAP))

target_map, cur = {}, None
for _, row in target_summary.iterrows():
    label = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
    g = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
    t = row.iloc[3]
    if g in ('分组','nan','') or g == '合计':
        if label in ('线上','线下门店'): cur = label
        continue
    if label in ('线上','线下门店'): cur = label
    try: target_map[g] = float(t)
    except: continue

online_tgt = sum(v for k,v in target_map.items() if k in ['课程1组','课程2组','课程3组','销售4组','销售5组','班主任2组','班主任3组'])
offline_tgt = sum(v for k,v in target_map.items() if k in STORE_MAP.values())

cm_df = pd.read_excel(r'C:\Users\fanny\Desktop\数据日日\南区-目标.xlsx', sheet_name='线下CM')
putuo_locked = sum(float(row.iloc[3]) for _, row in cm_df.iterrows() if str(row.iloc[0]).strip()=='普陀' and pd.notna(row.iloc[3]))

july_mask = (orders['pay_at']>='2026-07-01')&(orders['pay_at']<'2026-07-11')
june_mask = (orders['pay_at']>='2026-06-01')&(orders['pay_at']<'2026-06-11')

def parse_back(v):
    try:
        t=pd.to_datetime(v)
        return t if t.year>=2000 else pd.NaT
    except: return pd.NaT
orders['bd'] = orders['back_at'].apply(parse_back)
jref = (orders['bd']>='2026-07-01')&(orders['bd']<'2026-07-11')
jref_june = (orders['bd']>='2026-06-01')&(orders['bd']<'2026-06-11')

TP = 10/31
def sd(a,b): return a/b if b!=0 else 0
def fmt(n):
    if abs(n)>=1e8: return f'{n/1e8:.2f}亿'
    if abs(n)>=1e4: return f'{n/1e4:.0f}万'
    return f'{n:,.0f}'
def p2(n): return f'{n*100:.1f}%'
def p1(n): return f'{n*100:+.1f}%' if n and n!=float('inf') else '—'

# Role calc
role_data = []
for role in ['销售','班主任','门店']:
    grp = orders[orders['role']==role]
    tgt = sum(target_map.get(g,0) for g in grp['target_group'].dropna().unique())
    done = grp[july_mask]['pay_amount'].sum()
    prev = grp[june_mask]['pay_amount'].sum()
    ref = orders[orders['role']==role][jref]['pay_amount'].sum()
    prev_ref = orders[orders['role']==role][jref_june]['pay_amount'].sum()
    if role == '门店':
        po = orders[(orders['target_group']=='普陀')&july_mask]['pay_amount'].sum()
        done = done - po + putuo_locked
    rate = sd(done,tgt)
    net = done - ref; pnet = prev - prev_ref
    role_data.append({'role':role,'tgt':tgt,'done':done,'rate':rate,'chg':sd(done-prev,prev),
        'ref':ref,'rrate':sd(ref,done),'net':net,'pnet':pnet,'nchg':sd(net-pnet,pnet)})

total_tgt = sum(r['tgt'] for r in role_data)
total_done = sum(r['done'] for r in role_data)
total_ref = sum(r['ref'] for r in role_data)
total_net = sum(r['net'] for r in role_data)
total_rate = sd(total_done,total_tgt)
july_all = orders[orders['role'].isin(['销售','班主任','门店'])&july_mask]['pay_amount'].sum()
june_all = orders[orders['role'].isin(['销售','班主任','门店'])&june_mask]['pay_amount'].sum()
total_chg = sd(july_all-june_all,june_all)

# Group data
GROUPS = ['课程1组','课程2组','课程3组','销售4组','销售5组','班主任2组','班主任3组',
          '普陀','昆明','成都','广州','深圳','珠海','福州','杭州','南京','武汉','徐汇']
group_data = []
for g in GROUPS:
    tgt = target_map.get(g,0)
    if tgt == 0: continue
    ga = orders[orders['target_group']==g]
    done = ga[july_mask]['pay_amount'].sum()
    prev = ga[june_mask]['pay_amount'].sum()
    ref = ga[jref]['pay_amount'].sum()
    prev_ref = ga[jref_june]['pay_amount'].sum()
    if g == '普陀': done = putuo_locked
    rate = sd(done,tgt); net = done-ref; pnet = prev-prev_ref
    r = classify_role(orders[orders['target_group']==g]['组织'].iloc[0]) if len(orders[orders['target_group']==g])>0 else ''
    group_data.append({'group':g,'role':r,'tgt':tgt,'done':done,'rate':rate,'chg':sd(done-prev,prev),
        'ref':ref,'rrate':sd(ref,done),'rchg':sd(ref-prev_ref,prev_ref),'net':net,'pnet':pnet,'nchg':sd(net-pnet,pnet)})

def rate_bar(rate):
    p=min(rate*100,100)
    r=rate/TP if TP>0 else 0
    if r>=1: c,bg,bb,badge='#22c55e','#dcfce7','#f0fdf4',''
    elif r>=0.5: c,bg,bb,badge='#f59e0b','#fef3c7','#fffbeb','<span style="font-size:10px;background:#fef3c7;color:#b45309;padding:1px 6px;border-radius:8px">⚠️</span>'
    else: c,bg,bb,badge='#dc2626','#fecaca','#fef2f2','<span style="font-size:10px;background:#fee2e2;color:#b91c1c;padding:1px 6px;border-radius:8px">🚫</span>'
    return f'<div style="display:flex;align-items:center;gap:6px;min-width:180px"><div style="flex:1;height:20px;background:{bb};border-radius:10px;overflow:hidden;position:relative;border:1px solid #e2e8f0"><div style="width:{p}%;height:100%;background:{c};border-radius:10px"></div><div style="position:absolute;top:0;left:0;width:100%;height:100%;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:#1e293b">{p2(rate)}</div></div>{badge}</div>'

def ch(v):
    if v==0 or v==float('inf'): return '<span style="color:#94a3b8">—</span>'
    c='#22c55e' if v>=0 else '#ef4444'; a='↑' if v>=0 else '↓'
    return f'<span style="color:{c};font-weight:600">{a}{abs(v)*100:.1f}%</span>'

def rb(role):
    cl={'销售':('#dbeafe','#1e40af'),'班主任':('#fce7f3','#9d174d'),'门店':('#d1fae5','#065f46')}
    bg,fg=cl.get(role,('#f3f4f6','#6b7280'))
    return f'<span style="background:{bg};color:{fg};padding:2px 10px;border-radius:12px;font-size:12px;font-weight:600">{role}</span>'

# HTML
html = f'''<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>南区7月目标进展快照</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box}}
body{{font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif;background:#f1f5f9;color:#1e293b;padding:24px}}
.container{{max-width:1200px;margin:0 auto}}
.header{{text-align:center;margin-bottom:24px}}
.header h1{{font-size:24px;margin-bottom:4px}}
.header p{{color:#64748b;font-size:13px}}
.kpi-grid{{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:24px}}
.kpi-card{{background:#fff;border-radius:12px;padding:16px 20px;box-shadow:0 1px 3px rgba(0,0,0,0.06);text-align:center}}
.kpi-card .label{{font-size:12px;color:#64748b;margin-bottom:4px}}
.kpi-card .value{{font-size:28px;font-weight:700}}
.kpi-card .sub{{font-size:11px;color:#94a3b8;margin-top:4px}}
.legend{{display:flex;gap:16px;flex-wrap:wrap;align-items:center;margin-bottom:20px;padding:10px 16px;background:#fff;border-radius:10px;font-size:12px}}
.section{{background:#fff;border-radius:12px;padding:20px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,0.06)}}
.section h2{{font-size:16px;margin-bottom:12px;padding-bottom:8px;border-bottom:1px solid #e2e8f0}}
table{{width:100%;border-collapse:collapse;font-size:13px}}
th{{text-align:left;padding:10px 12px;border-bottom:2px solid #e2e8f0;color:#64748b;font-weight:600;font-size:12px;white-space:nowrap}}
td{{padding:7px 12px;border-bottom:1px solid #f1f5f9}}
.num{{text-align:right}}
.subtotal td{{background:#f8fafc;font-weight:700;border-top:1px solid #e2e8f0}}
.footer{{text-align:center;color:#94a3b8;font-size:11px;margin-top:24px}}
@media(max-width:768px){{.kpi-grid{{grid-template-columns:repeat(2,1fr)}}}}
</style></head><body><div class="container">
<div class="header"><h1>📊 南区7月目标进展快照</h1><p>数据截止：2026年7月10日 | 环比基期：6月1-10日同口径 | 普陀业绩单独统计</p></div>

<div class="kpi-grid">
<div class="kpi-card"><div class="label">📌 月总目标</div><div class="value">{fmt(total_tgt)}</div><div class="sub">线上 {fmt(online_tgt)} + 线下 {fmt(offline_tgt)}</div></div>
<div class="kpi-card"><div class="label">✅ 累计完成</div><div class="value">{fmt(total_done)}</div>{rate_bar(total_rate)}<div class="sub">完成率 {p2(total_rate)} | 环比 {p1(total_chg)}</div></div>
<div class="kpi-card"><div class="label">↩️ 本月退费</div><div class="value">{fmt(total_ref)}</div><div class="sub">退费率 {p2(sd(total_ref,total_done))} | 净流水 {fmt(total_net)}</div></div>
<div class="kpi-card"><div class="label">📈 时间进度</div><div class="value">{p2(TP)}</div><div class="sub">7月已过10/31天</div></div>
</div>

<div class="legend">
<span style="font-weight:700">📐 图例：</span>
<span style="display:inline-block;width:16px;height:10px;background:#22c55e;border-radius:5px"></span> 正常（≥{p2(TP)}）
<span style="display:inline-block;width:16px;height:10px;background:#f59e0b;border-radius:5px"></span> ⚠️ 关注（落后时间进度，可能完不成）
<span style="display:inline-block;width:16px;height:10px;background:#dc2626;border-radius:5px"></span> 🚫 高危（严重落后，需特别关注）
</div>

<div class="section"><h2>🏢 按角色汇总</h2><table>
<thead><tr><th>角色</th><th class="num">目标</th><th class="num">完成</th><th>完成率</th><th class="num">完成环比</th><th class="num">本月退费</th><th class="num">退费率</th><th class="num">本月净流水</th></tr></thead><tbody>
'''
for r in role_data:
    html += f'<tr><td>{rb(r["role"])}</td><td class="num">{fmt(r["tgt"])}</td><td class="num"><b>{fmt(r["done"])}</b></td><td>{rate_bar(r["rate"])}</td><td class="num">{ch(r["chg"])}</td><td class="num">{fmt(r["ref"])}</td><td class="num">{p2(r["rrate"])}</td><td class="num"><b>{fmt(r["net"])}</b></td></tr>'
html += f'''<tr class="subtotal"><td>南区合计</td><td class="num">{fmt(total_tgt)}</td><td class="num">{fmt(total_done)}</td><td>{rate_bar(total_rate)}</td><td class="num">{ch(total_chg)}</td><td class="num">{fmt(total_ref)}</td><td class="num">{p2(sd(total_ref,total_done))}</td><td class="num">{fmt(total_net)}</td></tr>
</tbody></table></div>

<div class="section"><h2>📋 按分组明细（按完成率升序）</h2><table>
<thead><tr><th>分组</th><th>角色</th><th class="num">目标</th><th class="num">完成</th><th>完成率</th><th class="num">完成环比</th><th class="num">本月退费</th><th class="num">退费率</th><th class="num">本月净流水</th></tr></thead><tbody>
'''
group_data.sort(key=lambda x: x['rate'])
for g in group_data:
    note = ' <span style="font-size:10px;color:#f59e0b">⚠</span>' if g['group']=='普陀' else ''
    html += f'<tr><td><b>{g["group"]}</b>{note}</td><td>{rb(g["role"])}</td><td class="num">{fmt(g["tgt"])}</td><td class="num"><b>{fmt(g["done"])}</b></td><td>{rate_bar(g["rate"])}</td><td class="num">{ch(g["chg"])}</td><td class="num">{fmt(g["ref"])}</td><td class="num">{p2(g["rrate"])}</td><td class="num"><b>{fmt(g["net"])}</b></td></tr>'

html += f'''</tbody></table></div>
<div class="footer">📸 南区7月目标进展快照 | {pd.Timestamp.now().strftime("%Y-%m-%d %H:%M")} | 观心运营助手</div>
</div></body></html>'''

out = r'C:\Users\fanny\Desktop\数据日日\南区7月目标快照.html'
with open(out, 'w', encoding='utf-8') as f:
    f.write(html)
print(f'快照已生成: {out}')
print(f'总目标:{fmt(total_tgt)} 完成:{fmt(total_done)}({p2(total_rate)}) 退费:{fmt(total_ref)} 净流水:{fmt(total_net)}')
