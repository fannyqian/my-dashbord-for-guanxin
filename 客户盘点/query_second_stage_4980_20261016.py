#!/usr/bin/env python3
"""Read-only second-stage 4980 allocation/conversion report for 2026-10-01..16."""
from collections import defaultdict
from pathlib import Path
from decimal import Decimal
import argparse, json, sys, hashlib
import pandas as pd
from openpyxl import load_workbook

WORK = Path(__file__).resolve().parent
sys.path.insert(0, str(WORK))
import _run_roster_query as source

DEFAULT_START = '2026-10-01'
DEFAULT_END = '2026-10-16'


def runtime_config():
    parser = argparse.ArgumentParser(description='查询二阶及线下销售4980新增/历史销转')
    parser.add_argument('--start', default=DEFAULT_START, help='窗口起始日期，含当天，格式YYYY-MM-DD')
    parser.add_argument('--end', default=DEFAULT_END, help='窗口结束日期，含当天，格式YYYY-MM-DD')
    parser.add_argument('--snapshot', default=None, help='DLC快照分区YYYYMMDD；默认自动取四表共同最新分区')
    args = parser.parse_args()
    start = pd.Timestamp(args.start).normalize()
    end = pd.Timestamp(args.end).normalize()
    if end < start:
        parser.error('--end must be on or after --start')
    end_exclusive = end + pd.Timedelta(days=1)
    return args, start, end, end_exclusive


def snapshot_partition(dlc):
    sql = """select max(dt) dt from (
      select dt from guanxin_dim.dim_order_clue_allocation_clue_info_df
      intersect select dt from guanxin_dim.dim_order_clue_cooperator_clue_df
      intersect select dt from guanxin_dwd.dwd_order_order_detail_df
      intersect select dt from guanxin_dim.dim_participant_partner_assistant_df
    ) x"""
    rows = source.cached_query(dlc, 'second_stage_common_partition', sql, retries=3)
    if not rows or not rows[0].get('dt'):
        raise RuntimeError('无法取得四张核心表共同最新分区')
    return str(rows[0]['dt'])


def q(dlc, label, sql):
    rows = source.cached_query(dlc, 'second_stage_' + label, sql, retries=3)
    return rows


def n(v):
    try: return float(v or 0)
    except (TypeError, ValueError): return 0.0


def ts(v):
    x = pd.to_datetime(v, errors='coerce', format='mixed')
    if pd.isna(x): return None
    if getattr(x, 'tzinfo', None) is not None: x = x.tz_convert('Asia/Shanghai').tz_localize(None)
    return x


def sql_strings(xs):
    return ','.join("'" + str(x).replace("'", "''") + "'" for x in xs)


def canonical_orders(rows):
    by = {}
    for r in rows:
        uid, oid = str(r.get('uid') or ''), str(r.get('order_id') or '')
        if not uid or not oid: continue
        key = (uid, oid)
        if key not in by: by[key] = dict(r)
        else:
            # Keep the row with the most complete monetary/identity fields.
            old = by[key]
            score = sum(bool(old.get(k)) for k in ('pay_at','seller_id','actual_pay','pay_amount','main_order_id'))
            newscore = sum(bool(r.get(k)) for k in ('pay_at','seller_id','actual_pay','pay_amount','main_order_id'))
            if newscore > score: by[key] = dict(r)
    children = defaultdict(list)
    for key, r in by.items():
        parent = str(r.get('main_order_id') or '').strip()
        if parent and parent not in ('0','None','null'): children[(key[0], parent)].append(key)
    excluded, missing_parent = set(), []
    for (uid, parent), keys in children.items():
        if (uid, parent) in by: excluded.update(keys)
        else: missing_parent.append({'uid': uid, 'main_order_id': parent, 'child_order_ids': [k[1] for k in keys]})
    out = {}
    for key, r in by.items():
        if key in excluded: continue
        x = dict(r)
        x['pay_dt'] = ts(x.get('pay_at'))
        x['entry_amount'] = n(x.get('actual_pay')) if x.get('actual_pay') not in (None, '') else n(x.get('pay_amount'))
        x['pay_amount_num'] = n(x.get('pay_amount'))
        x['refund_num'] = n(x.get('refund_total'))
        x['valid_date'] = x['pay_dt'] is not None
        x['amount_conflict'] = len({round(n(x.get(k)), 2) for k in ('actual_pay','pay_amount','order_amount') if x.get(k) not in (None, '')}) > 1
        x['is_4980'] = round(x['entry_amount'], 2) in (3980, 4980)
        out[key] = x
    return out, {'raw_rows': len(rows), 'unique_user_orders': len(by), 'parent_child_rows_excluded': len(excluded), 'missing_parent_groups': missing_parent, 'canonical_orders': len(out)}


def main():
    global DT, START, END, OUT, STATUS
    args, start_dt, end_dt, end_exclusive = runtime_config()
    dlc = source.dashboard.DLC()
    DT = args.snapshot or snapshot_partition(dlc)
    START = start_dt.strftime('%Y-%m-%d 00:00:00')
    END = end_exclusive.strftime('%Y-%m-%d 00:00:00')
    OUT = WORK / f"second_stage_4980_{start_dt:%Y%m%d}_{end_dt:%d}"
    OUT.mkdir(exist_ok=True)
    STATUS = f"已覆盖至{DT}快照"
    book = WORK / '一阶线索分配给二阶和线下销售计划表.xlsx'
    ws = load_workbook(book, data_only=True, read_only=True)['二阶+线下销售']
    roster = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        name, belong, group = [str(row[i] or '').strip() for i in (0,1,2)]
        if name and name not in ('None','销售'):
            roster.append({'name': name, 'belong': belong, 'group': group})
    names = sorted({r['name'] for r in roster})
    ids = sql_strings(names)
    staff_sql = f"""
      select cast(id as varchar) staff_id, trim(realname) staff_name,
             coalesce(trim(org_full_name),'') org_full_name,
             coalesce(trim(assistant_group_name),'') assistant_group_name,
             coalesce(cast(is_online as varchar),'') is_online
      from guanxin_dim.dim_participant_partner_assistant_df
      where dt='{DT}' and trim(realname) in ({ids})
    """
    staff_rows = q(dlc, 'staff', staff_sql)
    byname = defaultdict(list)
    for r in staff_rows: byname[str(r.get('staff_name') or '').strip()].append(r)
    # Match workbook context when possible; duplicate names remain explicitly unresolved.
    staff = []
    for r in roster:
        hits = byname.get(r['name'], [])
        if len(hits) == 1:
            staff.append({**r, 'staff_id': str(hits[0]['staff_id']), 'match_status': '唯一匹配', **hits[0]})
        elif not hits:
            staff.append({**r, 'staff_id': None, 'match_status': '未匹配'})
        else:
            staff.append({**r, 'staff_id': None, 'match_status': '重名', 'candidate_ids': ','.join(sorted({str(x['staff_id']) for x in hits}))})
    usable = [r for r in staff if r.get('staff_id')]
    usable_ids = sql_strings({r['staff_id'] for r in usable}) or "''"
    pool_sql = f"""
      with staff as (
        select cast(id as varchar) staff_id, trim(realname) staff_name
        from guanxin_dim.dim_participant_partner_assistant_df
        where dt='{DT}' and cast(id as varchar) in ({usable_ids})
      ), own as (
        select distinct cast(a.clue_assistant_id as varchar) staff_id,
          cast(a.clue_user_id as varchar) uid, '归属' ownership_type,
          cast(a.clue_created_at as varchar) clue_created_at,
          cast(a.clue_updated_at as varchar) clue_updated_at,
          cast(a.clue_status as varchar) clue_status
        from guanxin_dim.dim_order_clue_allocation_clue_info_df a join staff s
          on cast(a.clue_assistant_id as varchar)=s.staff_id
        where a.dt='{DT}' and a.clue_user_id is not null
      ), collab as (
        select distinct cast(c.cooperator_id as varchar) staff_id,
          cast(c.user_id as varchar) uid, '协作' ownership_type,
          cast(null as varchar) clue_created_at, cast(null as varchar) clue_updated_at,
          cast(null as varchar) clue_status
        from guanxin_dim.dim_order_clue_cooperator_clue_df c join staff s
          on cast(c.cooperator_id as varchar)=s.staff_id
        where c.dt='{DT}' and c.user_id is not null and c.cooperator_id>0
      ) select * from own union all select * from collab
    """
    pool_rows = q(dlc, 'pool', pool_sql)
    order_sql = f"""
      with scope as (
        select distinct cast(a.clue_user_id as varchar) uid
        from guanxin_dim.dim_order_clue_allocation_clue_info_df a
        where a.dt='{DT}' and a.clue_user_id is not null
          and cast(a.clue_assistant_id as varchar) in ({usable_ids})
        union
        select distinct cast(c.user_id as varchar) uid
        from guanxin_dim.dim_order_clue_cooperator_clue_df c
        where c.dt='{DT}' and c.user_id is not null and c.cooperator_id>0
          and cast(c.cooperator_id as varchar) in ({usable_ids})
      )
      select cast(o.user_id as varchar) uid, cast(o.order_id as varchar) order_id,
        cast(o.main_order_id as varchar) main_order_id, cast(o.pay_at as varchar) pay_at,
        cast(o.product_id as varchar) product_id, coalesce(o.product_name,'') product_name,
        coalesce(o.service_type_name,'') service_type_name,
        cast(o.actual_pay as varchar) actual_pay, cast(o.pay_amount as varchar) pay_amount,
        cast(o.order_amount as varchar) order_amount, cast(o.refund_total as varchar) refund_total,
        cast(o.seller_id as varchar) seller_id, cast(o.assistant_id as varchar) assistant_id,
        coalesce(o.order_status_desc,'') order_status_desc
      from guanxin_dwd.dwd_order_order_detail_df o join scope s
        on cast(o.user_id as varchar)=s.uid
      where o.dt='{DT}'
        and o.order_status_desc in ('已支付','全部退款','部分退款','部分收款','退款中','已发货','已完成')
        and o.pay_at is not null
    """
    order_rows = q(dlc, 'orders', order_sql)
    orders, order_audit = canonical_orders(order_rows)
    byuid = defaultdict(list)
    for x in orders.values(): byuid[str(x['uid'])].append(x)
    for xs in byuid.values(): xs.sort(key=lambda x: (x['pay_dt'] or pd.Timestamp.max, x['order_id']))
    staff_by_id = {r['staff_id']: r for r in usable}
    pool = defaultdict(list)
    for r in pool_rows:
        if r.get('uid') and r.get('staff_id') in staff_by_id: pool[(str(r['staff_id']), str(r['uid']))].append(r)
    customers, order_details = [], []
    excluded_historical_5000 = set()
    summary = []
    for (sid, uid), rels in sorted(pool.items()):
        own = '归属+协作' if {str(x.get('ownership_type')) for x in rels} == {'归属','协作'} else str(rels[0].get('ownership_type'))
        entries = [o for o in byuid.get(uid, []) if o['is_4980'] and o['valid_date']]
        entry_ids = {o['order_id'] for o in entries}
        entry_start = min((o['pay_dt'] for o in entries), default=None)
        historical_5000 = any(
            o['valid_date'] and o['order_id'] not in entry_ids and o['pay_amount_num'] >= 1000
            and entry_start is not None and o['pay_dt'] > entry_start and o['pay_dt'] < pd.Timestamp(START)
            for o in byuid.get(uid, [])
        )
        if historical_5000:
            excluded_historical_5000.add(uid)
            continue
        before = [o for o in entries if o['pay_dt'] < pd.Timestamp(START)]
        window_alloc = [x for x in rels if START <= str(x.get('clue_created_at') or '') < END]
        queues = []
        if window_alloc and entries: queues.append(('新增4980_窗口新分配', window_alloc, entries))
        if before: queues.append(('历史4980_10月前', rels, before))
        for queue, allocs, anchors in queues:
            eligible = []
            for o in byuid.get(uid, []):
                if not o['valid_date'] or not (pd.Timestamp(START) <= o['pay_dt'] < pd.Timestamp(END)): continue
                if o['pay_amount_num'] < 1000: continue
                eligible.append(o)
            # One customer row; order IDs are separately auditable.
            customers.append({'销售': staff_by_id[sid]['name'], '销售ID': sid, '归属': staff_by_id[sid]['belong'], '组别': staff_by_id[sid]['group'], 'user_id': uid, '队列': queue, '客户归属类型': own, '分配时间': min([str(x.get('clue_created_at')) for x in allocs if x.get('clue_created_at')] or ['']), '入口订单数': len(anchors), '入口首单时间': str(min(o['pay_dt'] for o in anchors)), '入口末单时间': str(max(o['pay_dt'] for o in anchors)), '入口金额': ';'.join(str(o['entry_amount']) for o in anchors), '窗口转化': bool(eligible), '窗口转化订单数': len(eligible), '窗口转化业绩_pay_amount': sum(o['pay_amount_num'] for o in eligible), '窗口退款': sum(o['refund_num'] for o in eligible), '实际seller_ids': ';'.join(sorted({str(o.get('seller_id') or '') for o in eligible})), '承接assistant_ids': ';'.join(sorted({str(o.get('assistant_id') or '') for o in eligible}))})
            for o in eligible:
                order_details.append({'销售': staff_by_id[sid]['name'], '销售ID': sid, 'user_id': uid, '队列': queue, 'order_id': o['order_id'], 'main_order_id': o.get('main_order_id'), 'pay_at': str(o['pay_dt']), 'order_amount': o.get('order_amount'), 'pay_amount': o.get('pay_amount'), 'actual_pay': o.get('actual_pay'), 'refund_total': o.get('refund_total'), 'product_id': o.get('product_id'), 'product_name': o.get('product_name'), 'seller_id': o.get('seller_id'), 'assistant_id': o.get('assistant_id'), 'cooperator_id': o.get('cooperator_id'), 'amount_conflict': o['amount_conflict'], 'is_4980_entry': o['is_4980'], 'seller_matches_customer': True})
    cdf = pd.DataFrame(customers)
    odf = pd.DataFrame(order_details)
    for r in staff:
        sid = r.get('staff_id'); name = r['name']
        out = {'销售': name, '销售ID': sid or '', '归属': r['belong'], '组别': r['group'], '匹配状态': r['match_status']}
        for queue, prefix in [('新增4980_窗口新分配','新增4980'), ('历史4980_10月前','历史4980')]:
            sub = cdf[(cdf['销售ID'].astype(str)==str(sid)) & (cdf['队列']==queue)] if sid and not cdf.empty else pd.DataFrame()
            osub = odf[(odf['销售ID'].astype(str)==str(sid)) & (odf['队列']==queue)] if sid and not odf.empty else pd.DataFrame()
            out[prefix+'客户数'] = len(sub)
            out[prefix+'转化人数'] = int(sub['窗口转化'].sum()) if not sub.empty else 0
            out[prefix+'转化订单数'] = len(osub)
            out[prefix+'转化业绩_pay_amount'] = float(osub['pay_amount'].map(n).sum()) if not osub.empty else 0.0
            out[prefix+'退款'] = float(osub['refund_total'].map(n).sum()) if not osub.empty else 0.0
        out['状态'] = STATUS if r['match_status']=='唯一匹配' else r['match_status']
        summary.append(out)
    pd.DataFrame(summary).to_csv(OUT/'销售汇总.csv', index=False, encoding='utf-8-sig')
    cdf.to_csv(OUT/'客户明细.csv', index=False, encoding='utf-8-sig')
    odf.to_csv(OUT/'订单审计明细.csv', index=False, encoding='utf-8-sig')
    report = {'snapshot': DT, 'window': [START, END], 'status': STATUS, 'staff_rows': len(staff), 'unique_matched_staff': len(usable), 'unmatched_or_ambiguous': [r for r in staff if not r.get('staff_id')], 'pool_rows': len(pool_rows), 'pool_sales_customers': len(pool), 'unique_pool_users': len({uid for _,uid in pool}), 'excluded_historical_1000_users': len(excluded_historical_5000), 'order_audit': order_audit, 'customer_detail_rows': len(cdf), 'order_detail_rows': len(odf), 'checks': {'window_boundary': f'half-open [{START}, {END})', 'conversion_uses_seller_id': True, 'amount_uses_pay_amount': True, 'refund_separate': True, 'personal_data_exported': False}}
    (OUT/'查询审计.json').write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding='utf-8')
    print(json.dumps(report, ensure_ascii=False, indent=2, default=str))

if __name__ == '__main__': main()
