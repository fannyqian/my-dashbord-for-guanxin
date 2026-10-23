"""Local-only preview for the second-stage 3980/4980 conversion funnel."""
from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
from pathlib import Path
from typing import Any

import pandas as pd

ROOT = Path(__file__).resolve().parent
DEFAULT_BASE = ROOT / "南区10月目标看板.html"
DEFAULT_INPUT = ROOT / "客户盘点" / "second_stage_4980_20261001_22"
DEFAULT_OUTPUT = ROOT / "南区10月目标看板_二阶4980本地预览.html"

REQUIRED_FILES = ("查询审计.json", "销售汇总.csv", "客户明细.csv", "订单审计明细.csv")
PROTECTED = (DEFAULT_BASE, ROOT / "daily_update_10.py", ROOT / "deploy" / "index.html")
QUEUE_NEW = "新增4980_窗口新分配"
QUEUE_HISTORY = "历史4980_10月前"


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def _id(value: Any) -> str:
    text = _text(value)
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def _num(value: Any) -> float:
    try:
        return float(value) if not pd.isna(value) else 0.0
    except (TypeError, ValueError):
        return 0.0


def _money(value: float) -> float:
    return round(float(value), 2)


def _rate(count: int, denominator: int) -> float:
    return round(count / denominator * 100, 1) if denominator else 0.0


def _entry_amounts(value: Any) -> set[int]:
    result: set[int] = set()
    for part in _text(value).split(";"):
        try:
            amount = round(float(part.strip()))
        except (TypeError, ValueError):
            continue
        if amount in (3980, 4980):
            result.add(int(amount))
    return result


def _prepare_frames(sales: pd.DataFrame, customers: pd.DataFrame, orders: pd.DataFrame):
    sales = sales.copy()
    customers = customers.copy()
    orders = orders.copy()
    for frame, columns in (
        (sales, ("销售", "销售ID", "归属", "组别", "匹配状态")),
        (customers, ("销售", "销售ID", "user_id", "队列", "归属", "组别")),
        (orders, ("销售", "销售ID", "user_id", "队列", "order_id", "seller_id")),
    ):
        for col in columns:
            if col not in frame.columns:
                frame[col] = ""
    for frame, columns in ((sales, ("销售ID",)), (customers, ("销售ID", "user_id")), (orders, ("销售ID", "user_id", "order_id", "seller_id"))):
        for col in columns:
            frame[col] = frame[col].map(_id)
    for col in ("窗口转化",):
        if col in customers:
            customers[col] = customers[col].map(lambda v: _text(v).lower() in ("true", "1", "yes"))
    for col in ("pay_amount", "refund_total", "order_amount", "actual_pay"):
        if col in orders:
            orders[col] = orders[col].map(_num)
        else:
            orders[col] = 0.0
    if "pay_at" not in orders:
        orders["pay_at"] = ""
    if "入口首单时间" not in customers:
        customers["入口首单时间"] = ""
    if "product_name" not in orders:
        orders["product_name"] = ""
    if "product_id" not in orders:
        orders["product_id"] = ""
    return sales, customers, orders


def _unique_orders(frame: pd.DataFrame) -> pd.DataFrame:
    if frame.empty:
        return frame.copy()
    x = frame.copy()
    x["order_id"] = x["order_id"].map(_id)
    x = x[x["order_id"] != ""]
    return x.drop_duplicates(subset=["order_id"], keep="first")


def _aggregate(frame: pd.DataFrame, denominator: int | None = None) -> dict[str, Any]:
    users = set(frame["user_id"].map(_id)) if not frame.empty else set()
    users.discard("")
    unique_orders = _unique_orders(frame)
    count = len(users)
    denominator = count if denominator is None else denominator
    return {
        "cohort": count,
        "up": count,
        "orders": int(len(unique_orders)),
        "amount": _money(unique_orders["pay_amount"].sum()) if not unique_orders.empty else 0.0,
        "refund": _money(unique_orders["refund_total"].sum()) if not unique_orders.empty else 0.0,
        "rate": _rate(count, denominator),
        "arpu": _money(unique_orders["pay_amount"].sum() / count) if count else 0.0,
    }


def _dimension_row(name: str, cohort_users: set[str], order_frame: pd.DataFrame, **extra: Any) -> dict[str, Any]:
    orders = _unique_orders(order_frame)
    converted = set(orders["user_id"].map(_id)) if not orders.empty else set()
    converted.discard("")
    row = {
        "name": name,
        "cohort": len(cohort_users),
        "up": len(converted),
        "orders": int(len(orders)),
        "amount": _money(orders["pay_amount"].sum()) if not orders.empty else 0.0,
        "refund": _money(orders["refund_total"].sum()) if not orders.empty else 0.0,
        "rate": _rate(len(converted), len(cohort_users)),
        "arpu": _money(orders["pay_amount"].sum() / len(converted)) if converted else 0.0,
    }
    row.update(extra)
    return row


def build_payload(audit: dict[str, Any], sales: pd.DataFrame, customers: pd.DataFrame, orders: pd.DataFrame) -> dict[str, Any]:
    sales, customers, orders = _prepare_frames(sales, customers, orders)
    orders = orders[orders.apply(lambda row: _id(row.get("seller_id")) == _id(row.get("销售ID")), axis=1)].copy()
    customer_users_all = customers.groupby("user_id", dropna=False)
    new_customers = customers[customers["队列"] == QUEUE_NEW].copy()
    history_customers = customers[customers["队列"] == QUEUE_HISTORY].copy()
    new_users = {_id(x) for x in new_customers["user_id"] if _id(x)}
    history_users = {_id(x) for x in history_customers["user_id"] if _id(x)}
    customer_users = new_customers.groupby("user_id", dropna=False)
    all_users = new_users
    conversion_orders_all = _unique_orders(orders)
    converted_users = {_id(x) for x in conversion_orders_all["user_id"] if _id(x)} & all_users
    conversion_orders = conversion_orders_all[conversion_orders_all["user_id"].map(_id).isin(converted_users)]
    history_converted_users = {_id(x) for x in conversion_orders_all["user_id"] if _id(x)} & history_users
    history_conversion_orders = conversion_orders_all[conversion_orders_all["user_id"].map(_id).isin(history_converted_users)]
    total_entry = len(new_users)
    total_converted = len(converted_users)
    total_amount = _money(conversion_orders["pay_amount"].sum()) if not conversion_orders.empty else 0.0
    total_refund = _money(conversion_orders["refund_total"].sum()) if not conversion_orders.empty else 0.0
    history_amount = _money(history_conversion_orders["pay_amount"].sum()) if not history_conversion_orders.empty else 0.0
    history_refund = _money(history_conversion_orders["refund_total"].sum()) if not history_conversion_orders.empty else 0.0
    history_cycle_days = []
    history_entry = history_customers[["user_id", "入口首单时间"]].copy()
    history_entry["user_id"] = history_entry["user_id"].map(_id)
    history_entry["entry_dt"] = pd.to_datetime(history_entry["入口首单时间"], errors="coerce")
    history_entry = history_entry.dropna(subset=["entry_dt"]).groupby("user_id", as_index=False)["entry_dt"].min()
    if not history_conversion_orders.empty:
        first_orders = history_conversion_orders.copy()
        first_orders["user_id"] = first_orders["user_id"].map(_id)
        first_orders["pay_dt"] = pd.to_datetime(first_orders["pay_at"], errors="coerce")
        first_orders = first_orders.dropna(subset=["pay_dt"]).sort_values(["user_id", "pay_dt", "order_id"]).drop_duplicates("user_id")
        cycles = first_orders.merge(history_entry, on="user_id", how="inner")
        history_cycle_days = ((cycles["pay_dt"] - cycles["entry_dt"]).dt.total_seconds() / 86400).tolist()
    history_avg_cycle_days = round(sum(history_cycle_days) / len(history_cycle_days), 1) if history_cycle_days else 0.0
    amounts_by_user: dict[str, set[int]] = {}
    for uid, group in customer_users:
        uid = _id(uid)
        if uid:
            values: set[int] = set()
            for value in group["入口金额"] if "入口金额" in group else []:
                values |= _entry_amounts(value)
            amounts_by_user[uid] = values
    entry_3980 = sum(3980 in values for values in amounts_by_user.values())
    entry_4980 = sum(4980 in values for values in amounts_by_user.values())

    def queue_row(queue: str) -> dict[str, Any]:
        queue_users = {_id(x) for x in new_customers.loc[new_customers["队列"] == queue, "user_id"] if _id(x)}
        frame = conversion_orders_all[conversion_orders_all["user_id"].isin(queue_users)]
        row = _dimension_row(queue, queue_users, frame)
        row["customers"] = row.pop("cohort")
        row["converted"] = row.pop("up")
        return row

    queue = [queue_row(QUEUE_NEW)]

    store_rows = []
    for group, subset in new_customers[new_customers["归属"] == "线下"].groupby("组别", dropna=False):
        name = _text(group) or "未映射"
        users = {_id(x) for x in subset["user_id"] if _id(x)}
        store_rows.append(_dimension_row(name, users, orders[orders["user_id"].isin(users)], belong="线下", store=name))
    store_rows.sort(key=lambda row: (-row["cohort"], row["name"]))

    cm_rows = []
    for (belong, group), subset in new_customers.groupby(["归属", "组别"], dropna=False):
        label = " / ".join(x for x in (_text(belong), _text(group)) if x) or "未映射"
        users = {_id(x) for x in subset["user_id"] if _id(x)}
        cm_rows.append(_dimension_row(label, users, orders[orders["user_id"].isin(users)], belong=_text(belong), group=_text(group), cm=label))
    cm_rows.sort(key=lambda row: (-row["cohort"], row["name"]))

    seller_rows = []
    for _, person in sales.iterrows():
        sid = _id(person.get("销售ID"))
        subset = new_customers[new_customers["销售ID"] == sid] if sid else new_customers.iloc[0:0]
        users = {_id(x) for x in subset["user_id"] if _id(x)}
        person_orders = orders[(orders["销售ID"] == sid) & (orders["user_id"].isin(all_users))] if sid else orders.iloc[0:0]
        row = _dimension_row(_text(person.get("销售")) or "未命名销售", users, person_orders,
                             seller_id=sid, belong=_text(person.get("归属")), group=_text(person.get("组别")),
                             status=_text(person.get("匹配状态")))
        seller_rows.append(row)
    seller_rows.sort(key=lambda row: (-row["up"], -row["cohort"], row["name"]))

    product_rows = []
    if not conversion_orders.empty:
        conversion_orders = conversion_orders.copy()
        conversion_orders["product"] = conversion_orders.apply(
            lambda row: _text(row.get("product_name")) or (_id(row.get("product_id")) or "未标注"), axis=1
        )
        for product, subset in conversion_orders.groupby("product", dropna=False):
            orders_unique = _unique_orders(subset)
            product_rows.append({
                "product": _text(product) or "未标注",
                "orders": int(len(orders_unique)),
                "users": int(orders_unique["user_id"].map(_id).nunique()),
                "amount": _money(orders_unique["pay_amount"].sum()),
                "refund": _money(orders_unique["refund_total"].sum()),
                "share": round(orders_unique["pay_amount"].sum() / total_amount * 100, 1) if total_amount else 0.0,
            })
    product_rows.sort(key=lambda row: (-row["amount"], row["product"]))

    window = audit.get("window") or ["", ""]
    return {
        "title": "二阶升单漏斗",
        "cohort_month": str(window[0])[:7] if window else "2026-10",
        "window": {"start": window[0] if len(window) > 0 else "", "end": window[1] if len(window) > 1 else "", "snapshot": audit.get("snapshot", "")},
        "stages": [
            {"label": "二阶本月新增3980/4980客户", "count": total_entry, "amount": None},
            {"label": "本月升单客户", "count": total_converted, "amount": total_amount, "rate_label": "升单率"},
        ],
        "overall": {
            "entry": {"count": total_entry, "new_count": len(new_users), "entry_3980_users": entry_3980, "entry_4980_users": entry_4980},
            "conversion": {"count": total_converted, "orders": int(len(conversion_orders)), "amount": total_amount, "refund": total_refund, "rate": _rate(total_converted, total_entry), "arpu": _money(total_amount / total_converted) if total_converted else 0.0},
            "history": {"count": len(history_users), "converted": len(history_converted_users), "orders": int(len(history_conversion_orders)), "amount": history_amount, "refund": history_refund, "rate": _rate(len(history_converted_users), len(history_users)), "arpu": _money(history_amount / len(history_converted_users)) if history_converted_users else 0.0, "avg_cycle_days": history_avg_cycle_days},
        },
        "queue": queue,
        "store": store_rows,
        "cm": cm_rows,
        "seller": seller_rows,
        "product": product_rows,
        "audit": {"staff_rows": audit.get("staff_rows"), "matched_staff": audit.get("unique_matched_staff"), "unmatched_or_ambiguous": audit.get("unmatched_or_ambiguous", []), "order_audit": audit.get("order_audit", {}), "customer_detail_rows": audit.get("customer_detail_rows"), "order_detail_rows": audit.get("order_detail_rows")},
        "limitations": ["仅支持本月新增入口客户→本月升单两段式漏斗", "无电话、加微、入群、D1-D3等中间行为数据", "无渠道字段和入口产品明细", "入口金额合并3980与4980，客户按user_id去重", "组别为名单快照维度；线下组别仅作为院区展示，二阶观鼎组不视为门店", "成交产品表仅代表成交侧产品分析"],
    }


def _html_table(title: str, headers: list[str], rows: list[list[str]]) -> str:
    def esc(value: Any) -> str:
        return html_escape(str(value))
    out = '<div class="f4980-subtitle">' + esc(title) + '</div><table><thead><tr>'
    out += ''.join('<th' + (' class="num"' if i else '') + '>' + esc(h) + '</th>' for i, h in enumerate(headers))
    out += '</tr></thead><tbody>'
    for row in rows:
        out += '<tr>' + ''.join('<td' + (' class="num"' if i else '') + '>' + esc(v) + '</td>' for i, v in enumerate(row)) + '</tr>'
    return out + '</tbody></table>' if rows else ''


def html_escape(value: str) -> str:
    return (value.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&#39;"))


def _js_section() -> str:
    return r'''function renderFunnel4980(f){
  if(!f||!f.overall) return '';
  var o=f.overall,e=o.entry,c=o.conversion,h='';
  function esc(x){return String(x==null?'':x).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
  function rate(x){return Number(x||0).toFixed(1)+'%';}
  function color(x){return x>=30?'#10b981':(x>=15?'#f59e0b':'#ef4444');}
  function money(x){return '¥'+(Number(x||0)/10000).toFixed(1)+'万';}
  function metric(label,value,colorValue){return _fMetric(label,value,colorValue);}
  h+='<div style="background:#fff;border-radius:14px;padding:18px 20px;box-shadow:0 1px 3px rgba(0,0,0,.06);border:2px solid #f59e0b;margin-bottom:16px">';
  h+='<div style="font-size:15px;font-weight:700;color:#1e293b;margin-bottom:6px">📈 '+esc(f.title)+'<span style="font-weight:400;color:#94a3b8;font-size:11px;margin-left:8px">本地预览 · 快照 '+esc(f.window.snapshot)+' · '+esc(f.window.start)+' 至 '+esc(f.window.end)+'（不含）</span></div>';
  h+='<div style="font-size:12px;line-height:1.7;color:#64748b;margin-bottom:14px">口径：仅统计本月新分配且购买过3980/4980入口的客户，3980与4980合并并按user_id去重；升单统计本月窗口内实际seller_id命中且pay_amount≥1000的后续成交，订单按order_id去重，金额为pay_amount毛额，退款单列；平均客单=升单金额÷升单客户数。</div>';
  h+='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px">';
  h+=metric('🎯 二阶本月新增3980/4980客户',fmtInt(e.count)+' 人','#6366f1');
  h+=metric('🚀 升单客户数',fmtInt(c.count)+' 人','#10b981');
  h+=metric('📊 升单率',rate(c.rate),color(c.rate));
  h+=metric('💰 升单金额',money(c.amount),'#ef4444');
  h+=metric('🧮 平均客单',money(c.arpu),'#8b5cf6');
  h+='</div>';
  h+='<div style="font-size:12px;color:#475569;margin:4px 0 10px">本月新增入口金额用户：3980 '+fmtInt(e.entry_3980_users)+' 人 · 4980 '+fmtInt(e.entry_4980_users)+' 人（两者可能重叠）</div>';
  if(f.overall.history){var hi=f.overall.history;h+='<div style="font-size:13px;font-weight:700;color:#334155;margin:18px 0 8px">🔄 历史4980客户本月升单（独立统计，不并入本月新增漏斗）</div><div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:12px">';h+=metric('历史4980客户',fmtInt(hi.count)+' 人','#64748b');h+=metric('本月升单客户',fmtInt(hi.converted)+' 人','#06b6d4');h+=metric('升单率',rate(hi.rate),color(hi.rate));h+=metric('升单金额',money(hi.amount),'#ef4444');h+=metric('平均升单周期',Number(hi.avg_cycle_days||0).toFixed(1)+'天','#0f766e');h+='</div>';}
  if(f.stages&&f.stages.length){var max=Math.max(f.stages[0].count,1);f.stages.forEach(function(s,i){var w=Math.min(100,s.count/max*100);h+='<div style="display:flex;align-items:center;gap:10px;margin:6px 0"><div style="width:180px;font-size:12px;color:#475569;text-align:right;flex-shrink:0">'+esc(s.label)+'</div><div style="flex:1;height:34px;background:#f1f5f9;border-radius:8px;position:relative;overflow:hidden"><div style="position:absolute;left:0;top:0;height:100%;width:'+w+'%;background:linear-gradient(90deg,#f59e0b,#ef4444);border-radius:8px;opacity:.3"></div><div style="position:relative;height:100%;display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:#1e293b">'+fmtInt(s.count)+' 人</div></div><div style="width:90px;font-size:12px;color:#64748b;text-align:right;flex-shrink:0">'+(s.amount==null?'—':money(s.amount))+'</div></div>';if(i<f.stages.length-1)h+='<div style="text-align:center;font-size:12px;color:#94a3b8;margin:3px 0 9px">↓ '+esc(s.rate_label||'转化率')+' <b style="color:'+color(c.rate)+'">'+rate(c.rate)+'</b></div>';});}
  if(f.queue&&f.queue.length)h+=_f4980Table('🧭 队列拆分',['队列','客户','成交','成交率','订单','成交金额','退款'],f.queue.map(function(r){return [r.name,fmtInt(r.customers),fmtInt(r.converted),rate(r.rate),fmtInt(r.orders),money(r.amount),money(r.refund)];}));
  if(f.store&&f.store.length)h+=_f4980Table('🏬 线下院区维度',['院区','入口','成交','成交率','订单','成交金额','退款'],f.store.map(function(r){return [r.name,fmtInt(r.cohort),fmtInt(r.up),rate(r.rate),fmtInt(r.orders),money(r.amount),money(r.refund)];}));
  if(f.cm&&f.cm.length)h+=_f4980Table('👤 近似CM/组别维度（非独立CM组织树）',['归属/组别','入口','成交','成交率','订单','成交金额','退款'],f.cm.map(function(r){return [r.name,fmtInt(r.cohort),fmtInt(r.up),rate(r.rate),fmtInt(r.orders),money(r.amount),money(r.refund)];}));
  if(f.seller&&f.seller.length)h+=_f4980Table('🧑 销售维度',['销售','归属/组别','入口','成交','成交率','订单','成交金额','退款'],f.seller.map(function(r){return [r.name,esc(r.belong+'/'+r.group),fmtInt(r.cohort),fmtInt(r.up),rate(r.rate),fmtInt(r.orders),money(r.amount),money(r.refund)];}));
  if(f.product&&f.product.length)h+=_f4980Table('📦 成交产品分析（不是入口产品漏斗）',['成交产品','订单','用户','成交金额','退款','金额占比'],f.product.map(function(r){return [r.product,fmtInt(r.orders),fmtInt(r.users),money(r.amount),money(r.refund),rate(r.share)];}));
  h+='<div style="font-size:12px;line-height:1.8;color:#64748b;margin-top:16px"><b>限制：</b>'+f.limitations.map(esc).join('；')+'</div></div>';return h;
}
function _f4980Table(title,headers,rows){if(!rows||!rows.length)return '';var h='<div style="font-size:13px;font-weight:700;color:#334155;margin:18px 0 8px">'+title+'</div><table><thead><tr>';headers.forEach(function(x,i){h+='<th'+(i?' class="num"':'')+'>'+x+'</th>';});h+='</tr></thead><tbody>';rows.forEach(function(r){h+='<tr>';r.forEach(function(x,i){h+='<td'+(i?' class="num"':'')+'>'+x+'</td>';});h+='</tr>';});return h+'</tbody></table>';}
'''


def build_preview(input_dir: Path = DEFAULT_INPUT, base: Path = DEFAULT_BASE, output: Path = DEFAULT_OUTPUT) -> Path:
    input_dir = input_dir.resolve()
    base = base.resolve()
    output = output.resolve()
    if output == base or output.parent.name.lower() == "deploy":
        raise ValueError("预览输出不能覆盖正式看板或写入 deploy")
    if not base.exists():
        raise FileNotFoundError(base)
    paths = {name: input_dir / name for name in REQUIRED_FILES}
    missing = [str(path) for path in paths.values() if not path.exists()]
    if missing:
        raise FileNotFoundError("缺少二阶结果文件: " + ", ".join(missing))
    protected_paths = tuple(path for path in (*PROTECTED, *paths.values()) if path.exists())
    before = {path: sha(path) for path in protected_paths}
    audit = json.loads(paths["查询审计.json"].read_text(encoding="utf-8-sig"))
    sales = pd.read_csv(paths["销售汇总.csv"], encoding="utf-8-sig")
    customers = pd.read_csv(paths["客户明细.csv"], encoding="utf-8-sig")
    orders = pd.read_csv(paths["订单审计明细.csv"], encoding="utf-8-sig")
    payload = build_payload(audit, sales, customers, orders)

    text = base.read_text(encoding="utf-8-sig")
    marker = re.search(r"window\._M\s*=\s*", text)
    if not marker:
        raise ValueError("正式看板缺少 window._M")
    dashboard, end = json.JSONDecoder().raw_decode(text[marker.end():])
    original = json.loads(json.dumps(dashboard, ensure_ascii=False))
    dashboard.setdefault("funnel", {})["funnel_4980"] = payload
    encoded = json.dumps(dashboard, ensure_ascii=False).replace("</", "<\\/")
    text = text[:marker.end()] + encoded + text[marker.end() + end:]
    old = "if(D.funnel_19){ h += renderFunnel19(D.funnel_19); }"
    new = old + "\n  if(D.funnel_4980){ h += renderFunnel4980(D.funnel_4980); }"
    if old not in text:
        raise ValueError("找不到现有 renderFunnel 注入锚点")
    text = text.replace(old, new, 1)
    insertion = _js_section()
    anchor = "function renderFunnel19(f){"
    if anchor not in text:
        raise ValueError("找不到 renderFunnel19 函数")
    text = text.replace(anchor, insertion + anchor, 1)
    text = text.replace("</body>", '<script>window.addEventListener("load",function(){switchTab("funnel");renderFunnel();});</script>\n</body>', 1)
    text = re.sub(r"<title>.*?</title>", "<title>本地预览｜二阶及线下3980/4980销转</title>", text, count=1)
    output.write_text(text, encoding="utf-8")
    assert dashboard["funnel"]["funnel_4980"] == payload
    for path, digest in before.items():
        if sha(path) != digest:
            raise AssertionError(f"保护文件发生变化: {path}")
    print(output)
    print(f"Verified: entry {payload['overall']['entry']['count']:,}; conversion {payload['overall']['conversion']['count']:,}/{payload['overall']['conversion']['orders']:,}; amount {payload['overall']['conversion']['amount']:,.2f}; local preview only")
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--preview", action="store_true", required=True, help="必须显式指定，仅生成本地预览")
    parser.add_argument("--input-dir", type=Path, default=DEFAULT_INPUT)
    parser.add_argument("--base", type=Path, default=DEFAULT_BASE)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()
    build_preview(args.input_dir, args.base, args.output)


if __name__ == "__main__":
    main()
