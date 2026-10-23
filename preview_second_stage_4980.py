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


def _metric_block(users: set[str], orders: pd.DataFrame) -> dict[str, Any]:
    users = {_id(x) for x in users if _id(x)}
    unique = _unique_orders(orders[orders["user_id"].map(_id).isin(users)]) if not orders.empty else orders.copy()
    converted = {_id(x) for x in unique["user_id"] if _id(x)}
    amount = _money(unique["pay_amount"].sum()) if not unique.empty else 0.0
    refund = _money(unique["refund_total"].sum()) if not unique.empty else 0.0
    return {
        "cohort_users": len(users),
        "converted_users": len(converted),
        "rate": _rate(len(converted), len(users)),
        "orders": int(len(unique)),
        "amount": amount,
        "refund": refund,
        "arpu": _money(amount / len(converted)) if converted else 0.0,
    }


def _metric_legacy(block: dict[str, Any]) -> dict[str, Any]:
    return {"cohort": block["cohort_users"], "up": block["converted_users"], "rate": block["rate"], "orders": block["orders"], "amount": block["amount"], "refund": block["refund"], "arpu": block["arpu"]}


def _build_organization(sales: pd.DataFrame, customer_map: dict[str, dict[str, Any]], orders: pd.DataFrame, history_users: set[str], new_users: set[str], union_users: set[str]) -> list[dict[str, Any]]:
    seller_info = {_id(row.get("销售ID")): row for _, row in sales.iterrows() if _id(row.get("销售ID"))}
    assigned = [record for record in customer_map.values() if record.get("primary_seller_id")]
    def metric_for(sids: set[str], cohort: set[str]) -> dict[str, Any]:
        users = {uid for uid in cohort if customer_map.get(uid, {}).get("primary_seller_id") in sids}
        return {"history": _metric_block(users & history_users, orders), "new": _metric_block(users & new_users, orders), "combined": _metric_block(users & union_users, orders)}
    by_belong: dict[str, list[dict[str, Any]]] = {}
    for record in assigned:
        belong = record["belong"] or "未映射"
        by_belong.setdefault(belong, []).append(record)
    result = []
    for belong, records in sorted(by_belong.items()):
        group_map: dict[str, list[dict[str, Any]]] = {}
        for record in records:
            group_map.setdefault(record["group"] or "未映射", []).append(record)
        belong_node = {"level": "belong", "key": "belong:" + belong, "name": belong, "metrics": metric_for({_id(r["primary_seller_id"]) for r in records}, union_users), "children": []}
        for group, group_records in sorted(group_map.items()):
            group_sids = {_id(r["primary_seller_id"]) for r in group_records}
            group_node = {"level": "group", "key": "group:" + belong + "/" + group, "name": group, "belong": belong, "metrics": metric_for(group_sids, union_users), "children": []}
            for sid in sorted(group_sids, key=lambda value: (_text(seller_info.get(value, {}).get("销售")), value)):
                person = seller_info.get(sid, {})
                person_record = [r for r in group_records if _id(r["primary_seller_id"]) == sid]
                person_users = {r["user_id"] for r in person_record}
                group_node["children"].append({"level": "seller", "key": "seller:" + sid, "seller_id": sid, "name": _text(person.get("销售")) or "未命名销售", "belong": belong, "group": group, "status": _text(person.get("匹配状态")), "metrics": {"history": _metric_block(person_users & history_users, orders[orders["seller_id"] == sid]), "new": _metric_block(person_users & new_users, orders[orders["seller_id"] == sid]), "combined": _metric_block(person_users & union_users, orders[orders["seller_id"] == sid])}})
            belong_node["children"].append(group_node)
        result.append(belong_node)
    return result


def build_payload(audit: dict[str, Any], sales: pd.DataFrame, customers: pd.DataFrame, orders: pd.DataFrame) -> dict[str, Any]:
    sales, customers, orders = _prepare_frames(sales, customers, orders)
    orders = _unique_orders(orders)
    raw_users = {_id(x) for x in customers["user_id"] if _id(x)}
    window = audit.get("window") or ["", ""]
    start_dt = pd.Timestamp(window[0]) if window and window[0] else pd.Timestamp.min
    order_dates = pd.to_datetime(orders["pay_at"], errors="coerce")
    entry_times = customers.groupby("user_id")["入口首单时间"].first().map(lambda x: pd.to_datetime(x, errors="coerce"))
    historical_5000_users = set()
    for uid in raw_users:
        entry_dt = entry_times.get(uid)
        if pd.isna(entry_dt):
            continue
        mask = (orders["user_id"].map(_id) == uid) & (order_dates > entry_dt) & (order_dates < start_dt) & (orders["pay_amount"] >= 5000)
        if bool(mask.any()):
            historical_5000_users.add(uid)
    waiting_users = raw_users - historical_5000_users
    excluded_from_query = int(audit.get("excluded_historical_1000_users") or 0)
    customers = customers[customers["user_id"].map(_id).isin(waiting_users)].copy()
    history_users = {_id(x) for x in customers.loc[customers["队列"] == QUEUE_HISTORY, "user_id"] if _id(x)}
    new_users = {_id(x) for x in customers.loc[customers["队列"] == QUEUE_NEW, "user_id"] if _id(x)}
    union_users = history_users | new_users
    overlap_users = history_users & new_users
    customer_map: dict[str, dict[str, Any]] = {}
    primary_conflicts = []
    for uid, group in customers.groupby("user_id", dropna=False):
        uid = _id(uid)
        if not uid:
            continue
        if "客户归属类型" in group.columns:
            primary = group[group["客户归属类型"].astype(str).str.contains("归属")]
        else:
            primary = group
        primary_ids = {_id(x) for x in primary["销售ID"] if _id(x)}
        if len(primary_ids) > 1:
            primary_conflicts.append({"user_id": uid, "seller_ids": sorted(primary_ids)})
        sid = next(iter(primary_ids)) if len(primary_ids) == 1 else ""
        row = primary.iloc[0] if len(primary) else group.iloc[0]
        customer_map[uid] = {"user_id": uid, "primary_seller_id": sid, "belong": _text(row.get("归属")), "group": _text(row.get("组别")), "collaborator_ids": sorted({_id(x) for x in group.loc[~group.index.isin(primary.index), "销售ID"] if _id(x)})}
    history_orders = orders[orders["user_id"].map(_id).isin(history_users)]
    new_orders = orders[orders["user_id"].map(_id).isin(new_users)]
    combined_orders = orders[orders["user_id"].map(_id).isin(union_users)]
    history_block = _metric_block(history_users, history_orders)
    new_block = _metric_block(new_users, new_orders)
    combined_block = _metric_block(union_users, combined_orders)
    overlap_order_ids = set(history_orders["order_id"]) & set(new_orders["order_id"])
    entry_amounts = {"3980": 0, "4980": 0}
    for uid, group in customers.groupby("user_id", dropna=False):
        if _id(uid) not in new_users:
            continue
        amounts = set().union(*(_entry_amounts(x) for x in group.loc[group["队列"] == QUEUE_NEW, "入口金额"] if "入口金额" in group))
        for amount in amounts:
            entry_amounts[str(amount)] += 1
    history_entry = customers[customers["队列"] == QUEUE_HISTORY][["user_id", "入口首单时间"]].copy()
    history_entry["user_id"] = history_entry["user_id"].map(_id)
    history_entry["entry_dt"] = pd.to_datetime(history_entry["入口首单时间"], errors="coerce")
    history_entry = history_entry.dropna(subset=["entry_dt"]).groupby("user_id", as_index=False)["entry_dt"].min()
    first_orders = history_orders.copy()
    first_orders["pay_dt"] = pd.to_datetime(first_orders["pay_at"], errors="coerce")
    first_orders = first_orders.dropna(subset=["pay_dt"]).sort_values(["user_id", "pay_dt", "order_id"]).drop_duplicates("user_id")
    cycles = first_orders.merge(history_entry, on="user_id", how="inner")
    cycle_days = ((cycles["pay_dt"] - cycles["entry_dt"]).dt.total_seconds() / 86400).tolist()
    history_avg_cycle_days = round(sum(cycle_days) / len(cycle_days), 1) if cycle_days else 0.0
    seller_combined = []
    seller_info = {_id(row.get("销售ID")): row for _, row in sales.iterrows() if _id(row.get("销售ID"))}
    for sid, person in seller_info.items():
        person_users = {uid for uid, record in customer_map.items() if record.get("primary_seller_id") == sid}
        person_orders = orders[orders["seller_id"] == sid]
        seller_combined.append({"name": _text(person.get("销售")) or "未命名销售", "seller_id": sid, "belong": _text(person.get("归属")), "group": _text(person.get("组别")), "status": _text(person.get("匹配状态")), "history": _metric_block(person_users & history_users, person_orders), "new": _metric_block(person_users & new_users, person_orders), "combined": _metric_block(person_users & union_users, person_orders)})
    seller_combined.sort(key=lambda row: (-row["combined"]["converted_users"], -row["combined"]["cohort_users"], row["name"]))
    product_rows = []
    if not combined_orders.empty:
        for product, subset in combined_orders.assign(product=combined_orders.apply(lambda row: _text(row.get("product_name")) or (_id(row.get("product_id")) or "未标注"), axis=1)).groupby("product"):
            unique = _unique_orders(subset)
            product_rows.append({"product": product, "orders": int(len(unique)), "users": int(unique["user_id"].nunique()), "amount": _money(unique["pay_amount"].sum()), "refund": _money(unique["refund_total"].sum()), "share": round(unique["pay_amount"].sum() / combined_block["amount"] * 100, 1) if combined_block["amount"] else 0.0})
    product_rows.sort(key=lambda row: (-row["amount"], row["product"]))
    organization = _build_organization(sales, customer_map, orders, history_users, new_users, union_users)
    window = audit.get("window") or ["", ""]
    legacy_history = _metric_legacy(history_block)
    legacy_new = _metric_legacy(new_block)
    legacy_combined = _metric_legacy(combined_block)
    return {"schema_version": "2.0", "title": "二阶升单漏斗", "cohort_month": str(window[0])[:7] if window else "2026-10", "window": {"start": window[0] if len(window) > 0 else "", "end": window[1] if len(window) > 1 else "", "snapshot": audit.get("snapshot", "")}, "cohort_summary": {"raw_entry_users": len(raw_users), "excluded_historical_1000_users": excluded_from_query or len(historical_5000_users), "waiting_users": len(waiting_users), "history": len(history_users), "new": len(new_users), "overlap": len(overlap_users), "union": len(union_users)}, "stages": [{"label": "3980/4980待升单客户（排除历史5000+）", "count": combined_block["cohort_users"], "amount": None}, {"label": "本月升单客户", "count": combined_block["converted_users"], "amount": combined_block["amount"], "rate_label": "升单率"}], "overall": {"entry": {"count": new_block["cohort_users"], "new_count": new_block["cohort_users"], "entry_3980_users": entry_amounts["3980"], "entry_4980_users": entry_amounts["4980"]}, "conversion": legacy_new, "history": {**history_block, "count": history_block["cohort_users"], "converted": history_block["converted_users"], "avg_cycle_days": history_avg_cycle_days}, "new": new_block, "combined": combined_block, "overlap": {"users": len(overlap_users), "orders": len(overlap_order_ids)}}, "queue": [{"name": "历史4980_10月前", **legacy_history}, {"name": "新增4980_窗口新分配", **legacy_new}, {"name": "历史+本月合计", **legacy_combined}], "organization": organization, "seller_combined": seller_combined, "store": [], "cm": [], "seller": [], "product": product_rows, "audit": {"staff_rows": audit.get("staff_rows"), "matched_staff": audit.get("unique_matched_staff"), "unmatched_or_ambiguous": audit.get("unmatched_or_ambiguous", []), "order_audit": audit.get("order_audit", {}), "customer_detail_rows": audit.get("customer_detail_rows"), "order_detail_rows": audit.get("order_detail_rows"), "history_unique_users": len(history_users), "new_unique_users": len(new_users), "overlap_users": len(overlap_users), "union_unique_users": len(union_users), "history_unique_orders": len(set(history_orders["order_id"])), "new_unique_orders": len(set(new_orders["order_id"])), "cross_queue_duplicate_orders": len(overlap_order_ids), "union_unique_orders": len(set(combined_orders["order_id"])), "primary_owner_conflicts": primary_conflicts}, "limitations": ["入口池=3980/4980客户中排除入口后历史1000+订单的待升单客户", "本月升单按pay_amount≥1000，统计待升单客户的首次升单", "纯协作关系不增加主客户数，订单金额按实际seller_id归属", "无电话、加微、入群、D1-D3等中间行为数据", "无渠道字段和入口产品明细", "入口金额合并3980与4980", "组织层级为归属→组别→销售ID，不将观鼎组当作门店或独立CM组织树", "成交产品表仅代表成交侧产品分析"]}


def _build_payload_legacy(audit: dict[str, Any], sales: pd.DataFrame, customers: pd.DataFrame, orders: pd.DataFrame) -> dict[str, Any]:
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
  if(typeof _funnelCollapsed==='undefined') window._funnelCollapsed={};
  if(typeof _funnelCollapsed.funnel_4980==='undefined') _funnelCollapsed.funnel_4980=true;
  if(!f||!f.overall) return '';
  var o=f.overall,e=o.entry||{},n=o.new||{},hi=o.history||{},z=o.combined||{},key='funnel_4980',collapsed=_funnelCollapsed[key] !== false,h='';
  function esc2(x){return String(x==null?'':x).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
  function pct(x){return Number(x||0).toFixed(1)+'%';}
  function moneyWan(x){return '¥'+(Number(x||0)/10000).toFixed(1)+'万';}
  function moneyYuan(x){return '¥'+fmtInt(Math.round(Number(x||0)));}
  function metric(label,value,color){return _fMetric(label,value,color);}
  h+='<div class="funnel-card" id="funnel-card-4980">';
  h+='<div class="funnel-card-head"><div><div class="funnel-card-title">📈 二阶升单漏斗</div><div class="funnel-card-summary">待升单客户 '+fmtInt(z.cohort_users)+' 人 · 本月升单 '+fmtInt(z.converted_users)+' 人 · 升单率 '+pct(z.rate)+'</div></div><div class="funnel-card-actions"><button type="button" class="funnel-export-btn" onclick="exportFunnel4980()">导出</button>'+_fToggle('funnel_4980','二阶升单漏斗')+'</div></div>';
  h+='<div class="funnel-card-body" style="display:'+(collapsed?'none':'block')+'">';
  h+='<div style="font-size:15px;font-weight:700;color:#1e293b;margin-bottom:14px">📈 二阶升单漏斗<span style="font-weight:400;color:#94a3b8;font-size:11px;margin-left:8px">合并历史与本月新增，按user_id/order_id全局去重</span></div>';
  h+='<div style="font-size:12px;line-height:1.7;color:#64748b;margin-bottom:14px">统计窗口：'+esc2(f.window.start)+' 至 '+esc2(f.window.end)+'（不含） · 快照：'+esc2(f.window.snapshot)+'。历史与本月重叠客户 '+fmtInt((f.cohort_summary||{}).overlap||0)+' 人，跨队列重复订单 '+fmtInt((o.overlap||{}).orders||0)+' 单。</div>';
  h+='<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:12px;margin-bottom:16px">';
  h+=metric('🎯 3980/4980待升单客户',fmtInt(z.cohort_users)+' 人','#6366f1');
  h+=metric('🚀 合计本月升单',fmtInt(z.converted_users)+' 人','#10b981');
  h+=metric('📊 合计升单率',pct(z.rate),_fRateColor(z.rate));
  h+=metric('💰 合计升单金额',moneyWan(z.amount),'#ef4444');
  h+=metric('🧮 合计升单ARPU',moneyYuan(z.arpu),'#8b5cf6');
  h+=metric('🔄 历史队列本月升单',fmtInt(hi.converted_users||hi.converted||0)+'人 · '+moneyWan(hi.amount),'#06b6d4');
  h+='</div>';
  h+='<div style="font-size:12px;color:#475569;margin:0 0 10px">历史入口 '+fmtInt(hi.cohort_users||hi.count||0)+' 人 / '+fmtInt(hi.converted_users||hi.converted||0)+' 人升单 / '+moneyWan(hi.amount)+'；本月新增 '+fmtInt(n.cohort_users||0)+' 人 / '+fmtInt(n.converted_users||0)+' 人升单 / '+moneyWan(n.amount)+'；历史平均升单周期 '+Number(hi.avg_cycle_days||0).toFixed(1)+' 天。</div>';
  h+=_fBars(f.stages||[]);
  h+='<div style="display:flex;gap:8px;align-items:center;margin:20px 0 8px;flex-wrap:wrap"><span style="font-size:13px;font-weight:700;color:#334155">组织与人员维度</span><button class="f4980-view active" data-v="combined">合计</button><button class="f4980-view" data-v="history">历史</button><button class="f4980-view" data-v="new">本月新增</button></div>';
  function rowsFor(nodes,view){return (nodes||[]).map(function(x){return {name:x.name,metrics:(x.metrics||{})[view]||{}};});}
  function tableView(title,nodes,view){var rows=rowsFor(nodes,view);if(!rows.length)return '';var x='<div class="f4980-view-table" data-view="'+view+'"><div style="font-size:13px;font-weight:700;color:#334155;margin:14px 0 8px">'+title+'</div><table><thead><tr><th>组织/人员</th><th class="num">4980客户数</th><th class="num">升单</th><th class="num">升单率</th><th class="num">升单金额</th><th class="num">ARPU</th></tr></thead><tbody>';rows.forEach(function(r){var m=r.metrics||{};x+='<tr><td>'+esc2(r.name||'')+'</td><td class="num">'+fmtInt(m.cohort_users)+'</td><td class="num">'+fmtInt(m.converted_users)+'</td><td class="num" style="color:'+_fRateColor(m.rate)+';font-weight:700">'+pct(m.rate)+'</td><td class="num">'+moneyWan(m.amount)+'</td><td class="num">'+moneyYuan(m.arpu)+'</td></tr>';});return x+'</tbody></table></div>';}
  var groups=[],people=[];(f.organization||[]).forEach(function(b){(b.children||[]).forEach(function(g){groups.push(g);(g.children||[]).forEach(function(p){people.push(p);});});});
  ['combined','history','new'].forEach(function(v){h+=tableView('🏢 归属组织',f.organization,v);h+=tableView('🏬 组别/院区',groups,v);h+=tableView('🧑 销售人员',people,v);});
  if(f.product&&f.product.length)h+=_f4980Table('📦 合并成交产品',['成交产品','订单数','成交金额','覆盖用户','金额占比'],f.product.map(function(r){return [r.product,fmtInt(r.orders),moneyWan(r.amount),fmtInt(r.users),pct(r.share)];}));
  h+='<div style="font-size:12px;line-height:1.8;color:#64748b;margin-top:16px"><b>口径：</b>合计按user_id和canonical order全局去重；协作关系不增加主客户数；订单金额按实际seller_id归属；观鼎组为二阶组织组，线下组别作为院区，不等同于门店或CM。</div></div>';
  setTimeout(function(){var root=document.getElementById('funnel_cards');if(!root)return;root.querySelectorAll('.f4980-view').forEach(function(btn){btn.onclick=function(){var card=btn.closest('div[style*="border:1px solid #eef2f7"]');card.querySelectorAll('.f4980-view').forEach(function(b){b.classList.remove('active');});btn.classList.add('active');card.querySelectorAll('.f4980-view-table').forEach(function(t){t.style.display=t.getAttribute('data-view')===btn.getAttribute('data-v')?'block':'none';});};});root.querySelectorAll('.f4980-view-table').forEach(function(t){t.style.display=t.getAttribute('data-view')==='combined'?'block':'none';});},0);
  return h;
}
function exportFunnel4980(){
  if(typeof XLSX==='undefined'){alert('Excel 库未加载');return;}
  var f=window._M&&window._M.funnel&&window._M.funnel.funnel_4980;if(!f){alert('当前漏斗暂无数据');return;}
  var wb=XLSX.utils.book_new(),o=f.overall||{},z=o.combined||{};
  function add(name,rows){XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows&&rows.length?rows:[{'说明':'暂无数据'}]),name);}
  add('概览',[{'指标':'3980/4980待升单客户','人数':z.cohort_users,'本月升单客户':z.converted_users,'升单率':(z.rate||0)+'%','升单金额（元）':z.amount,'ARPU（元）':z.arpu},{'指标':'历史排除5000+升单客户','人数':(f.cohort_summary||{}).excluded_historical_1000_users||0}]);
  add('漏斗阶段',(f.stages||[]).map(function(r){return {'阶段':r.label,'人数':r.count,'金额（元）':r.amount,'转化率':r.rate_label||''};}));
  var people=[];(f.organization||[]).forEach(function(b){(b.children||[]).forEach(function(g){(g.children||[]).forEach(function(p){people.push({归属:b.name,组别:g.name,销售:p.name,销售ID:p.seller_id,历史客户:p.metrics.history.cohort_users,历史升单:p.metrics.history.converted_users,本月新增客户:p.metrics.new.cohort_users,本月新增升单:p.metrics.new.converted_users,合计客户:p.metrics.combined.cohort_users,合计升单:p.metrics.combined.converted_users,合计金额:p.metrics.combined.amount});});});});
  add('销售人员',people);add('成交产品',f.product);var month=(f.cohort_month||'2026-10').replace(/-/g,'');XLSX.writeFile(wb,'升单漏斗_二阶_'+month+'.xlsx');
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
    print(f"Verified: union entry {payload['overall']['combined']['cohort_users']:,}; conversion {payload['overall']['combined']['converted_users']:,}/{payload['overall']['combined']['orders']:,}; amount {payload['overall']['combined']['amount']:,.2f}; local preview only")
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
