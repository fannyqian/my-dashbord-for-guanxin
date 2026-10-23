import unittest

import pandas as pd

from preview_second_stage_4980 import build_payload


class SecondStagePreviewTest(unittest.TestCase):
    def setUp(self):
        self.audit = {
            "snapshot": "20261022",
            "window": ["2026-10-01 00:00:00", "2026-10-23 00:00:00"],
            "staff_rows": 2,
            "unique_matched_staff": 2,
        }
        self.sales = pd.DataFrame([
            {"销售": "A", "销售ID": "1", "归属": "二阶", "组别": "观鼎组", "匹配状态": "唯一匹配"},
            {"销售": "B", "销售ID": "2", "归属": "线下", "组别": "南京院区", "匹配状态": "唯一匹配"},
        ])
        self.customers = pd.DataFrame([
            {"销售": "A", "销售ID": "1", "归属": "二阶", "组别": "观鼎组", "user_id": "u1", "队列": "新增4980_窗口新分配", "入口金额": "3980.0;4980.0"},
            {"销售": "A", "销售ID": "1", "归属": "二阶", "组别": "观鼎组", "user_id": "u1", "队列": "历史4980_10月前", "入口金额": "4980.0"},
            {"销售": "B", "销售ID": "2", "归属": "线下", "组别": "南京院区", "user_id": "u2", "队列": "历史4980_10月前", "入口金额": "3980.0"},
        ])
        self.orders = pd.DataFrame([
            {"销售": "A", "销售ID": "1", "user_id": "u1", "队列": "新增4980_窗口新分配", "order_id": "o1", "seller_id": "1", "pay_amount": "4980", "refund_total": "100", "product_id": "p1", "product_name": "产品A"},
            {"销售": "A", "销售ID": "1", "user_id": "u1", "队列": "历史4980_10月前", "order_id": "o1", "seller_id": "1", "pay_amount": "4980", "refund_total": "100", "product_id": "p1", "product_name": "产品A"},
            {"销售": "B", "销售ID": "2", "user_id": "u2", "队列": "历史4980_10月前", "order_id": "o2", "seller_id": "2", "pay_amount": "1980", "refund_total": "0", "product_id": "p2", "product_name": "产品B"},
            {"销售": "B", "销售ID": "2", "user_id": "u3", "队列": "历史4980_10月前", "order_id": "o3", "seller_id": "999", "pay_amount": "6600", "refund_total": "0", "product_id": "p3", "product_name": "错误归属"},
        ])

    def test_global_and_queue_dedup(self):
        payload = build_payload(self.audit, self.sales, self.customers, self.orders)
        entry = payload["overall"]["entry"]
        conversion = payload["overall"]["conversion"]
        self.assertEqual(entry["count"], 1)
        self.assertEqual(entry["new_count"], 1)
        self.assertEqual(entry["entry_3980_users"], 1)
        self.assertEqual(entry["entry_4980_users"], 1)
        self.assertEqual(conversion["count"], 1)
        self.assertEqual(conversion["orders"], 1)
        self.assertEqual(conversion["amount"], 4980.0)
        self.assertEqual(conversion["refund"], 100.0)
        self.assertEqual(conversion["rate"], 100.0)
        self.assertEqual(payload["queue"][0]["customers"], 1)

    def test_seller_and_product_dimensions(self):
        payload = build_payload(self.audit, self.sales, self.customers, self.orders)
        sellers = {row["seller_id"]: row for row in payload["seller"]}
        self.assertEqual(sellers["1"]["up"], 1)
        self.assertEqual(sellers["2"]["up"], 0)
        self.assertEqual(sellers["2"]["amount"], 0.0)
        products = {row["product"]: row for row in payload["product"]}
        self.assertEqual(products["产品A"]["orders"], 1)
        self.assertEqual(products["产品A"]["amount"], 4980.0)
        self.assertNotIn("产品B", products)
        stores = {row["store"]: row for row in payload["store"]}
        self.assertNotIn("南京院区", stores)


if __name__ == "__main__":
    unittest.main()
