#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
企业微信文档内容获取脚本
读取企业微信文档并保存到本地
"""

import requests
import json
import os
import time

# 企业微信配置
CORP_ID = "wwe032718ce9485b7d"
CORP_SECRET = "QEmipYLHRVvYHXOi2nVq6igk4oGBmVo7j_9lnA6fHXY"

# 文档列表
DOCS = [
    {
        "name": "系统相关知识库",
        "url": "https://doc.weixin.qq.com/wiki/w.AI4A4AcKABA.R2hWBEPzw3c",
        "docid": "AI4A4AcKABA",
        "type": "wiki"
    },
    {
        "name": "门店个案知识库",
        "url": "https://doc.weixin.qq.com/wiki/w.AI4A4AcKABA.Jx-CZq7P3Rs",
        "docid": "AI4A4AcKABA",
        "type": "wiki"
    },
    {
        "name": "单独QA文档",
        "url": "https://doc.weixin.qq.com/sheet/e3_AXAAoQYrAIQCN61pqVd8uSQWwhAqa",
        "docid": "AXAAoQYrAIQCN61pqVd8uSQWwhAqa",
        "type": "sheet"
    }
]

OUTPUT_DIR = "./raw_docs"


def get_access_token():
    """获取企业微信Access Token"""
    url = f"https://qyapi.weixin.qq.com/cgi-bin/gettoken?corpid={CORP_ID}&corpsecret={CORP_SECRET}"

    print(f"正在获取Access Token...")
    response = requests.get(url)
    result = response.json()

    if result.get("errcode") == 0:
        token = result.get("access_token")
        print(f"[OK] Access Token获取成功: {token[:20]}...")
        return token
    else:
        print(f"[ERROR] Access Token获取失败: {result}")
        return None


def fetch_wiki_content(access_token, docid):
    """获取wiki文档内容 - 尝试多个API端点"""

    # 尝试方法1: wedoc API
    url1 = f"https://qyapi.weixin.qq.com/cgi-bin/wedoc/doc_get_content?access_token={access_token}"
    payload1 = {"docid": docid}

    print(f"  正在获取wiki内容 (docid: {docid})...")
    print(f"  尝试API: wedoc/doc_get_content")
    response = requests.post(url1, json=payload1)
    result = response.json()

    if result.get("errcode") == 0:
        return result

    print(f"  失败 (errcode: {result.get('errcode')}), 尝试其他API...")

    # 尝试方法2: 文档详情API
    url2 = f"https://qyapi.weixin.qq.com/cgi-bin/wedoc/doc_get?access_token={access_token}"
    payload2 = {"docid": docid}

    print(f"  尝试API: wedoc/doc_get")
    response = requests.post(url2, json=payload2)
    result = response.json()

    if result.get("errcode") == 0:
        return result
    else:
        print(f"  [ERROR] 所有API均失败: {result}")
        return None


def fetch_sheet_content(access_token, docid):
    """获取表格文档内容 - 尝试多个API端点"""

    # 尝试方法1: smartsheet API
    url1 = f"https://qyapi.weixin.qq.com/cgi-bin/wedoc/smartsheet/get_sheet?access_token={access_token}"
    payload1 = {"docid": docid}

    print(f"  正在获取表格内容 (docid: {docid})...")
    print(f"  尝试API: wedoc/smartsheet/get_sheet")
    response = requests.post(url1, json=payload1)
    result = response.json()

    if result.get("errcode") == 0:
        return result

    print(f"  失败 (errcode: {result.get('errcode')}), 尝试其他API...")

    # 尝试方法2: 文档详情API
    url2 = f"https://qyapi.weixin.qq.com/cgi-bin/wedoc/doc_get?access_token={access_token}"
    payload2 = {"docid": docid}

    print(f"  尝试API: wedoc/doc_get")
    response = requests.post(url2, json=payload2)
    result = response.json()

    if result.get("errcode") == 0:
        return result
    else:
        print(f"  [ERROR] 所有API均失败: {result}")
        return None


def save_content(doc_name, content):
    """保存文档内容到本地"""
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    filename = f"{doc_name.replace('/', '_')}.txt"
    filepath = os.path.join(OUTPUT_DIR, filename)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(json.dumps(content, ensure_ascii=False, indent=2))

    print(f"  [OK] 已保存到: {filepath}")


def main():
    print("=" * 60)
    print("企业微信文档内容获取工具")
    print("=" * 60)
    print()

    # 1. 获取Access Token
    access_token = get_access_token()
    if not access_token:
        print("\n程序终止：无法获取Access Token")
        return

    print()
    print("-" * 60)
    print()

    # 2. 逐个获取文档内容
    for i, doc in enumerate(DOCS, 1):
        print(f"[{i}/{len(DOCS)}] {doc['name']}")
        print(f"  URL: {doc['url']}")

        if doc['type'] == 'wiki':
            content = fetch_wiki_content(access_token, doc['docid'])
        elif doc['type'] == 'sheet':
            content = fetch_sheet_content(access_token, doc['docid'])
        else:
            print(f"  [ERROR] 未知文档类型: {doc['type']}")
            continue

        if content:
            save_content(doc['name'], content)
            print(f"  [OK] 成功获取内容")
        else:
            print(f"  [ERROR] 获取失败")

        print()
        time.sleep(1)  # 避免请求过快

    print("=" * 60)
    print("完成！所有文档已保存到 ./raw_docs 目录")
    print("=" * 60)


if __name__ == "__main__":
    main()
