#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
QA知识库生成脚本
读取企业微信文档内容，使用DeepSeek API生成高质量QA对
"""

import os
import json
import time
from openai import OpenAI

# DeepSeek API配置
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY")
DEEPSEEK_BASE_URL = "https://api.deepseek.com"
DEEPSEEK_MODEL = "deepseek-chat"

# 目录配置
RAW_DOCS_DIR = "./raw_docs"
QA_OUTPUT_DIR = "./qa_output"

# QA生成提示词
SYSTEM_PROMPT = """你是一个专业的企业知识库QA生成专家。你的任务是将企业文档内容转换为高质量的问答对。

要求：
1. 站在员工角度思考：员工真正会怎么问？不要机械地把标题变成问题
2. 多样化问法：同一个知识点生成3-5种不同的问法
   - 例如："怎么请假"、"请假流程是什么"、"我要请假怎么操作"、"如何申请休假"
3. 答案要口语化、实用：像一个熟悉业务的同事在回答，不要生硬的官方文档语气
4. 操作类问题必须包含具体步骤：第一步、第二步、第三步...
5. 保留关键信息：时间、人名、系统名称、具体数字等
6. 避免模糊回答：不要说"请联系相关部门"，要说清楚联系谁

输出格式：
返回JSON数组，每个元素包含：
{
  "question": "员工可能问的问题",
  "answer": "清晰实用的答案",
  "keywords": ["关键词1", "关键词2"],
  "category": "分类（如：考勤、报销、系统操作等）"
}
"""


def init_deepseek_client():
    """初始化DeepSeek客户端"""
    if not DEEPSEEK_API_KEY:
        print("[ERROR] 未设置环境变量 DEEPSEEK_API_KEY")
        print("请运行: export DEEPSEEK_API_KEY='your_api_key'")
        return None

    client = OpenAI(
        api_key=DEEPSEEK_API_KEY,
        base_url=DEEPSEEK_BASE_URL
    )
    print("[OK] DeepSeek客户端初始化成功")
    return client


def read_raw_docs():
    """读取原始文档"""
    if not os.path.exists(RAW_DOCS_DIR):
        print(f"[ERROR] 目录不存在: {RAW_DOCS_DIR}")
        return []

    docs = []
    for filename in os.listdir(RAW_DOCS_DIR):
        if filename.endswith('.txt'):
            filepath = os.path.join(RAW_DOCS_DIR, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
                docs.append({
                    "filename": filename,
                    "content": content
                })
            print(f"[OK] 已读取: {filename} ({len(content)} 字符)")

    return docs


def generate_qa_for_doc(client, doc_name, doc_content):
    """为单个文档生成QA"""
    print(f"\n正在为 {doc_name} 生成QA...")

    user_prompt = f"""请为以下企业文档内容生成高质量的QA对：

文档名称：{doc_name}

文档内容：
{doc_content[:8000]}  # 限制长度避免超出token限制

请生成至少20个QA对，覆盖文档中的主要知识点。
"""

    try:
        response = client.chat.completions.create(
            model=DEEPSEEK_MODEL,
            messages=[
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": user_prompt}
            ],
            temperature=0.7,
            max_tokens=4000
        )

        result_text = response.choices[0].message.content

        # 尝试解析JSON
        # 如果返回的是markdown格式的json，需要提取
        if "```json" in result_text:
            json_start = result_text.find("```json") + 7
            json_end = result_text.find("```", json_start)
            result_text = result_text[json_start:json_end].strip()
        elif "```" in result_text:
            json_start = result_text.find("```") + 3
            json_end = result_text.find("```", json_start)
            result_text = result_text[json_start:json_end].strip()

        qa_list = json.loads(result_text)
        print(f"[OK] 成功生成 {len(qa_list)} 个QA对")
        return qa_list

    except Exception as e:
        print(f"[ERROR] 生成失败: {e}")
        return []


def save_qa_output(doc_name, qa_list):
    """保存QA到文件"""
    os.makedirs(QA_OUTPUT_DIR, exist_ok=True)

    # 生成文件名
    base_name = doc_name.replace('.txt', '').replace('.json', '')
    output_file = os.path.join(QA_OUTPUT_DIR, f"{base_name}_qa.json")

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(qa_list, f, ensure_ascii=False, indent=2)

    print(f"[OK] 已保存到: {output_file}")


def merge_all_qa():
    """合并所有QA到一个文件"""
    all_qa = []

    for filename in os.listdir(QA_OUTPUT_DIR):
        if filename.endswith('_qa.json'):
            filepath = os.path.join(QA_OUTPUT_DIR, filename)
            with open(filepath, 'r', encoding='utf-8') as f:
                qa_list = json.load(f)
                all_qa.extend(qa_list)

    # 保存合并后的文件
    merged_file = os.path.join(QA_OUTPUT_DIR, "all_qa_merged.json")
    with open(merged_file, 'w', encoding='utf-8') as f:
        json.dump(all_qa, f, ensure_ascii=False, indent=2)

    print(f"\n[OK] 已合并所有QA到: {merged_file}")
    print(f"[OK] 总计: {len(all_qa)} 个QA对")


def main():
    print("=" * 60)
    print("QA知识库生成工具")
    print("=" * 60)
    print()

    # 1. 初始化DeepSeek客户端
    client = init_deepseek_client()
    if not client:
        return

    print()
    print("-" * 60)
    print()

    # 2. 读取原始文档
    docs = read_raw_docs()
    if not docs:
        print("[ERROR] 没有找到任何文档")
        print(f"请确保 {RAW_DOCS_DIR} 目录中有 .txt 文件")
        return

    print(f"\n找到 {len(docs)} 个文档")
    print("-" * 60)

    # 3. 为每个文档生成QA
    for i, doc in enumerate(docs, 1):
        print(f"\n[{i}/{len(docs)}] 处理: {doc['filename']}")
        qa_list = generate_qa_for_doc(client, doc['filename'], doc['content'])

        if qa_list:
            save_qa_output(doc['filename'], qa_list)

        # 避免API请求过快
        if i < len(docs):
            print("等待3秒...")
            time.sleep(3)

    print()
    print("-" * 60)

    # 4. 合并所有QA
    merge_all_qa()

    print()
    print("=" * 60)
    print("完成！QA知识库已生成")
    print("=" * 60)


if __name__ == "__main__":
    main()
