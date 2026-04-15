#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Coze知识库自动同步脚本
功能：监听本地MD文档变化，自动同步到Coze知识库
"""

import os
import json
import time
from pathlib import Path

# 配置
TOKEN = "pat_zfHHM2sTlCivOrjoqZFIw36OWSmHsf7s1t4rb6SKa1vGrVsZbdHxchd7fTVTR9at"
KNOWLEDGE_BASE_NAME = "观心CRM知识库网站"
DOCS_DIR = "./知识库"
BOT_ID = "7621136728413274122"

def sync_documents():
    """同步所有MD文档到Coze"""
    print(f"开始同步文档到Coze知识库: {KNOWLEDGE_BASE_NAME}")

    # 遍历所有MD文件
    md_files = list(Path(DOCS_DIR).rglob("*.md"))
    print(f"找到 {len(md_files)} 个MD文件")

    for md_file in md_files:
        print(f"  - {md_file.relative_to(DOCS_DIR)}")

    print("\n提示：由于Coze API文档限制，当前脚本仅列出文件")
    print("请手动在Coze平台上传这些文档，或等待API文档更新后完善脚本")

if __name__ == "__main__":
    sync_documents()
