#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
知识库文档打包工具
自动将MD文档打包，方便上传到Coze
"""

import os
import shutil
from pathlib import Path
from datetime import datetime

DOCS_DIR = "./知识库"
OUTPUT_DIR = "./coze_upload"

def package_docs():
    """打包所有MD文档"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_name = f"观心CRM知识库_{timestamp}"

    # 创建临时目录
    temp_dir = Path(OUTPUT_DIR) / output_name
    temp_dir.mkdir(parents=True, exist_ok=True)

    # 复制所有MD文件
    md_files = list(Path(DOCS_DIR).rglob("*.md"))
    print(f"正在打包 {len(md_files)} 个文档...")

    for md_file in md_files:
        rel_path = md_file.relative_to(DOCS_DIR)
        dest = temp_dir / rel_path
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(md_file, dest)
        print(f"  [OK] {rel_path}")

    # 创建压缩包
    archive_path = shutil.make_archive(
        str(Path(OUTPUT_DIR) / output_name),
        'zip',
        temp_dir
    )

    # 清理临时目录
    shutil.rmtree(temp_dir)

    print(f"\n[SUCCESS] 打包完成: {archive_path}")
    print(f"文件大小: {os.path.getsize(archive_path) / 1024:.2f} KB")
    print(f"\n请将此文件上传到Coze知识库")

if __name__ == "__main__":
    package_docs()
