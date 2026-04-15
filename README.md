# 企业微信智能问答系统

## 项目目标
搭建一个企业微信智能问答系统，自动读取企业微信文档内容并生成高质量的QA知识库。

## 系统架构

### 第一阶段：文档内容获取 (fetch_docs.py)
- 使用 Corp ID + Secret 获取 Access Token
- 调用企业微信文档API读取文档内容
- 保存原始内容到 `./raw_docs` 目录

### 第二阶段：QA知识库生成 (generate_qa.py)
- 读取 `./raw_docs` 中的文档内容
- 调用 DeepSeek API 生成高质量QA对
- 保存到 `./qa_output` 目录（JSON格式）

## 配置信息

### 企业微信配置
- **Corp ID**: `wwe032718ce9485b7d`
- **Corp Secret**: `QEmipYLHRVvYHXOi2nVq6igk4oGBmVo7j_9lnA6fHXY`

### 文档列表
1. **系统相关知识库**
   - URL: https://doc.weixin.qq.com/wiki/w.AI4A4AcKABA.R2hWBEPzw3c
   - DocID: AI4A4AcKABA
   - 类型: wiki

2. **门店个案知识库**
   - URL: https://doc.weixin.qq.com/wiki/w.AI4A4AcKABA.Jx-CZq7P3Rs
   - DocID: AI4A4AcKABA
   - 类型: wiki

3. **单独QA文档**
   - URL: https://doc.weixin.qq.com/sheet/e3_AXAAoQYrAIQCN61pqVd8uSQWwhAqa
   - DocID: AXAAoQYrAIQCN61pqVd8uSQWwhAqa
   - 类型: sheet

### DeepSeek API配置
- **Base URL**: https://api.deepseek.com
- **Model**: deepseek-chat
- **API Key**: 从环境变量 `DEEPSEEK_API_KEY` 读取

## 当前状态

### ✅ 已完成
- [x] 创建项目目录结构 (`raw_docs/`, `qa_output/`)
- [x] 实现 `fetch_docs.py` 脚本
  - Access Token 获取功能正常
  - 支持多种API端点尝试
  - 文件保存逻辑完整

### ⚠️ 待解决问题
**API权限问题 (错误码: 48002)**
- 当前使用的 Corp Secret 对应的应用没有开通企业微信文档API权限
- 需要在企业微信管理后台配置应用权限

### 📋 下一步计划
1. **解决权限问题**
   - 在企业微信管理后台找到对应应用
   - 开通"微文档"相关API权限
   - 可能需要的权限：
     - `wedoc:doc_get_content` - 获取文档内容
     - `wedoc:doc_get` - 获取文档详情
     - `wedoc:smartsheet_get` - 获取表格内容

2. **测试文档获取**
   - 权限配置完成后运行 `python fetch_docs.py`
   - 验证文档内容是否正确保存到 `./raw_docs`

3. **开发第二阶段**
   - 实现 `generate_qa.py`
   - 调用 DeepSeek API 生成QA对

## 临时解决方案
在等待API权限开通期间，可以手动导出文档：
1. 打开文档链接
2. 导出为 txt 格式
3. 保存到 `./raw_docs/` 目录：
   - `系统相关知识库.txt`
   - `门店个案知识库.txt`
   - `单独QA文档.txt`

## 使用说明

### 运行文档获取脚本
```bash
python fetch_docs.py
```

### 运行QA生成脚本（待开发）
```bash
# 需要先设置环境变量
export DEEPSEEK_API_KEY="your_api_key_here"

python generate_qa.py
```

## 目录结构
```
.
├── fetch_docs.py          # 文档获取脚本
├── generate_qa.py         # QA生成脚本（待开发）
├── raw_docs/              # 原始文档存储目录
├── qa_output/             # 生成的QA知识库目录
└── README.md              # 项目说明文档
```

## QA生成要求
生成的QA需要满足以下标准：
- ✅ 站在员工角度思考真实问法，不要机械转换标题
- ✅ 同一知识点生成多种问法（如"怎么请假"、"请假流程"、"我要请假怎么操作"）
- ✅ 答案口语化、实用，像熟悉业务的同事在回答
- ✅ 操作类问题包含具体步骤
- ✅ 输出格式为JSON，便于后续导入知识库系统

## 联系方式
如有问题，请联系项目负责人。
