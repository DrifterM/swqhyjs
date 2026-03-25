# archive_catalog 技能

## 用途
自动生成符合规范的“期货交易咨询业务档案卷内文件目录”Word文档。

## 输入参数
- `project_path`: 项目文件夹路径（必需）
- `client_name`: 客户名称（可选，自动从文件名推断）
- `client_code`: 客户代码（可选，优先从台账匹配）
- `year`: 年份（默认为 "2025"）
- `seq_num`: 档号序号（4位，如 "0001"）

## 输出
在指定项目文件夹中生成 `卷内文件目录.docx`

## 使用方式
```python
# 方式一：基本调用（手动指定参数）
use_skill("archive_catalog", project_path="D:\\...\\新项目名称")

# 方式二：自动匹配客户代码（推荐）
use_skill("archive_catalog", 
             project_path="D:\\...\\新项目名称",
             client_name="客户姓名",
             year="2025",
             seq_num="0001",
             auto_match_code=True,
             ledger_path="D:\\申银万国期货\\3.档案管理\\2、业务档案\\3. 台账相关-每年随档案报送\\期货交易咨询跟踪台账-魏子骥-*.xlsx"
)
```

## 功能特性
- ✅ 自动扫描以数字开头的文件（如 1.xxx, 2.xxx）
- ✅ 支持任意数量的文件（动态生成行数）
- ✅ 严格遵循8列表格与字体格式规范
- ✅ 前四行内容居中对齐
- ✅ 包含“小计”行
- ✅ 可通过命令行或 agent 调用
- ✅ 支持自动从台账匹配客户代码（`auto_match_code=True`）
- 💾 生成结果持久化，不依赖临时环境