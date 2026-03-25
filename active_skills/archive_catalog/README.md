# archive_catalog

> **技能状态：已就绪**

本技能用于自动生成标准化的“期货交易咨询业务档案卷内文件目录”Word文档。

## 快速开始

```bash
# 1. 进入技能目录
cd C:\Users\H3C\.copaw\active_skills\archive_catalog

# 2. 直接运行（需激活 venv）
D:\CoPaw\venv\Scripts\python.exe generate_catalog.py "D:\\路径\\到\\项目文件夹" "客户名称" "客户代码" "2025" "0001"
```

## 参数说明
| 参数 | 说明 |
|------|------|
| `project_path` | 项目文件夹路径（必需） |
| `client_name` | 客户名称（可选） |
| `client_code` | 客户代码（可选） |
| `year` | 年份（默认 2025） |
| `seq_num` | 档号序号（4位） |

## 验证
- ✅ 已验证可在目标环境下生成 `.docx` 文件
- ✅ 中文路径与编码问题已规避
- ✅ 输出格式经人工确认无误

---

*此技能由夜澜为柏林构建，2026年3月13日。*