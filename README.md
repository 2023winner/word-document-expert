# Word Document Expert - Word 文档处理专家

专业处理 Word 文档（.docx）的创建、编辑和格式优化的技能。特别擅长处理中文文档、LaTeX 公式、图表插入和格式统一。

## ✨ 功能特点

- 📝 创建和编辑 Word 文档
- 🔢 支持 LaTeX 数学公式
- 📊 生成和插入图表
- 🌏 完美的中文支持
- 🎨 格式统一和优化
- 📋 作业报告生成

## 📦 安装

```bash
git clone https://github.com/2023winner/word-document-expert.git
cd word-document-expert
pip install pypandoc_binary python-docx matplotlib numpy
```
## 🚀 使用方法

### 基本使用

```python
import pypandoc

# 将 Markdown 转换为 Word 文档
output = pypandoc.convert_file(
    'input.md',
    'docx',
    outputfile='output.docx',
    extra_args=['--variable', 'CJKmainfont=SimSun']
)
```

### 生成报告

```python
from word_document_expert import ReportGenerator

generator = ReportGenerator()
generator.create_report('data.csv', 'report.docx')
```
## 🛠️ 技术栈

- Python
- python-docx
- pypandoc
- matplotlib

## 📁 目录结构

```
word-document-expert/
├── src/              # 源代码目录
├── docs/             # 文档目录
├── tests/            # 测试文件
├── examples/         # 示例代码
├── README.md         # 项目说明
└── .gitignore        # Git 忽略文件
```

## ⚠️ 注意事项

处理中文文档时请确保系统已安装中文字体。
## 📄 许可证

本项目采用 MIT 许可证 - 详见 LICENSE 文件
