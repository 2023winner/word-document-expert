# Word Document Expert - Word 文档处理专家

📚 专业处理 Word 文档（.docx）的创建、编辑和格式优化的技能。

## ✨ 功能特点

- 📝 创建和编辑 Word 文档
- 🔢 支持 LaTeX 数学公式
- 📊 生成和插入图表
- 🌏 完美的中文支持
- 🎨 格式统一和优化
- 📋 作业报告生成

## 🚀 快速开始

### 安装依赖

```bash
pip install pypandoc_binary
pip install python-docx
pip install matplotlib
pip install numpy
```

### 使用示例

```python
import pypandoc

# 将 Markdown 转换为 Word 文档
output = pypandoc.convert_file(
    'input.md',
    'docx',
    outputfile='output.docx',
    extra_args=[
        '--variable', 'CJKmainfont=SimSun',
        '--variable', 'fontsize=12pt'
    ]
)
```

## 📖 文档

详细使用说明请查看 [SKILL.md](SKILL.md)

## 🛠️ 核心功能

### 1. 文档创建
- 标准作业文档模板
- LaTeX 公式支持
- 图表自动生成

### 2. 格式处理
- 中文字体配置
- 统一格式规范
- 图片标题自动编号

### 3. 质量验证
- 12 项质量检查清单
- 自动化验证脚本
- 格式一致性保证

## 📋 使用场景

- 📚 作业报告生成
- 📄 学术论文排版
- 📊 数据报告制作
- 🎓 毕业论文格式
- 📝 办公文档处理

## 🔧 工具要求

- Python 3.7+
- Pandoc（pypandoc_binary 已包含）
- Matplotlib（图表生成）
- python-docx（文档处理）

## 📦 项目结构

```
word-document-expert/
├── SKILL.md          # 技能详细说明
├── README.md         # 本文件
├── requirements.txt  # Python 依赖
└── .gitignore       # Git 忽略文件
```

## 🎯 最佳实践

1. **使用 Markdown 源文件** - 便于版本控制
2. **配置中文字体** - 避免乱码问题
3. **使用 LaTeX 公式** - 保证数学公式质量
4. **设置图片 DPI** - 300 及以上保证打印质量
5. **保留源文件** - 便于后续修改

## 📝 许可证

MIT License

## 👥 贡献

欢迎提交 Issue 和 Pull Request！

## 🙏 致谢

基于神经网络作业实践经验总结
