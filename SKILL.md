---
name: word-document-expert
description: 专业处理 Word 文档（.docx）的创建、编辑和格式优化。当你需要处理 Word 文档、生成作业报告、创建正式文档或解决文档格式问题时使用此技能。特别擅长处理中文文档、LaTeX 公式、图表插入和格式统一。
---

# Word 文档处理专家

## 兼容性
- 必需工具：pandoc, pypandoc, python-docx, matplotlib
- 可选工具：pywin32（Windows 专用，用于完整格式支持）

## 核心工作流程

### 1. 需求分析阶段
在开始处理文档前，必须明确以下信息：

**必须确认的要点：**
1. 文档类型（作业、报告、论文等）
2. 目标格式要求（字体、字号、标题层级等）
3. 是否包含 LaTeX 公式
4. 是否需要插入图表
5. 中文字体要求（等线、宋体、黑体等）
6. 最终提交格式（.docx, .pdf 等）

**关键问题模板：**
```
在开始之前，请确认：
1. 文档的用途和提交要求是什么？
2. 需要匹配原文档的格式吗？（字体、字号、样式）
3. 文档中包含数学公式吗？（需要使用 LaTeX）
4. 需要插入图片吗？（图表、截图等）
5. 有特殊的中文字体要求吗？
```

### 2. 技术选型决策树

**根据需求选择工具：**

```
需求分析 → 工具选择
├─ 简单文本转换 → python-docx
├─ 包含 LaTeX 公式 → pandoc（首选）
├─ 需要完整格式保留 → pywin32（Windows only）
├─ 需要生成图表 → matplotlib + pandoc
└─ 复杂格式编辑 → pywin32 读取 → pandoc 转换
```

**工具选择原则：**
- **pandoc**: 处理 LaTeX 公式的首选，格式转换最可靠
- **python-docx**: 简单文档创建，不支持公式
- **pywin32**: 最完整的 Word API 支持，但仅 Windows 可用

### 3. 文档创建标准流程

#### 步骤 1: 准备 Markdown 源文件
```markdown
# 一级标题
## 二级标题
### 三级标题

正文内容使用标准 Markdown 语法。

**LaTeX 公式示例：**
$$y_{i} = softmax\left( x_{i} \right) = \frac{\exp\left( x_{i} \right)}{\sum_{j = 1}^{n}{\exp\left( x_{j} \right)}}$$

**图片插入：**
![图片描述](image.png)

**图片标题（居中）：**
<div style="text-align: center; font-weight: bold;">图 1: 图片标题</div>
```

#### 步骤 2: 配置 Pandoc 参数
```bash
# 标准中文文档转换
pandoc input.md -o output.docx \
  --standalone \
  --variable CJKmainfont="SimSun" \
  --variable fontsize=12pt
```

**常用字体映射：**
- SimSun - 宋体（默认）
- SimHei - 黑体
- DengXian - 等线
- KaiTi - 楷体

#### 步骤 3: 图片处理规范

**图表生成要求：**
1. 使用 matplotlib 时设置中文字体
2. DPI 设置为 300（打印质量）
3. 图片格式使用 PNG
4. 每个图片必须有编号和标题

**matplotlib 中文配置：**
```python
from matplotlib import rcParams
rcParams['font.sans-serif'] = ['SimHei']  # 黑体
rcParams['axes.unicode_minus'] = False  # 解决负号显示
```

**图片标题格式：**
- 位置：图片下方
- 对齐：居中
- 格式：`图 X: 标题文字`
- 样式：加粗

#### 步骤 4: 格式验证清单

**必须验证的项目：**
- [ ] 中文字体正确（无乱码）
- [ ] 公式渲染正确
- [ ] 图片清晰且位置正确
- [ ] 图片标题居中且编号连续
- [ ] 标题层级清晰
- [ ] 字号统一（正文 12pt）
- [ ] 无多余下划线或格式
- [ ] 参考文献格式统一

### 4. 常见问题解决方案

#### 问题 1: 中文乱码
**原因：** 字体配置不正确
**解决：**
```bash
# 方法 1: 使用 SimSun
pandoc input.md -o output.docx --variable CJKmainfont="SimSun"

# 方法 2: 使用等线
pandoc input.md -o output.docx --variable CJKmainfont="DengXian"
```

#### 问题 2: 公式不显示
**原因：** 使用了不支持 LaTeX 的工具
**解决：** 必须使用 pandoc，不能用 python-docx

#### 问题 3: 图片标题不居中
**解决：** 使用 HTML div 标签
```markdown
<div style="text-align: center; font-weight: bold;">图 1: 标题</div>
```

#### 问题 4: 图表文字乱码
**原因：** matplotlib 未配置中文字体
**解决：**
```python
rcParams['font.sans-serif'] = ['SimHei']
```

### 5. 文档结构模板

**标准作业文档结构：**
```markdown
# [课程代码] 课程名称 - 作业 X

**课程名称：** [课程代码] 课程名称
**作业名称：** 作业 X
**提交邮箱：** email@example.com
**文件命名：** 作业 X-学号 - 姓名
**截止时间：** YYYY 年 MM 月 DD 日

---

## 1. 题目 1

**【参考答案】**

答案内容...

### 1.1 小标题

内容...

## 2. 题目 2

**【参考答案】**

### 2.1 理论推导

$$公式$$

### 2.2 代码实现

```python
# 代码
```

### 2.3 运行结果

结果说明...

### 2.4 讨论

讨论要点...

## 参考文献

1. 作者。(年份). 标题。期刊.
2. ...
```

### 6. 代码示例库

#### 示例 1: 生成带中文的图表
```python
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams

# 设置中文字体
rcParams['font.sans-serif'] = ['SimHei']
rcParams['axes.unicode_minus'] = False

# 创建图表
fig, ax = plt.subplots(figsize=(8, 6))
ax.plot([1, 2, 3], [1, 4, 9])
ax.set_title('图 1: 示例图表', fontsize=14)
ax.set_xlabel('X 轴', fontsize=12)
ax.set_ylabel('Y 轴', fontsize=12)

plt.savefig('figure.png', dpi=300, bbox_inches='tight')
```

#### 示例 2: Pandoc 转换脚本
```python
import pypandoc

output = pypandoc.convert_file(
    'input.md',
    'docx',
    outputfile='output.docx',
    extra_args=[
        '--standalone',
        '--variable', 'CJKmainfont=SimSun',
        '--variable', 'fontsize=12pt'
    ]
)
```

#### 示例 3: 文档验证脚本
```python
import pypandoc
import os

def verify_document(docx_path):
    content = pypandoc.convert_file(docx_path, 'markdown')
    
    checks = {
        '标题': '作业' in content,
        '公式': '$' in content,
        '图片': '图 1' in content,
        '参考文献': '参考文献' in content
    }
    
    for check, result in checks.items():
        status = '✓' if result else '✗'
        print(f'{status} {check}')
```

### 7. 最佳实践总结

#### 格式统一原则
1. **字体统一**：全文使用同一种中文字体
2. **字号统一**：正文 12pt，标题自动分级
3. **颜色统一**：默认黑色，避免彩色文字
4. **无多余装饰**：不使用下划线（除非必要）

#### 图片处理原则
1. **高质量**：DPI ≥ 300
2. **中文清晰**：使用正确字体
3. **编号连续**：图 1、图 2、图 3...
4. **标题居中**：使用 HTML div 标签

#### 公式处理原则
1. **使用 LaTeX**：所有数学公式用 LaTeX
2. **行内公式**：`$...$`
3. **行间公式**：`$$...$$`
4. **符号规范**：使用标准 LaTeX 符号

#### 文档组织原则
1. **结构清晰**：使用标题层级
2. **内容完整**：题目 + 答案 + 讨论
3. **格式一致**：所有题目格式统一
4. **易于阅读**：合理的段落间距

### 8. 工具安装指南

#### 必需工具
```bash
# Python 库
pip install pypandoc_binary
pip install python-docx
pip install matplotlib
pip install numpy

# Pandoc（自动安装或使用内置版本）
```

#### 可选工具（Windows）
```bash
# pywin32（完整 Word API 支持）
pip install pywin32

# 运行后处理脚本
python -m win32com.client.Dispatch("Word.Application")
```

### 9. 质量检查清单

**提交前必须检查：**
- [ ] 所有中文字符显示正确
- [ ] 所有公式渲染正确
- [ ] 所有图片清晰且位置正确
- [ ] 所有图片标题居中
- [ ] 图片标题编号连续
- [ ] 标题层级正确（1 → 1.1 → 1.1.1）
- [ ] 正文字号统一（12pt）
- [ ] 无多余格式（下划线、颜色等）
- [ ] 参考文献格式统一
- [ ] 文件大小合理（< 10MB）

### 10. 提交指南

**标准提交流程：**
1. 重命名文件：`作业 X-学号 - 姓名.docx`
2. 检查文件大小
3. 发送到指定邮箱
4. 确认邮件主题格式
5. 保留备份副本

**邮件主题模板：**
```
【作业提交】作业 X - 课程名称 - 学号 - 姓名
```

---

## 使用示例

### 示例 1: 创建作业文档
**用户：** "我需要完成神经网络作业，包含公式和图表"

**技能触发后的操作：**
1. 询问具体要求（题目数量、格式要求等）
2. 创建 Markdown 源文件
3. 生成必要的图表（配置中文字体）
4. 使用 pandoc 转换为 docx
5. 验证文档完整性
6. 提供提交指南

### 示例 2: 修复文档格式
**用户：** "这个 Word 文档中文乱码了"

**技能触发后的操作：**
1. 检查原文档格式
2. 识别乱码原因（字体问题）
3. 重新生成文档（指定正确字体）
4. 验证修复效果

### 示例 3: 统一文档格式
**用户：** "需要匹配原文档的格式（等线字体、黑色、无下划线）"

**技能触发后的操作：**
1. 分析原文档格式（使用 pywin32）
2. 提取字体、字号、样式信息
3. 创建 pandoc 配置
4. 生成匹配格式的新文档
5. 验证格式一致性

---

## 注意事项

### 常见错误
1. **忘记设置中文字体** → 导致乱码
2. **使用 python-docx 处理公式** → 公式丢失
3. **图片 DPI 过低** → 打印模糊
4. **忘记添加图片标题** → 格式不完整
5. **标题层级混乱** → 结构不清晰

### 性能优化
1. 使用 pandoc 而非 python-docx 处理复杂文档
2. 图片使用 PNG 格式（质量和大小平衡）
3. 批量生成时先测试单个样本
4. 使用脚本自动化重复任务

### 兼容性提醒
1. pywin32 仅 Windows 可用
2. 某些字体可能跨平台不兼容
3. 建议在目标平台上测试最终文档
4. 保留 Markdown 源文件便于修改

---

## 版本历史

- **v1.0** - 初始版本
  - 基础文档创建流程
  - Pandoc 转换配置
  - 图表生成规范
  - 常见问题解决方案

---

## 参考资源

### 官方文档
- Pandoc: https://pandoc.org/
- Matplotlib: https://matplotlib.org/
- python-docx: https://python-docx.readthedocs.io/

### 字体参考
- Windows 中文字体：SimSun, SimHei, DengXian, KaiTi
- macOS 中文字体：STSong, Heiti, PingFang
- Linux 中文字体：WenQuanYi, Noto CJK

### 示例文件
- Markdown 模板：homework_template.md
- 图表生成脚本：generate_figures.py
- 验证脚本：verify_document.py
