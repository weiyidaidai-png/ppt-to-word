# Word 转 PPT 自动化程序

一个可以将 Word 文档自动转换为 PowerPoint 演示文稿的 Python 工具，支持图形界面和命令行两种模式。

## 功能特性

- 📄 读取 Word 文档中的所有内容
- 🔍 智能识别标题和正文（支持样式和字号判断）
- 🎨 自动生成 PPT 幻灯片
- 📝 保持原有的段落结构和顺序
- 📊 支持多种标题样式和字号
- 🎯 空行和特殊格式处理
- 🖥️ 图形用户界面 (GUI) 和命令行界面 (CLI) 双模式
- 📁 可视化文件选择和保存对话框

## 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 1. 图形界面模式（推荐）

直接运行程序，无需命令行参数：

```bash
python word_to_ppt.py
```

程序会打开文件选择对话框，引导您选择 Word 文件和设置保存位置。

### 2. 命令行方式

```bash
python word_to_ppt.py [Word 文档路径] [PPT 保存路径]
```

### 3. 交互式命令行方式

如果只提供部分参数，程序会提示您输入剩余信息：

```bash
python word_to_ppt.py
```

## 快速测试

我们提供了一个完整的示例供您测试：

```bash
# 创建示例 Word 文档
python example_script.py

# 使用命令行转换
python word_to_ppt.py "example.docx" "output.pptx"

# 或者使用图形界面转换
python word_to_ppt.py
```

## 工作原理

1. **内容解析**：程序读取 Word 文档，根据字体大小或样式识别标题
2. **章节划分**：每个标题对应一个章节，标题后的所有内容作为该章节的正文
3. **PPT 生成**：为每个章节创建一张幻灯片，标题放在标题区域，正文放在内容区域
4. **保存文件**：生成的 PPT 自动保存到指定路径

## 标题识别规则

- 使用 "Heading 1"、"Heading 2" 或 "Heading 3" 样式的段落
- 字号大于等于 14 号的段落
- 如果没有识别到标题，会将第一段落作为标题

## 注意事项

- 请确保输入的 Word 文件路径正确且文件不为空
- 生成的 PPT 文件会覆盖同名文件，请谨慎操作
- 对于复杂格式的 Word 文档，可能需要手动调整结果

## 文件说明

- `word_to_ppt.py` - 主程序文件，实现 Word 转 PPT 的核心功能
- `example_script.py` - 创建示例 Word 文档的脚本
- `test_program.py` - 测试程序功能的脚本
- `requirements.txt` - 所需依赖包列表
- `README.md` - 项目说明文档

