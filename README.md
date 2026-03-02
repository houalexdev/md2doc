# md2doc

一个将极简 Markdown 转换为 DOCX 文件的工具，支持自定义样式配置。

## 功能特性

- ✅ 支持标题（H1-H5）
- ✅ 支持加粗文本（**加粗**）
- ✅ 支持表格（带表头样式）
- ✅ 支持列表（无序转有序，多级列表）
- ✅ 支持代码块（带语法高亮样式）
- ✅ 自定义样式配置（styles.json）
- ✅ 命令行操作，简单易用

## 安装

1. 确保已安装 Python 3.6+：
   ```bash
   python --version
   ```

2. 安装依赖库：
   ```bash
   pip install python-docx
   ```

3. 克隆或下载项目：
   ```bash
   git clone <repository-url>
   cd md2doc
   ```

## 使用方法

### 基本用法

```bash
python md2doc.py input.md output.docx
```

- `input.md`：输入的 Markdown 文件路径
- `output.docx`：输出的 DOCX 文件路径

### 示例

```bash
# 使用默认输入文件（input.md）和输出文件（output.docx）
python md2doc.py

# 指定自定义输入和输出文件
python md2doc.py my_document.md result.docx
```

## 支持的 Markdown 语法

### 标题

```markdown
# 一级标题
## 二级标题
### 三级标题
#### 四级标题
##### 五级标题
```

### 加粗文本

```markdown
普通文本 **加粗文本** 普通文本
```

### 表格

```markdown
| 列1 | 列2 | 列3 |
|-----|-----|-----|
| 行1 | 行1 | 行1 |
| 行2 | 行2 | 行2 |
```

### 列表

```markdown
- 一级列表项
  - 二级列表项
    - 三级列表项
- 一级列表项
```

### 代码块

```markdown
```python
def hello():
    print("Hello, World!")
```
```

## 样式配置

所有样式配置都在 `styles.json` 文件中定义，无需修改 Python 代码。

### 配置项说明

#### 段落样式

```json
{
  "paragraph_styles": [
    {
      "id": "StyleID",
      "name": "样式名称",
      "base_style": null,
      "next_style": "Normal",
      "alignment": "LEFT",
      "outline_lvl": 1,
      "font": {
        "name": "宋体",
        "size": 12,
        "bold": false,
        "color": "000000"
      },
      "paragraph_format": {
        "space_before": 0,
        "space_after": 0,
        "line_spacing": 1.5,
        "left_indent": 0,
        "first_line_indent": 24
      }
    }
  ]
}
```

#### 字符样式

```json
{
  "character_styles": [
    {
      "id": "Strong",
      "name": "强调",
      "font": {
        "bold": true,
        "color": "000000"
      }
    }
  ]
}
```

#### 表格样式

```json
{
  "table_styles": [
    {
      "id": "CustomTable",
      "name": "自定义表格",
      "font": {
        "name": "宋体",
        "size": 10.5,
        "color": "000000"
      },
      "paragraph_format": {
        "space_after": 0,
        "line_spacing": 1.0
      },
      "table_properties": {
        "border": "single",
        "border_width": "1pt",
        "border_color": "000000"
      },
      "header_row": {
        "font": {
          "bold": true,
          "color": "000000"
        },
        "fill_color": "D9D9D9",
        "paragraph_format": {
          "alignment": "CENTER"
        }
      }
    }
  ]
}
```

### 预设样式

项目默认提供了以下样式：

- **标题 1-5**：不同级别的标题样式
- **正文**：默认文本样式
- **强调**：加粗文本样式
- **CodeBlock**：代码块样式
- **自定义表格**：带表头样式的表格

## 示例输入输出

### 输入（input.md）

```markdown
# 项目介绍

## 功能特性

- **核心功能**：将 Markdown 转换为 DOCX
- **样式支持**：自定义标题、正文、表格样式
- **代码块**：支持代码高亮样式

## 使用示例

| 命令 | 说明 |
|------|------|
| `python md2doc.py` | 使用默认文件 |
| `python md2doc.py input.md output.docx` | 指定输入输出 |

## 代码示例

```python
# 简单的转换示例
import json
from docx import Document

def md_to_docx(md_text, out_file):
    # 转换逻辑
    pass
```
```

### 输出（output.docx）

转换后的 DOCX 文件将包含：
- 格式化的标题
- 加粗的文本
- 带表头样式的表格
- 带样式的代码块
- 有序列表（自动转换无序列表）

## 注意事项

1. 仅支持极简 Markdown 语法，复杂语法可能无法正确转换
2. 表格需要使用标准的 Markdown 表格格式
3. 列表转换时会自动将无序列表转为有序列表
4. 代码块需要用三个反引号（```）包裹
5. 所有样式都在 styles.json 中配置，无需修改 Python 代码

## 许可证

MIT License

## 贡献

欢迎提交 Issue 和 Pull Request 来改进这个项目！

## 联系方式

如有问题或建议，请通过以下方式联系：
- 项目地址：<repository-url>
- Issue 页面：<repository-url>/issues