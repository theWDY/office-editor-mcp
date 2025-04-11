# Office文档处理助手 MCP服务器

[![EN](https://img.shields.io/badge/Language-English-blue)](README.md)
[![CN](https://img.shields.io/badge/语言-中文-red)](README_CN.md)

![MCP Server](https://img.shields.io/badge/MCP-Server-blue)
![Python](https://img.shields.io/badge/Python-3.7+-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

基于MCP(Model Context Protocol)的Office文档处理助手，支持在MCP Client中创建和编辑Word、Excel、PowerPoint文档，无需离开AI助手环境。

## 概述

Office-Editor-MCP实现了[Model Context Protocol](https://modelcontextprotocol.io/)标准，将Office文档操作暴露为工具和资源。它为AI助手和Microsoft Office文档之间搭建了桥梁，让您能够通过AI助手创建、编辑、格式化和分析各类Office文档。

<!-- 建议: 在此处添加使用截图 -->

## 功能特性

### Word文档操作

#### 文档管理
- 创建新的Word文档，支持元数据（标题、作者等）
- 提取文本内容和分析文档结构
- 查看文档属性和统计信息
- 列出目录中的可用文档
- 创建文档副本

#### 内容创建
- 添加不同级别的标题
- 插入段落（支持可选样式）
- 创建自定义数据表格
- 添加图片（支持比例缩放）
- 插入分页符

#### 文本格式化
- 格式化特定文本部分（粗体、斜体、下划线）
- 更改文本颜色和字体属性
- 应用自定义样式到文本元素
- 在整个文档中搜索和替换文本

### Excel表格操作

#### 表格管理
- 创建新的Excel工作簿
- 打开现有Excel文件
- 添加/删除/重命名工作表

#### 数据处理
- 读写单元格内容
- 插入/删除行列
- 数据排序和筛选
- 公式与函数应用

### PowerPoint演示文稿操作

#### 演示文稿管理
- 创建新的PowerPoint演示文稿
- 添加/删除/重排幻灯片
- 设置幻灯片主题和背景

#### 内容编辑
- 添加文本和图形元素
- 插入表格和图表
- 添加动画和转场效果

### 高级功能

- OCR识别（从图片提取文本）
- 文档比较（对比两个文档的差异）
- 文档翻译
- 文档加密和解密
- 表格数据导入导出（与数据库交互）

## 安装指南

### 前提条件
- Python 3.7 或更高版本
- pip 包管理器
- Microsoft Office或兼容组件（如python-docx, openpyxl）

### 基本安装

```bash
# 克隆仓库
git clone https://github.com/theWDY/office-editor-mcp.git
cd office-editor-mcp

# 安装依赖
pip install -r requirements.txt
```

## 配置说明

### 在Cursor中配置

#### 方法一：UI配置

1. 打开Cursor
2. 进入设置 > Features > MCP
3. 点击"+ Add New MCP Server"
4. 填写配置信息：
   - 名称：`Office助手`（可根据喜好修改）
   - 类型：选择`stdio`
   - 命令：输入运行服务器的完整路径，例如：
     ```
     python /path/to/office_server.py
     ```
     注意替换为您实际的文件路径

#### 方法二：JSON配置文件（推荐）

1. 在项目目录中创建 `.cursor` 文件夹（如果不存在）
2. 在该文件夹中创建 `mcp.json` 文件，内容如下：

```json
{
  "mcpServers": {
    "office-assistant": {
      "command": "python",
      "args": ["/path/to/office_server.py"],
      "env": {}
    }
  }
}
```

### 在Claude for Desktop中配置

1. 编辑Claude配置文件：
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`

2. 添加以下配置：

```json
{
  "mcpServers": {
    "office-document-server": {
      "command": "python",
      "args": [
        "/path/to/office_server.py"
      ]
    }
  }
}
```

3. 重启Claude使配置生效。

## 使用示例

配置完成后，您可以向AI助手发出如下指令：

### Word文档操作
- "创建一个名为'季度报告.docx'的文档，包含标题页"
- "在文档中添加一个标题和三个段落"
- "插入一个4x4的销售数据表格"
- "将第2段中的'重要'一词设为粗体红色"
- "搜索并替换所有'旧术语'为'新术语'"

### Excel操作
- "创建一个新的Excel工作簿，命名为'财务分析.xlsx'"
- "在A1单元格插入'季度销售额'作为标题"
- "创建一个包含部门销售数据的表格，并计算总和"
- "为销售数据创建一个柱状图"
- "对B列数据进行降序排序"

### PowerPoint操作
- "创建一个名为'项目演示.pptx'的演示文稿"
- "添加一个标题为'项目概述'的新幻灯片"
- "在第2张幻灯片中插入公司logo图片"
- "为标题添加飞入动画效果"

## API参考

### Word文档操作

```python
# 文档创建与属性
create_document(filename, title=None, author=None)
get_document_info(filename)
get_document_text(filename)
get_document_outline(filename)
list_available_documents(directory=".")
copy_document(source_filename, destination_filename=None)

# 内容添加
add_heading(filename, text, level=1)
add_paragraph(filename, text, style=None)
add_table(filename, rows, cols, data=None)
add_picture(filename, image_path, width=None)
add_page_break(filename)

# 文本格式化
format_text(filename, paragraph_index, start_pos, end_pos, bold=None, 
            italic=None, underline=None, color=None, font_size=None, font_name=None)
search_and_replace(filename, find_text, replace_text)
delete_paragraph(filename, paragraph_index)
create_custom_style(filename, style_name, bold=None, italic=None, 
                    font_size=None, font_name=None, color=None, base_style=None)
```

### Excel操作

```python
# 工作簿操作
create_workbook(filename)
open_workbook(filename)
save_workbook(filename, new_filename=None)
add_worksheet(filename, sheet_name=None)
list_worksheets(filename)

# 单元格操作
read_cell(filename, sheet_name, cell_reference)
write_cell(filename, sheet_name, cell_reference, value)
format_cell(filename, sheet_name, cell_reference, **format_args)
```

### PowerPoint操作

```python
# 演示文稿操作
create_presentation(filename)
open_presentation(filename)
save_presentation(filename, new_filename=None)
add_slide(filename, layout=None)
```

## 故障排除

### 常见问题

1. **缺少样式**
   - 部分文档可能缺少标题和表格操作所需的样式
   - 服务器将尝试创建缺少的样式或使用直接格式化
   - 为获得最佳效果，请使用具有标准Office样式的模板

2. **权限问题**
   - 确保服务器有权读/写文档路径
   - 使用`copy_document`函数创建锁定文档的可编辑副本
   - 操作失败时检查文件所有权和权限

3. **图片插入问题**
   - 使用图片的绝对路径
   - 验证图片格式兼容性（推荐JPEG、PNG）
   - 检查图片文件大小和权限

### 调试

通过设置环境变量启用详细日志记录：

```bash
export MCP_DEBUG=1  # Linux/macOS
set MCP_DEBUG=1     # Windows
```

## 实施进度

- ✅ 构建MCP服务器基础框架
- ✅ 实现与AI助手的成功集成
- ✅ Word文档的基本操作
- ✅ Excel工作簿的基本操作
- ✅ PowerPoint演示文稿的基本操作
- ✅ 高级功能完善
- ✅ 性能优化
- ✅ 跨平台兼容性测试

## 贡献指南

欢迎贡献！请随时提交Pull Request。

1. Fork本仓库
2. 创建您的特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交您的更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 开启一个Pull Request

## 许可证

本项目采用MIT许可证 - 详情请参阅[LICENSE](LICENSE)文件。

## 致谢

- [Model Context Protocol](https://modelcontextprotocol.io/)提供协议规范
- [python-docx](https://python-docx.readthedocs.io/)提供Word文档处理
- [openpyxl](https://openpyxl.readthedocs.io/)提供Excel处理
- [python-pptx](https://python-pptx.readthedocs.io/)提供PowerPoint处理

---

*注意：此服务器与您系统上的文档文件交互。在AI助手或其他MCP客户端中确认操作前，请始终验证所请求操作的适当性。*
