# 安装和使用指南

## 1. 安装依赖

本项目需要以下依赖：

1. Python 3.7或更高版本
2. MCP SDK
3. python-docx库（用于Word文档操作）
4. Pillow库（用于图像处理）
5. pywin32（可选，Windows系统用于增强功能）

### 使用pip安装

```bash
pip install -r requirements.txt
```

## 2. 服务器配置

项目提供了两个主要的服务器实现：

1. `create_txt_server.py` - 简单文本文件创建服务器（仅用于功能测试）
2. `office_server.py` - 完整的Office文档处理服务器（需要下载完整项目）

## 3. 在Cursor中配置

### 方法一：通过UI配置

1. 打开Cursor
2. 进入设置 > Features > MCP
3. 点击"+ Add New MCP Server"
4. 填写配置信息：
   - 名称：`Office助手`（可自定义）
   - 类型：选择`stdio`
   - 命令：输入运行服务器的完整路径，例如：
     ```
     python C:/path/to/office_server.py
     ```

### 方法二：通过配置文件配置（推荐）

1. 在项目目录中创建 `.cursor` 文件夹（如果不存在）
2. 在该文件夹中创建 `mcp.json` 文件，内容如下：

```json
{
  "mcpServers": {
    "office-editor": {
      "command": "python",
      "args": ["C:/path/to/office_server.py"],
      "env": {
        "OFFICE_EDIT_PATH": "C:/path/to/output/folder"
      }
    }
  }
}
```

请替换路径为您实际的文件路径。`OFFICE_EDIT_PATH`环境变量指定文档的默认保存位置，如不设置则默认为桌面。

3. 重启Cursor使配置生效。

## 4. 功能测试

项目包含一个测试脚本 `test_server.py`，它可以：

1. 检查MCP SDK是否正确安装
2. 验证服务器脚本是否存在
3. 尝试启动服务器并测试基本功能

运行测试脚本：

```bash
python test_server.py
```

如果测试通过，您会看到一系列成功消息，表明服务器已正确配置。

## 5. 使用示例

一旦服务器配置好并在Cursor中启用，您可以通过以下方式使用：

1. 在Cursor中，与AI助手进行对话
2. 要求AI助手帮您创建或编辑Word文档，例如：
   - "创建一个名为'项目计划'的Word文档"
   - "在该文档中添加一个标题'2025年度项目计划'"
   - "添加一个3行4列的表格"

AI助手将使用配置好的MCP服务器执行这些操作，并在指定的目录中创建和编辑文档。