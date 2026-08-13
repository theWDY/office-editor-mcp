# 安装和使用指南

## 1. 安装依赖

本项目需要以下依赖：

1. Python 3.10或更高版本
2. MCP SDK和Office文档处理依赖（由`requirements.txt`统一安装）
3. Microsoft Office（仅Windows COM增强功能需要）
4. Tesseract OCR（仅OCR功能需要）

### 使用pip安装

```bash
pip install -r requirements.txt
```

## 2. 服务器配置

项目提供四个独立的服务器实现：

1. `word_server.py` - Word文档操作
2. `excel_server.py` - Excel工作簿操作
3. `powerpoint_server.py` - PowerPoint演示文稿操作
4. `general_server.py` - OCR、比较、翻译、加密和批处理等通用功能

仓库中没有`office_server.py`。客户端可以只配置需要的服务器，也可以参考
`.cursor/mcp.json`一次配置全部四个服务器。

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
     python C:/path/to/office-editor-mcp/word_server.py
     ```

### 方法二：通过配置文件配置（推荐）

1. 在项目目录中创建 `.cursor` 文件夹（如果不存在）
2. 在该文件夹中创建 `mcp.json` 文件，内容如下：

```json
{
  "mcpServers": {
    "office-word": {
      "command": "python",
      "args": ["C:/path/to/office-editor-mcp/word_server.py"],
      "env": {
        "OFFICE_EDIT_PATH": "C:/path/to/output/folder"
      }
    }
  }
}
```

请替换路径为实际的绝对路径。`OFFICE_EDIT_PATH`指定文档工作目录。通用服务器只允许
访问该目录内的路径；移动和删除默认禁用，只有显式设置
`OFFICE_ALLOW_DESTRUCTIVE=true`后才会开放。

新加密文件使用随机盐和Scrypt密钥派生；旧版加密文件仍可读取。批处理请求最多处理
100个文件，并限制为最多16个工作线程。

3. 重启Cursor使配置生效。

## 4. 功能测试

运行静态编译和安全边界测试：

```bash
python -m compileall -q .
python -m unittest discover -s tests -v
```

## 5. 使用示例

一旦服务器配置好并在Cursor中启用，您可以通过以下方式使用：

1. 在Cursor中，与AI助手进行对话
2. 要求AI助手帮您创建或编辑Word文档，例如：
   - "创建一个名为'项目计划'的Word文档"
   - "在该文档中添加一个标题'2025年度项目计划'"
   - "添加一个3行4列的表格"

AI助手将使用配置好的MCP服务器执行这些操作，并在指定的目录中创建和编辑文档。
