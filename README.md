# Office文档处理助手 MCP服务器

这是一个功能强大的MCP（Model Context Protocol）服务器，旨在提供全面的Microsoft Office文档处理能力。通过与MCP Client（如Claude Desktop、Cursor等）的集成，可以实现对Word、Excel、PowerPoint等文档的创建、编辑和管理操作，无需离开Client环境。

## 项目背景与目标

在软件开发过程中，开发者经常需要同时处理代码和文档。传统工作流程需要在IDE和Office应用程序之间不断切换，降低了工作效率。本项目旨在通过MCP协议将Office文档处理能力集成到Cursor中，使开发者能够在同一环境中完成所有工作。

## 功能需求

### 1. Word文档操作

#### 1.1 文档创建与管理
- 创建新的Word文档，支持指定文件名和保存位置
- 打开现有Word文档
- 保存文档（另存为不同格式，如.docx、.doc、.pdf）
- 关闭文档

#### 1.2 内容编辑
- 插入文本内容（支持指定位置插入）
- 插入标题（支持多级标题H1-H6）
- 编辑现有内容（查找替换、删除内容）
- 格式设置（字体、大小、颜色、加粗、斜体、下划线等）
- 段落设置（对齐方式、行距、段前段后间距）

#### 1.3 高级功能
- 插入图片（支持本地图片和网络图片）
- 插入表格（创建表格、编辑单元格内容）
- 插入目录（自动生成目录、更新目录）
- 添加页眉页脚
- 设置页面布局（纸张大小、页边距、方向）
- 添加批注和修订
- 文档合并

### 2. Excel表格操作

#### 2.1 表格创建与管理
- 创建新的Excel工作簿
- 打开现有Excel文件
- 保存工作簿（支持.xlsx、.xls、.csv格式）
- 添加/删除/重命名工作表

#### 2.2 内容编辑
- 读取单元格内容
- 写入单元格内容（支持文本、数字、日期等类型）
- 清除单元格内容
- 设置单元格格式（字体、背景色、边框、对齐方式等）
- 合并/拆分单元格

#### 2.3 数据处理
- 插入/删除行列
- 数据排序（升序/降序）
- 数据筛选
- 数据透视表创建
- 数据有效性设置
- 条件格式设置
- 批量数据处理（批量填充、批量替换）

#### 2.4 公式与函数
- 插入基本计算公式（加减乘除）
- 插入高级函数（SUM、AVERAGE、COUNT、IF、VLOOKUP等）
- 创建和编辑图表（柱状图、折线图、饼图等）

### 3. PowerPoint演示文稿操作

#### 3.1 演示文稿创建与管理
- 创建新的PowerPoint演示文稿
- 打开现有演示文稿
- 保存演示文稿（支持.pptx、.ppt、.pdf格式）

#### 3.2 幻灯片操作
- 添加新幻灯片（支持不同版式）
- 删除幻灯片
- 调整幻灯片顺序
- 设置幻灯片主题和背景

#### 3.3 内容编辑
- 添加文本框和编辑文本内容
- 插入图片和形状
- 插入表格和图表
- 添加动画效果（入场、强调、退场）
- 设置幻灯片切换效果
- 添加备注

### 4. 通用功能

#### 4.1 文件管理
- 批量创建文档
- 文件重命名
- 文件复制和移动
- 文件格式转换（Office格式之间的互转，转PDF等）

#### 4.2 模板功能
- 提供常用文档模板（简历、报告、表格等）
- 保存自定义模板
- 基于模板创建新文档

#### 4.3 高级功能
- OCR识别（从图片提取文本）
- 文档比较（对比两个文档的差异）
- 文档翻译
- 文档加密和解密
- 表格数据导入导出（与数据库交互）

## 技术要求

### 1. 系统兼容性
- 支持Windows操作系统
- 支持macOS操作系统（可选）
- 支持Linux操作系统（可选）

### 2. 依赖条件
- 用户本地已安装Microsoft Office或兼容组件
- Python 3.7+
- MCP Python SDK

### 3. 安全与权限
- 文档操作需用户明确授权
- 敏感操作需确认机制
- 支持只读模式操作

### 4. 性能要求
- 响应时间：常规操作≤1秒
- 大型文档处理≤5秒
- 批量操作进度可视化

## 用户体验

### 1. 交互方式
- 通过自然语言指令操作（例如："在桌面创建一个名为'季度报告'的Word文档"）
- 支持复合指令（例如："创建Excel表格，插入销售数据，并生成柱状图"）
- 支持上下文相关操作（例如："在上一个单元格下方插入新行"）

### 2. 错误处理
- 提供清晰的错误提示
- 建议解决方案
- 支持操作撤销

### 3. 辅助功能
- 操作历史记录
- 常用操作快捷指令
- 批处理脚本支持

## 实施阶段

### 第一阶段：基础框架（已完成）
- ✅ 构建MCP服务器基础框架
- ✅ 实现与Cursor的成功集成
- ✅ 验证基本功能（简单的TXT文件创建）

### 第二阶段：Word文档基础功能（已完成）
- ✅ 实现Word文档的创建、打开、保存
- ✅ 基本文本编辑功能
- ✅ 简单格式设置

### 第三阶段：Excel基础功能
- ✅实现Excel工作簿的创建、打开、保存
- ✅单元格内容读写
- ✅基本格式设置

### 第四阶段：PowerPoint基础功能
- ✅实现PowerPoint演示文稿的创建、打开、保存
- ✅幻灯片管理
- ✅基本内容编辑

### 第五阶段：高级功能与优化
- ✅实现各应用的高级功能
- ✅性能优化
- ✅用户体验提升

## 配置方法

### 在Cursor中配置（方法一：UI配置）

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

### 在Cursor中配置（方法二：JSON配置文件，推荐）

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

注意替换 `/path/to/office_server.py` 为您实际的文件路径。

3. 重启Cursor使配置生效。

## 许可

[MIT许可](LICENSE)
