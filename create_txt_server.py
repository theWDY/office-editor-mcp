"""
简易TXT文件创建MCP服务器
这是一个简化版的服务器，只提供创建TXT文件的功能
作为完整Office Editor服务器的简化版示例
"""

import os
from mcp.server.fastmcp import FastMCP

# 创建MCP服务器
mcp = FastMCP("txt creator")

@mcp.tool()
def create_empty_txt(filename: str) -> str:
    """
    在指定路径上创建一个空白的TXT文件。
    
    Args:
        filename: 要创建的文件名 (不需要包含.txt扩展名)
    
    Returns:
        包含操作结果的消息
    """
    # 确保文件名有.txt扩展名
    if not filename.lower().endswith('.txt'):
        filename += '.txt'
    
    # 从环境变量获取输出路径，如果未设置则使用默认桌面路径
    output_path = os.environ.get('OFFICE_EDIT_PATH')
    if not output_path:
        output_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    
    # 创建完整的文件路径
    file_path = os.path.join(output_path, filename)
    
    try:
        # 创建输出目录（如果不存在）
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 创建空白文件
        with open(file_path, 'w', encoding='utf-8') as f:
            pass
        return f"成功在 {output_path} 创建了空白文件: {filename}"
    except Exception as e:
        return f"创建文件时出错: {str(e)}"

@mcp.tool()
def create_txt_with_content(filename: str, content: str) -> str:
    """
    在指定路径上创建一个包含指定内容的TXT文件。
    
    Args:
        filename: 要创建的文件名 (不需要包含.txt扩展名)
        content: 文件内容
    
    Returns:
        包含操作结果的消息
    """
    # 确保文件名有.txt扩展名
    if not filename.lower().endswith('.txt'):
        filename += '.txt'
    
    # 从环境变量获取输出路径，如果未设置则使用默认桌面路径
    output_path = os.environ.get('OFFICE_EDIT_PATH')
    if not output_path:
        output_path = os.path.join(os.path.expanduser('~'), 'Desktop')
    
    # 创建完整的文件路径
    file_path = os.path.join(output_path, filename)
    
    try:
        # 创建输出目录（如果不存在）
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 创建并写入内容到文件
        with open(file_path, 'w', encoding='utf-8') as f:
            f.write(content)
        return f"成功在 {output_path} 创建了文件: {filename}，并写入内容"
    except Exception as e:
        return f"创建文件时出错: {str(e)}"

if __name__ == "__main__":
    # 运行MCP服务器
    print("启动TXT文件创建服务器...")
    mcp.run()