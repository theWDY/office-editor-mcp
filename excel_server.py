"""
MCP Server for Excel Operations

This server provides tools to create, edit and manage Excel workbooks.
It's implemented using the Model Context Protocol (MCP) Python SDK.
"""

import os
import sys
import io
from mcp.server.fastmcp import FastMCP
from typing import Optional, List, Dict, Any, Union, Tuple

# 标记库是否已安装
openpyxl_installed = True

# 尝试导入openpyxl库，如果没有安装则标记为未安装但不退出
try:
    import openpyxl
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
    from openpyxl.styles.colors import Color
    from openpyxl.utils import get_column_letter, column_index_from_string
except ImportError:
    print("警告: 未检测到openpyxl库，Excel功能将不可用")
    print("请使用以下命令安装: pip install openpyxl")
    openpyxl_installed = False

# 尝试导入Pandas库，用于数据处理
pandas_installed = True
try:
    import pandas as pd
    import numpy as np
except ImportError:
    print("警告: 未检测到pandas库，高级数据处理功能将受限")
    print("请使用以下命令安装: pip install pandas numpy")
    pandas_installed = False

# 创建一个MCP服务器，保持名称与配置文件一致
mcp = FastMCP("office editor")

@mcp.tool()
def create_excel_workbook(filename: str) -> str:
    """
    创建一个新的Excel工作簿。
    
    Args:
        filename: 要创建的文件名 (不需要包含.xlsx扩展名)
    
    Returns:
        包含操作结果的消息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法创建Excel工作簿，请先安装openpyxl库: pip install openpyxl"
    
    # 确保文件名有.xlsx扩展名
    if not filename.lower().endswith('.xlsx'):
        filename += '.xlsx'
    
    # 从环境变量获取输出路径，如果未设置则使用默认桌面路径
    output_path = os.environ.get('OFFICE_EDIT_PATH')
    if not output_path:
        output_path = os.path.join(os.path.expanduser('~'), '桌面')
    
    # 创建完整的文件路径
    file_path = os.path.join(output_path, filename)
    
    try:
        # 创建输出目录（如果不存在）
        os.makedirs(os.path.dirname(file_path), exist_ok=True)
        
        # 创建新的Excel工作簿
        wb = Workbook()
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在 {output_path} 创建了Excel工作簿: {filename}"
    except Exception as e:
        return f"创建Excel工作簿时出错: {str(e)}"

@mcp.tool()
def open_excel_workbook(file_path: str) -> str:
    """
    打开一个现有的Excel工作簿并读取其基本信息。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
    
    Returns:
        工作簿的基本信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法打开Excel工作簿，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path, data_only=True)
        
        # 获取工作表信息
        sheet_names = wb.sheetnames
        active_sheet = wb.active.title
        
        # 构建工作簿信息
        workbook_info = (
            f"文件名: {os.path.basename(file_path)}\n"
            f"工作表数量: {len(sheet_names)}\n"
            f"工作表列表: {', '.join(sheet_names)}\n"
            f"当前活动工作表: {active_sheet}\n"
        )
        
        return workbook_info
    except Exception as e:
        return f"打开Excel工作簿时出错: {str(e)}"

@mcp.tool()
def save_excel_workbook(file_path: str, format_type: str = "xlsx", new_filename: str = None) -> str:
    """
    保存Excel工作簿，可选择保存为不同格式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        format_type: 保存格式，可选值: "xlsx", "xls", "csv"
        new_filename: 新文件名(不含扩展名)，如果不提供则使用原文件名
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法保存Excel工作簿，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证格式类型
    supported_formats = ["xlsx", "xls", "csv"]
    if format_type not in supported_formats:
        return f"错误: 不支持的格式类型 '{format_type}'，支持的格式有: {', '.join(supported_formats)}"
    
    try:
        # 获取原文件名（不含扩展名）和目录
        file_dir = os.path.dirname(file_path)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        
        # 使用新文件名（如果提供）
        if new_filename:
            file_name = new_filename
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 创建新文件路径
        new_file_path = os.path.join(file_dir, f"{file_name}.{format_type}")
        
        # 根据格式类型保存
        if format_type == "xlsx":
            wb.save(new_file_path)
        
        elif format_type == "xls":
            # 尝试不同的方法保存为xls格式
            
            # 方法1: 使用win32com.client (需要安装pywin32)
            try:
                import win32com.client
                
                # 先保存为临时xlsx文件
                temp_path = os.path.join(file_dir, f"{file_name}_temp.xlsx")
                wb.save(temp_path)
                
                # 使用Excel应用程序打开并另存为xls
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                
                try:
                    workbook = excel.Workbooks.Open(temp_path)
                    workbook.SaveAs(new_file_path, FileFormat=56)  # 56 = xls格式
                    workbook.Close()
                finally:
                    excel.Quit()
                
                # 删除临时文件
                if os.path.exists(temp_path):
                    os.remove(temp_path)
                
            except (ImportError, Exception) as e:
                # 如果win32com方法失败，尝试使用第二种方法
                try:
                    # 方法2: 使用xlwt库
                    import xlwt
                    
                    # 创建新的xls工作簿
                    xls_wb = xlwt.Workbook()
                    
                    # 遍历所有工作表
                    for sheet_name in wb.sheetnames:
                        ws = wb[sheet_name]
                        xls_ws = xls_wb.add_sheet(sheet_name)
                        
                        # 复制数据
                        for row_idx, row in enumerate(ws.iter_rows(values_only=True)):
                            for col_idx, cell_value in enumerate(row):
                                xls_ws.write(row_idx, col_idx, cell_value)
                    
                    # 保存xls工作簿
                    xls_wb.save(new_file_path)
                    
                except ImportError:
                    # 如果xlwt也不可用，则返回错误信息
                    return f"错误: 无法保存为XLS格式，未找到必要的库。请安装pywin32或xlwt库: pip install pywin32 或 pip install xlwt"
        
        elif format_type == "csv":
            # 只保存第一个工作表为CSV
            ws = wb.active
            
            with open(new_file_path, 'w', newline='', encoding='utf-8') as f:
                import csv
                writer = csv.writer(f)
                for row in ws.iter_rows(values_only=True):
                    writer.writerow(row)
        
        return f"成功将工作簿保存为 {format_type} 格式: {os.path.basename(new_file_path)}"
    
    except Exception as e:
        return f"保存工作簿时出错: {str(e)}"

@mcp.tool()
def add_worksheet(file_path: str, sheet_name: str) -> str:
    """
    在Excel工作簿中添加新的工作表。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 新工作表的名称
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法添加工作表，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否已存在
        if sheet_name in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 已存在"
        
        # 创建新工作表
        ws = wb.create_sheet(title=sheet_name)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作簿 {os.path.basename(file_path)} 中添加工作表: {sheet_name}"
    except Exception as e:
        return f"添加工作表时出错: {str(e)}"

@mcp.tool()
def delete_worksheet(file_path: str, sheet_name: str) -> str:
    """
    从Excel工作簿中删除指定的工作表。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 要删除的工作表名称
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法删除工作表，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 删除工作表
        del wb[sheet_name]
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功从工作簿 {os.path.basename(file_path)} 中删除了工作表: {sheet_name}"
    except Exception as e:
        return f"删除工作表时出错: {str(e)}"

@mcp.tool()
def rename_worksheet(file_path: str, old_name: str, new_name: str) -> str:
    """
    重命名Excel工作簿中的工作表。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        old_name: 要重命名的工作表当前名称
        new_name: 工作表的新名称
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法重命名工作表，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查旧工作表名称是否存在
        if old_name not in wb.sheetnames:
            return f"错误: 工作表 '{old_name}' 不存在"
        
        # 检查新工作表名称是否已存在
        if new_name in wb.sheetnames:
            return f"错误: 工作表 '{new_name}' 已存在"
        
        # 重命名工作表
        ws = wb[old_name]
        ws.title = new_name
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功将工作簿 {os.path.basename(file_path)} 中的工作表 '{old_name}' 重命名为 '{new_name}'"
    except Exception as e:
        return f"重命名工作表时出错: {str(e)}"

@mcp.tool()
def read_cell(file_path: str, sheet_name: str, cell: str) -> str:
    """
    读取Excel工作簿中指定单元格的内容。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell: 单元格地址，如"A1", "B2"等
    
    Returns:
        单元格的内容
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法读取单元格，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path, data_only=True)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 读取单元格内容
        cell_value = ws[cell].value
        
        if cell_value is None:
            return f"单元格 {cell} 为空"
        else:
            return f"单元格 {cell} 的内容: {cell_value}"
    
    except Exception as e:
        return f"读取单元格时出错: {str(e)}"

@mcp.tool()
def read_cell_range(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """
    读取Excel工作簿中指定单元格范围的内容。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        start_cell: 起始单元格地址，如"A1"
        end_cell: 结束单元格地址，如"B3"
    
    Returns:
        单元格范围的内容
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法读取单元格范围，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path, data_only=True)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 读取单元格范围内容
        result = []
        for row in ws[f"{start_cell}:{end_cell}"]:
            row_values = []
            for cell in row:
                row_values.append(str(cell.value if cell.value is not None else ""))
            result.append(row_values)
        
        # 格式化输出
        formatted_result = ""
        for row in result:
            formatted_result += " | ".join(row) + "\n"
        
        return f"单元格范围 {start_cell}:{end_cell} 的内容:\n{formatted_result}"
    
    except Exception as e:
        return f"读取单元格范围时出错: {str(e)}"

@mcp.tool()
def write_cell(file_path: str, sheet_name: str, cell: str, value: str) -> str:
    """
    向Excel工作簿中指定单元格写入内容。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell: 单元格地址，如"A1", "B2"等
        value: 要写入的值
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法写入单元格，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 尝试转换为数字或日期（如果适用）
        try:
            # 尝试转换为整数
            int_value = int(value)
            ws[cell] = int_value
        except ValueError:
            try:
                # 尝试转换为浮点数
                float_value = float(value)
                ws[cell] = float_value
            except ValueError:
                # 作为文本保存
                ws[cell] = value
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功将值 '{value}' 写入到单元格 {sheet_name}!{cell}"
    
    except Exception as e:
        return f"写入单元格时出错: {str(e)}"

@mcp.tool()
def write_cell_range(file_path: str, sheet_name: str, start_cell: str, data: List[List[str]]) -> str:
    """
    向Excel工作簿中指定区域写入数据。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        start_cell: 起始单元格地址，如"A1"
        data: 要写入的数据，二维数组
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法写入数据区域，请先安装openpyxl库: pip install openpyxl"
    
    # 检查参数
    if not data:
        return "错误: 数据不能为空"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 解析起始单元格
        start_col_letter = ''.join(filter(str.isalpha, start_cell))
        start_row = int(''.join(filter(str.isdigit, start_cell)))
        
        # 写入数据
        for row_idx, row_data in enumerate(data):
            for col_idx, cell_value in enumerate(row_data):
                col_letter = get_column_letter(column_index_from_string(start_col_letter) + col_idx)
                cell_address = f"{col_letter}{start_row + row_idx}"
                
                # 尝试转换为数字（如果适用）
                try:
                    # 尝试转换为整数
                    int_value = int(cell_value)
                    ws[cell_address] = int_value
                except (ValueError, TypeError):
                    try:
                        # 尝试转换为浮点数
                        float_value = float(cell_value)
                        ws[cell_address] = float_value
                    except (ValueError, TypeError):
                        # 作为文本保存
                        ws[cell_address] = cell_value
        
        # 保存工作簿
        wb.save(file_path)
        
        # 计算结束单元格
        end_row = start_row + len(data) - 1
        end_col_letter = get_column_letter(column_index_from_string(start_col_letter) + len(data[0]) - 1)
        end_cell = f"{end_col_letter}{end_row}"
        
        return f"成功将数据写入区域 {sheet_name}!{start_cell}:{end_cell}"
    
    except Exception as e:
        return f"写入数据区域时出错: {str(e)}"

@mcp.tool()
def clear_cell(file_path: str, sheet_name: str, cell: str) -> str:
    """
    清除Excel工作簿中指定单元格的内容。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell: 单元格地址，如"A1", "B2"等
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法清除单元格，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 清除单元格内容
        ws[cell].value = None
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功清除单元格 {sheet_name}!{cell} 的内容"
    
    except Exception as e:
        return f"清除单元格时出错: {str(e)}"

@mcp.tool()
def format_cell(
    file_path: str,
    sheet_name: str,
    cell: str,
    font_name: str = None,
    font_size: int = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False,
    font_color: str = None,
    background_color: str = None,
    horizontal_alignment: str = None,
    vertical_alignment: str = None
) -> str:
    """
    设置Excel工作簿中指定单元格的格式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell: 单元格地址，如"A1", "B2"等，或单元格范围如"A1:B3"
        font_name: 字体名称
        font_size: 字体大小（点）
        bold: 是否加粗
        italic: 是否斜体
        underline: 是否下划线
        font_color: 字体颜色 (十六进制RGB格式，如"#FF0000"表示红色)
        background_color: 单元格背景色 (十六进制RGB格式)
        horizontal_alignment: 水平对齐方式 ("left", "center", "right", "justify")
        vertical_alignment: 垂直对齐方式 ("top", "center", "bottom")
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法设置单元格格式，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 定义对齐方式映射
    h_align_map = {
        "left": "left",
        "center": "center",
        "right": "right",
        "justify": "justify"
    }
    
    v_align_map = {
        "top": "top",
        "center": "center",
        "bottom": "bottom"
    }
    
    # 校验对齐方式参数
    if horizontal_alignment and horizontal_alignment.lower() not in h_align_map:
        return f"错误: 无效的水平对齐方式 '{horizontal_alignment}'，可选值为: left, center, right, justify"
    
    if vertical_alignment and vertical_alignment.lower() not in v_align_map:
        return f"错误: 无效的垂直对齐方式 '{vertical_alignment}'，可选值为: top, center, bottom"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 处理单元格范围
        cells = ws[cell] if ":" in cell else [ws[cell]]
        
        # 如果cells是单个单元格而不是列表，则转为列表
        if not isinstance(cells, list) and not isinstance(cells, tuple):
            cells = [cells]
        elif isinstance(cells, tuple) and len(cells) == 1:
            cells = [cells[0]]
        
        # 遍历所有单元格
        for cell_obj in cells:
            if isinstance(cell_obj, tuple):  # 处理元组行情况
                for c in cell_obj:
                    set_cell_format(c, font_name, font_size, bold, italic, underline, 
                                   font_color, background_color, 
                                   horizontal_alignment, vertical_alignment)
            else:
                set_cell_format(cell_obj, font_name, font_size, bold, italic, underline, 
                               font_color, background_color, 
                               horizontal_alignment, vertical_alignment)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功设置单元格 {sheet_name}!{cell} 的格式"
    
    except Exception as e:
        return f"设置单元格格式时出错: {str(e)}"

def set_cell_format(cell, font_name, font_size, bold, italic, underline, 
                   font_color, background_color, horizontal_alignment, vertical_alignment):
    """辅助函数：设置单元格格式"""
    # 创建字体对象
    if font_name or font_size or bold or italic or underline or font_color:
        font = Font(
            name=font_name,
            size=font_size,
            bold=bold,
            italic=italic,
            underline=underline if underline else None
        )
        
        # 设置字体颜色
        if font_color:
            try:
                # 解析十六进制颜色值
                if font_color.startswith("#"):
                    font_color = font_color[1:]
                font.color = Color(rgb=font_color)
            except ValueError:
                pass
        
        cell.font = font
    
    # 设置背景色
    if background_color:
        try:
            # 解析十六进制颜色值
            if background_color.startswith("#"):
                background_color = background_color[1:]
            fill = PatternFill(start_color=background_color, end_color=background_color, fill_type="solid")
            cell.fill = fill
        except ValueError:
            pass
    
    # 设置对齐方式
    if horizontal_alignment or vertical_alignment:
        alignment = Alignment()
        
        if horizontal_alignment:
            alignment.horizontal = horizontal_alignment.lower()
        
        if vertical_alignment:
            alignment.vertical = vertical_alignment.lower()
        
        cell.alignment = alignment

@mcp.tool()
def merge_cells(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """
    合并Excel工作簿中指定范围的单元格。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        start_cell: 起始单元格地址，如"A1"
        end_cell: 结束单元格地址，如"B3"
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法合并单元格，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 合并单元格
        ws.merge_cells(f"{start_cell}:{end_cell}")
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功合并单元格 {sheet_name}!{start_cell}:{end_cell}"
    
    except Exception as e:
        return f"合并单元格时出错: {str(e)}"

@mcp.tool()
def unmerge_cells(file_path: str, sheet_name: str, start_cell: str, end_cell: str) -> str:
    """
    拆分Excel工作簿中已合并的单元格。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        start_cell: 起始单元格地址，如"A1"
        end_cell: 结束单元格地址，如"B3"
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法拆分单元格，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 拆分单元格
        ws.unmerge_cells(f"{start_cell}:{end_cell}")
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功拆分单元格 {sheet_name}!{start_cell}:{end_cell}"
    
    except Exception as e:
        return f"拆分单元格时出错: {str(e)}"

@mcp.tool()
def insert_row(file_path: str, sheet_name: str, row_idx: int) -> str:
    """
    在Excel工作簿中插入空行。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        row_idx: 要插入行的位置（从1开始）
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法插入行，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 检查行索引是否有效
        if row_idx < 1:
            return "错误: 行索引必须大于或等于1"
        
        # 插入行
        ws.insert_rows(row_idx)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 中第 {row_idx} 行位置插入了空行"
    
    except Exception as e:
        return f"插入行时出错: {str(e)}"

@mcp.tool()
def insert_column(file_path: str, sheet_name: str, col_idx: int) -> str:
    """
    在Excel工作簿中插入空列。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        col_idx: 要插入列的位置（从1开始）
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法插入列，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 检查列索引是否有效
        if col_idx < 1:
            return "错误: 列索引必须大于或等于1"
        
        # 插入列
        ws.insert_cols(col_idx)
        
        # 保存工作簿
        wb.save(file_path)
        
        # 获取列字母
        col_letter = get_column_letter(col_idx)
        
        return f"成功在工作表 {sheet_name} 中第 {col_letter} 列位置插入了空列"
    
    except Exception as e:
        return f"插入列时出错: {str(e)}"

@mcp.tool()
def delete_row(file_path: str, sheet_name: str, row_idx: int) -> str:
    """
    从Excel工作簿中删除指定行。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        row_idx: 要删除的行位置（从1开始）
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法删除行，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 检查行索引是否有效
        if row_idx < 1:
            return "错误: 行索引必须大于或等于1"
        
        # 删除行
        ws.delete_rows(row_idx)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功从工作表 {sheet_name} 中删除了第 {row_idx} 行"
    
    except Exception as e:
        return f"删除行时出错: {str(e)}"

@mcp.tool()
def delete_column(file_path: str, sheet_name: str, col_idx: int) -> str:
    """
    从Excel工作簿中删除指定列。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        col_idx: 要删除的列位置（从1开始）
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法删除列，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 检查列索引是否有效
        if col_idx < 1:
            return "错误: 列索引必须大于或等于1"
        
        # 获取列字母（用于返回消息）
        col_letter = get_column_letter(col_idx)
        
        # 删除列
        ws.delete_cols(col_idx)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功从工作表 {sheet_name} 中删除了第 {col_letter} 列"
    
    except Exception as e:
        return f"删除列时出错: {str(e)}"

@mcp.tool()
def sort_data(file_path: str, sheet_name: str, range_to_sort: str, sort_column: int, 
              ascending: bool = True, has_header: bool = True) -> str:
    """
    对Excel工作簿中的数据区域进行排序。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        range_to_sort: 要排序的数据区域，如"A1:D10"
        sort_column: 排序依据的列索引（在选定区域内，从1开始）
        ascending: 是否升序排序，True为升序，False为降序
        has_header: 是否包含标题行，如果为True，则第一行作为标题不参与排序
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法排序数据，请先安装openpyxl库: pip install openpyxl"
    
    # 如果pandas未安装，则无法使用此功能
    if not pandas_installed:
        return "错误: 排序功能需要pandas库支持，请使用以下命令安装: pip install pandas numpy"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 检查排序列索引是否有效
    if sort_column < 1:
        return "错误: 排序列索引必须大于或等于1"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 获取排序区域的数据
        data = []
        for row in ws[range_to_sort]:
            data.append([cell.value for cell in row])
        
        # 使用pandas进行排序
        header_row = 0 if has_header else None
        df = pd.DataFrame(data)
        
        if has_header:
            # 保存标题行
            headers = df.iloc[0].tolist()
            # 排序剩余数据
            sorted_df = df.iloc[1:].sort_values(by=sort_column-1, ascending=ascending)
            # 重新添加标题行
            sorted_df = pd.concat([pd.DataFrame([headers], columns=df.columns), sorted_df])
        else:
            # 直接排序所有数据
            sorted_df = df.sort_values(by=sort_column-1, ascending=ascending)
        
        # 将排序后的数据写回工作表
        sorted_data = sorted_df.values.tolist()
        
        # 确定区域的行列范围
        start_cell, end_cell = range_to_sort.split(":")
        start_col_letter = ''.join(filter(str.isalpha, start_cell))
        start_row = int(''.join(filter(str.isdigit, start_cell)))
        
        # 写回数据
        for row_idx, row_data in enumerate(sorted_data):
            for col_idx, value in enumerate(row_data):
                col_letter = get_column_letter(column_index_from_string(start_col_letter) + col_idx)
                cell_address = f"{col_letter}{start_row + row_idx}"
                ws[cell_address] = value
        
        # 保存工作簿
        wb.save(file_path)
        
        # 获取排序依据列的字母
        sort_col_letter = get_column_letter(column_index_from_string(start_col_letter) + sort_column - 1)
        
        return f"成功对工作表 {sheet_name} 中区域 {range_to_sort} 的数据按照第 {sort_col_letter} 列进行{'升序' if ascending else '降序'}排序"
    
    except Exception as e:
        return f"排序数据时出错: {str(e)}"

@mcp.tool()
def apply_formula(file_path: str, sheet_name: str, cell: str, formula: str) -> str:
    """
    在Excel工作簿中指定单元格添加公式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell: 单元格地址，如"A1", "B2"等
        formula: 要添加的公式，不需要包含等号前缀，如"SUM(A1:A10)"
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法添加公式，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 确保公式以等号开头
        if not formula.startswith("="):
            formula = "=" + formula
        
        # 添加公式
        ws[cell] = formula
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在单元格 {sheet_name}!{cell} 中添加公式: {formula}"
    
    except Exception as e:
        return f"添加公式时出错: {str(e)}"

@mcp.tool()
def batch_fill(file_path: str, sheet_name: str, range_to_fill: str, value: str, 
               is_formula: bool = False) -> str:
    """
    在Excel工作簿中批量填充单元格。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        range_to_fill: 要填充的单元格范围，如"A1:D10"
        value: 要填充的值或公式
        is_formula: 是否为公式，如果为True，则value为公式内容
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法批量填充单元格，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 如果是公式，确保以等号开头
        if is_formula and not value.startswith("="):
            value = "=" + value
        
        # 尝试转换为数字（如果适用）
        if not is_formula:
            try:
                # 尝试转换为整数
                value = int(value)
            except ValueError:
                try:
                    # 尝试转换为浮点数
                    value = float(value)
                except ValueError:
                    # 保持原样（字符串）
                    pass
        
        # 确定区域的行列范围
        start_cell, end_cell = range_to_fill.split(":")
        
        # 填充单元格
        cell_count = 0
        for row in ws[range_to_fill]:
            for cell in row:
                cell.value = value
                cell_count += 1
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 中填充了 {cell_count} 个单元格 ({range_to_fill})"
    
    except Exception as e:
        return f"批量填充单元格时出错: {str(e)}"

@mcp.tool()
def create_chart(
    file_path: str, 
    sheet_name: str, 
    data_range: str, 
    chart_type: str = "column",
    title: str = "",
    categories_range: str = None,
    position: str = "F1"
) -> str:
    """
    在Excel工作簿中创建图表。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        data_range: 数据区域，如"A1:D10"
        chart_type: 图表类型，可选值: "column" (柱状图), "line" (折线图), "pie" (饼图), "bar" (条形图)
        title: 图表标题
        categories_range: 类别标签区域，如"A1:A10" (可选)
        position: 图表放置的位置，如"F1"
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法创建图表，请先安装openpyxl库: pip install openpyxl"
    
    try:
        # 尝试导入图表相关模块
        from openpyxl.chart import BarChart, LineChart, PieChart, Reference
        from openpyxl.chart.series import DataPoint
    except ImportError:
        return "错误: 创建图表需要openpyxl的chart模块支持"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证图表类型
    valid_chart_types = ["column", "line", "pie", "bar"]
    if chart_type not in valid_chart_types:
        return f"错误: 无效的图表类型 '{chart_type}'，可选值为: {', '.join(valid_chart_types)}"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 解析数据区域
        try:
            start_cell, end_cell = data_range.split(":")
            start_col_letter = ''.join(filter(str.isalpha, start_cell))
            start_row = int(''.join(filter(str.isdigit, start_cell)))
            end_col_letter = ''.join(filter(str.isalpha, end_cell))
            end_row = int(''.join(filter(str.isdigit, end_cell)))
            
            start_col = column_index_from_string(start_col_letter)
            end_col = column_index_from_string(end_col_letter)
        except Exception:
            return f"错误: 无效的数据区域格式 '{data_range}'，应为如 'A1:D10' 的格式"
        
        # 创建数据引用
        data_ref = Reference(
            ws, 
            min_col=start_col, 
            min_row=start_row, 
            max_col=end_col, 
            max_row=end_row
        )
        
        # 创建类别引用（如果提供）
        if categories_range:
            try:
                cats_start, cats_end = categories_range.split(":")
                cats_col_letter = ''.join(filter(str.isalpha, cats_start))
                cats_start_row = int(''.join(filter(str.isdigit, cats_start)))
                cats_end_row = int(''.join(filter(str.isdigit, cats_end)))
                
                cats_col = column_index_from_string(cats_col_letter)
                
                categories = Reference(
                    ws,
                    min_col=cats_col,
                    min_row=cats_start_row,
                    max_row=cats_end_row
                )
            except Exception:
                return f"错误: 无效的类别区域格式 '{categories_range}'，应为如 'A1:A10' 的格式"
        else:
            categories = None
        
        # 创建适当类型的图表
        if chart_type == "column":
            chart = BarChart()
            chart.type = "col"
        elif chart_type == "bar":
            chart = BarChart()
            chart.type = "bar"
        elif chart_type == "line":
            chart = LineChart()
        elif chart_type == "pie":
            chart = PieChart()
        
        # 设置图表标题
        chart.title = title
        
        # 添加数据
        if chart_type == "pie":
            # 饼图通常只使用一个系列的数据
            chart.add_data(data_ref, titles_from_data=True)
            if categories:
                chart.set_categories(categories)
        else:
            # 柱状图、条形图和折线图可以使用多个系列
            chart.add_data(data_ref, titles_from_data=True)
            if categories:
                chart.set_categories(categories)
        
        # 添加图表到工作表
        ws.add_chart(chart, position)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 中添加了 {chart_type} 类型的图表"
    
    except Exception as e:
        return f"创建图表时出错: {str(e)}"

@mcp.tool()
def apply_filter(file_path: str, sheet_name: str, filter_range: str) -> str:
    """
    在Excel工作簿中应用数据筛选功能。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        filter_range: 要应用筛选的数据区域，如"A1:D10"
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法应用筛选，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 导入AutoFilter模块
        from openpyxl.worksheet.filters import AutoFilter
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 应用自动筛选
        ws.auto_filter.ref = filter_range
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 的区域 {filter_range} 应用了数据筛选功能"
    
    except Exception as e:
        return f"应用数据筛选时出错: {str(e)}"

@mcp.tool()
def filter_data(
    file_path: str, 
    sheet_name: str, 
    column_letter: str, 
    filter_criteria: str,
    filter_value: str
) -> str:
    """
    在已应用筛选的Excel工作簿中设置筛选条件。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        column_letter: 要筛选的列字母，如"A"、"B"等
        filter_criteria: 筛选条件，可选值: "equals", "not_equals", "greater_than", "less_than", "contains", "not_contains", "begins_with", "ends_with"
        filter_value: 筛选值
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法设置筛选条件，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证筛选条件
    valid_criteria = ["equals", "not_equals", "greater_than", "less_than", "contains", "not_contains", "begins_with", "ends_with"]
    if filter_criteria not in valid_criteria:
        return f"错误: 无效的筛选条件 '{filter_criteria}'，可选值为: {', '.join(valid_criteria)}"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 检查是否已应用自动筛选
        if not ws.auto_filter:
            return f"错误: 工作表 {sheet_name} 尚未应用自动筛选，请先使用apply_filter功能"
        
        # 获取列索引
        try:
            column_idx = column_index_from_string(column_letter)
        except:
            return f"错误: 无效的列字母 '{column_letter}'"
        
        # 注：由于openpyxl的限制，我们只能创建筛选定义，但不能实际执行筛选操作
        # 这需要Excel应用程序的交互才能完成，所以这里我们只能添加筛选定义
        
        # 创建筛选定义
        from openpyxl.worksheet.filters import CustomFilter, CustomFilters, FilterColumn
        
        # 根据筛选条件创建CustomFilter
        operator_map = {
            "equals": "equal",
            "not_equals": "notEqual",
            "greater_than": "greaterThan",
            "less_than": "lessThan",
            "contains": "contains",
            "not_contains": "notContains",
            "begins_with": "beginsWith",
            "ends_with": "endsWith"
        }
        
        custom_filter = CustomFilter(operator=operator_map[filter_criteria], val=filter_value)
        
        # 创建CustomFilters，可以添加多个条件（这里只用一个）
        custom_filters = CustomFilters(customFilter=[custom_filter])
        
        # 创建FilterColumn
        filter_column = FilterColumn(colId=column_idx-1, customFilters=custom_filters)
        
        # 将FilterColumn添加到auto_filter
        ws.auto_filter.filterColumn.append(filter_column)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"已在工作表 {sheet_name} 的列 {column_letter} 上设置筛选条件: {filter_criteria} {filter_value}。请注意，实际筛选操作需要在Excel应用中查看。"
    
    except Exception as e:
        return f"设置筛选条件时出错: {str(e)}"

@mcp.tool()
def clear_filter(file_path: str, sheet_name: str) -> str:
    """
    清除Excel工作簿中的数据筛选。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法清除筛选，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 检查是否已应用自动筛选
        if not ws.auto_filter:
            return f"工作表 {sheet_name} 没有应用数据筛选，无需清除"
        
        # 清除筛选
        ws.auto_filter.ref = None
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功清除了工作表 {sheet_name} 中的数据筛选"
    
    except Exception as e:
        return f"清除数据筛选时出错: {str(e)}"

@mcp.tool()
def create_pivot_table(
    file_path: str, 
    sheet_name: str, 
    data_range: str, 
    target_sheet: str,
    target_cell: str = "A1",
    rows: list = None, 
    columns: list = None, 
    values: list = None, 
    filters: list = None
) -> str:
    """
    在Excel工作簿中创建数据透视表。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 源数据工作表名称
        data_range: 数据区域，如"A1:D10"
        target_sheet: 放置数据透视表的目标工作表名称
        target_cell: 数据透视表放置的起始单元格，默认为"A1"
        rows: 行标签的字段列表，如["姓名", "部门"]
        columns: 列标签的字段列表，如["年份", "月份"]
        values: 值字段列表，如[{"字段": "销售额", "函数": "SUM"}, {"字段": "成本", "函数": "AVERAGE"}]
        filters: 筛选字段列表，如["区域", "产品"]
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法创建数据透视表，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 初始化默认参数
    if rows is None:
        rows = []
    if columns is None:
        columns = []
    if values is None:
        values = []
    if filters is None:
        filters = []
    
    try:
        # 导入必要的模块
        from openpyxl.pivot.table import PivotTable, PivotFields, PivotField, Reference
        from openpyxl.pivot.cache import PivotCache
        from openpyxl.utils import get_column_letter
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查源工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 源工作表 '{sheet_name}' 不存在"
        
        # 检查目标工作表名称
        if target_sheet not in wb.sheetnames:
            # 创建新的目标工作表
            wb.create_sheet(title=target_sheet)
        
        # 获取源工作表和目标工作表
        source_ws = wb[sheet_name]
        target_ws = wb[target_sheet]
        
        # 解析数据范围
        try:
            from openpyxl.utils.cell import range_boundaries
            min_col, min_row, max_col, max_row = range_boundaries(data_range)
        except:
            return f"错误: 无效的数据区域格式 '{data_range}'"
        
        # 创建数据引用
        data_ref = Reference(source_ws, min_col=min_col, min_row=min_row, 
                             max_col=max_col, max_row=max_row)
        
        # 解析目标单元格
        try:
            from openpyxl.utils.cell import coordinate_from_string
            target_coord = coordinate_from_string(target_cell)
            target_row = target_coord[1]
            target_col = column_index_from_string(target_coord[0])
        except:
            return f"错误: 无效的目标单元格格式 '{target_cell}'"
        
        # 创建透视缓存
        # 注意：目前仅能使用openpyxl的基本功能设置透视表结构
        # 真正的数据处理需要在Excel应用中进行
        pivot_cache = PivotCache(cacheSource=data_ref, cacheId=1)
        
        # 创建透视表
        pivot_table = PivotTable(name=f"PivotTable{len(wb.sheetnames)}", 
                                 cache=pivot_cache,
                                 ref=f"{target_cell}:J20",  # 初始范围，会自动调整
                                 location=target_cell)
        
        # 获取列字段名称（假设第一行是标题）
        headers = [cell.value for cell in source_ws[min_row]]
        
        # 设置行字段
        for row_field in rows:
            if row_field in headers:
                field_idx = headers.index(row_field)
                pivot_field = PivotField(name=row_field)
                pivot_field.axis = "axisRow"
                pivot_table.fields.append(pivot_field)
        
        # 设置列字段
        for col_field in columns:
            if col_field in headers:
                field_idx = headers.index(col_field)
                pivot_field = PivotField(name=col_field)
                pivot_field.axis = "axisCol"
                pivot_table.fields.append(pivot_field)
        
        # 设置值字段
        for value_item in values:
            if isinstance(value_item, dict) and "字段" in value_item:
                field_name = value_item["字段"]
                if field_name in headers:
                    pivot_field = PivotField(name=field_name)
                    pivot_field.dataField = True
                    
                    # 设置汇总函数，如果提供
                    if "函数" in value_item:
                        function = value_item["函数"]
                        valid_functions = ["SUM", "COUNT", "AVERAGE", "MAX", "MIN"]
                        if function in valid_functions:
                            pivot_field.subtotal = function
                    
                    pivot_table.data_fields.append(pivot_field)
        
        # 设置筛选字段
        for filter_field in filters:
            if filter_field in headers:
                field_idx = headers.index(filter_field)
                pivot_field = PivotField(name=filter_field)
                pivot_field.axis = "axisPage"
                pivot_table.fields.append(pivot_field)
        
        # 将透视表添加到目标工作表
        target_ws.add_pivot_table(pivot_table)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"""成功在工作表 {target_sheet} 创建了数据透视表。
请注意：由于技术限制，此API只能创建透视表结构，但不能填充实际数据。
要查看透视表数据，请在Excel应用程序中打开文件并刷新透视表。"""
    
    except Exception as e:
        return f"创建数据透视表时出错: {str(e)}"

@mcp.tool()
def update_pivot_table(
    file_path: str, 
    sheet_name: str,
    pivot_name: str,
    add_row: str = None,
    remove_row: str = None,
    add_column: str = None,
    remove_column: str = None,
    add_value: dict = None,
    remove_value: str = None,
    add_filter: str = None,
    remove_filter: str = None
) -> str:
    """
    更新Excel工作簿中的数据透视表配置。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 包含数据透视表的工作表名称
        pivot_name: 数据透视表名称
        add_row: 要添加的行字段名称
        remove_row: 要移除的行字段名称
        add_column: 要添加的列字段名称
        remove_column: 要移除的列字段名称
        add_value: 要添加的值字段，格式如{"字段": "销售额", "函数": "SUM"}
        remove_value: 要移除的值字段名称
        add_filter: 要添加的筛选字段名称
        remove_filter: 要移除的筛选字段名称
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法更新数据透视表，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 导入必要的模块
        from openpyxl.pivot.table import PivotField
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 查找指定的数据透视表
        pivot_table = None
        for pt in ws._pivots:
            if pt.name == pivot_name:
                pivot_table = pt
                break
        
        if pivot_table is None:
            return f"错误: 未在工作表 {sheet_name} 中找到名为 '{pivot_name}' 的数据透视表"
        
        # 获取所有字段名称
        field_names = [field.name for field in pivot_table.fields]
        
        # 添加行字段
        if add_row and add_row in field_names:
            row_field = PivotField(name=add_row)
            row_field.axis = "axisRow"
            pivot_table.fields.append(row_field)
        
        # 移除行字段
        if remove_row:
            for i, field in enumerate(pivot_table.fields):
                if field.name == remove_row and field.axis == "axisRow":
                    del pivot_table.fields[i]
                    break
        
        # 添加列字段
        if add_column and add_column in field_names:
            col_field = PivotField(name=add_column)
            col_field.axis = "axisCol"
            pivot_table.fields.append(col_field)
        
        # 移除列字段
        if remove_column:
            for i, field in enumerate(pivot_table.fields):
                if field.name == remove_column and field.axis == "axisCol":
                    del pivot_table.fields[i]
                    break
        
        # 添加值字段
        if add_value and isinstance(add_value, dict) and "字段" in add_value:
            field_name = add_value["字段"]
            if field_name in field_names:
                value_field = PivotField(name=field_name)
                value_field.dataField = True
                
                # 设置汇总函数，如果提供
                if "函数" in add_value:
                    function = add_value["函数"]
                    valid_functions = ["SUM", "COUNT", "AVERAGE", "MAX", "MIN"]
                    if function in valid_functions:
                        value_field.subtotal = function
                
                pivot_table.data_fields.append(value_field)
        
        # 移除值字段
        if remove_value:
            for i, field in enumerate(pivot_table.data_fields):
                if field.name == remove_value:
                    del pivot_table.data_fields[i]
                    break
        
        # 添加筛选字段
        if add_filter and add_filter in field_names:
            filter_field = PivotField(name=add_filter)
            filter_field.axis = "axisPage"
            pivot_table.fields.append(filter_field)
        
        # 移除筛选字段
        if remove_filter:
            for i, field in enumerate(pivot_table.fields):
                if field.name == remove_filter and field.axis == "axisPage":
                    del pivot_table.fields[i]
                    break
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"""成功更新了工作表 {sheet_name} 中的数据透视表 '{pivot_name}'。
请注意：由于技术限制，此API只能更新透视表结构，但不能更新实际数据。
要查看更新后的透视表数据，请在Excel应用程序中打开文件并刷新透视表。"""
    
    except Exception as e:
        return f"更新数据透视表时出错: {str(e)}"

@mcp.tool()
def set_data_validation(
    file_path: str, 
    sheet_name: str, 
    cell_range: str, 
    validation_type: str,
    operator: str = None,
    formula1: str = None,
    formula2: str = None,
    allow_blank: bool = True,
    error_title: str = None,
    error_message: str = None,
    prompt_title: str = None,
    prompt_message: str = None
) -> str:
    """
    在Excel工作簿中设置数据有效性规则。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要应用数据有效性的单元格区域，如"A1:B10"
        validation_type: 验证类型，可选值：
            - "whole": 整数
            - "decimal": 小数
            - "list": 列表
            - "date": 日期
            - "time": 时间
            - "textLength": 文本长度
            - "custom": 自定义公式
        operator: 条件运算符，可选值：
            - "between": 介于...之间
            - "notBetween": 不介于...之间
            - "equal": 等于
            - "notEqual": 不等于
            - "greaterThan": 大于
            - "lessThan": 小于
            - "greaterThanOrEqual": 大于等于
            - "lessThanOrEqual": 小于等于
        formula1: 条件值1或列表源（对于列表类型）
        formula2: 条件值2（仅用于"between"和"notBetween"运算符）
        allow_blank: 是否允许空值
        error_title: 错误提示标题
        error_message: 错误提示消息
        prompt_title: 输入提示标题
        prompt_message: 输入提示消息
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法设置数据有效性，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证验证类型
    valid_types = ["whole", "decimal", "list", "date", "time", "textLength", "custom"]
    if validation_type not in valid_types:
        return f"错误: 无效的验证类型 '{validation_type}'，可选值为: {', '.join(valid_types)}"
    
    # 验证运算符
    if operator is not None:
        valid_operators = ["between", "notBetween", "equal", "notEqual", "greaterThan", 
                           "lessThan", "greaterThanOrEqual", "lessThanOrEqual"]
        if operator not in valid_operators:
            return f"错误: 无效的运算符 '{operator}'，可选值为: {', '.join(valid_operators)}"
    
    # 检查必要参数
    if formula1 is None and validation_type != "custom":
        return "错误: 未提供必要的条件值 (formula1)"
    
    try:
        # 导入必要的模块
        from openpyxl.worksheet.datavalidation import DataValidation
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 创建数据验证对象
        dv = DataValidation(
            type=validation_type,
            operator=operator,
            formula1=formula1,
            formula2=formula2,
            allow_blank=allow_blank,
            showErrorMessage=True if error_message else False,
            showInputMessage=True if prompt_message else False,
            errorTitle=error_title,
            error=error_message,
            promptTitle=prompt_title,
            prompt=prompt_message
        )
        
        # 添加单元格范围
        dv.add(cell_range)
        
        # 将数据验证添加到工作表
        ws.add_data_validation(dv)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 的区域 {cell_range} 设置了数据有效性规则"
    
    except Exception as e:
        return f"设置数据有效性时出错: {str(e)}"

@mcp.tool()
def create_dropdown_list(
    file_path: str, 
    sheet_name: str, 
    cell_range: str, 
    values: list,
    allow_blank: bool = True,
    error_title: str = None,
    error_message: str = None,
    prompt_title: str = None,
    prompt_message: str = None
) -> str:
    """
    在Excel工作簿中创建下拉列表。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要应用下拉列表的单元格区域，如"A1:B10"
        values: 下拉列表的选项值列表
        allow_blank: 是否允许空值
        error_title: 错误提示标题
        error_message: 错误提示消息
        prompt_title: 输入提示标题
        prompt_message: 输入提示消息
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法创建下拉列表，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 检查下拉列表值
    if not values or not isinstance(values, list):
        return "错误: 未提供有效的下拉列表选项"
    
    try:
        # 导入必要的模块
        from openpyxl.worksheet.datavalidation import DataValidation
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 将列表值转换为逗号分隔的字符串
        values_str = ','.join([f'"{val}"' for val in values])
        
        # 创建数据验证对象（下拉列表）
        dv = DataValidation(
            type="list",
            formula1=values_str,
            allow_blank=allow_blank,
            showErrorMessage=True if error_message else False,
            showInputMessage=True if prompt_message else False,
            errorTitle=error_title,
            error=error_message,
            promptTitle=prompt_title,
            prompt=prompt_message
        )
        
        # 添加单元格范围
        dv.add(cell_range)
        
        # 将数据验证添加到工作表
        ws.add_data_validation(dv)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 的区域 {cell_range} 创建了包含 {len(values)} 个选项的下拉列表"
    
    except Exception as e:
        return f"创建下拉列表时出错: {str(e)}"

@mcp.tool()
def clear_data_validation(
    file_path: str, 
    sheet_name: str, 
    cell_range: str
) -> str:
    """
    清除Excel工作簿中指定区域的数据有效性规则。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要清除数据有效性的单元格区域，如"A1:B10"
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法清除数据有效性，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 导入必要的模块
        from openpyxl.utils.cell import range_boundaries
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 解析单元格范围
        min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        
        # 遍历数据验证规则，移除符合条件的规则
        dvs_to_remove = []
        
        for dv in ws.data_validations.dataValidation:
            ranges_to_remove = []
            
            for dv_range in dv.sqref.ranges:
                dv_min_col, dv_min_row, dv_max_col, dv_max_row = range_boundaries(str(dv_range))
                
                # 检查是否与目标范围有重叠
                if (dv_min_col <= max_col and dv_max_col >= min_col and
                    dv_min_row <= max_row and dv_max_row >= min_row):
                    ranges_to_remove.append(dv_range)
            
            # 从验证规则中移除重叠的范围
            for range_to_remove in ranges_to_remove:
                dv.sqref.ranges.remove(range_to_remove)
            
            # 如果验证规则没有剩余范围，则标记为需要删除
            if not dv.sqref.ranges:
                dvs_to_remove.append(dv)
        
        # 移除需要删除的验证规则
        for dv in dvs_to_remove:
            ws.data_validations.dataValidation.remove(dv)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功清除了工作表 {sheet_name} 的区域 {cell_range} 中的数据有效性规则"
    
    except Exception as e:
        return f"清除数据有效性规则时出错: {str(e)}"

@mcp.tool()
def add_conditional_formatting(
    file_path: str, 
    sheet_name: str, 
    cell_range: str, 
    condition_type: str,
    format_type: str,
    condition_value: str = None,
    condition_value2: str = None,
    color: str = None,
    fill_type: str = "solid",
    text_color: str = None,
    bold: bool = False,
    italic: bool = False,
    underline: bool = False
) -> str:
    """
    在Excel工作簿中添加条件格式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要应用条件格式的单元格区域，如"A1:B10"
        condition_type: 条件类型，可选值：
            - "cellIs": 单元格值
            - "expression": 使用公式
            - "colorScale": 色阶
            - "dataBar": 数据条
            - "iconSet": 图标集
            - "top10": 前10项
            - "aboveAverage": 高于平均值
            - "duplicateValues": 重复值
            - "uniqueValues": 唯一值
            - "containsText": 包含文本
        format_type: 对于cellIs类型，指定条件运算符，可选值：
            - "lessThan": 小于
            - "lessThanOrEqual": 小于等于
            - "equal": 等于
            - "notEqual": 不等于
            - "greaterThanOrEqual": 大于等于
            - "greaterThan": 大于
            - "between": 介于
            - "notBetween": 不介于
            - "containsText": 包含
            - "notContainsText": 不包含
            - "beginsWith": 开始于
            - "endsWith": 结束于
        condition_value: 条件值
        condition_value2: 第二个条件值（用于"between"和"notBetween"）
        color: 当条件满足时应用的背景颜色 (十六进制RGB格式，如"#FF0000"表示红色)
        fill_type: 填充类型，默认为"solid"
        text_color: 当条件满足时应用的文本颜色 (十六进制RGB格式)
        bold: 是否加粗文本
        italic: 是否使用斜体
        underline: 是否添加下划线
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法添加条件格式，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证条件类型
    valid_conditions = ["cellIs", "expression", "colorScale", "dataBar", "iconSet", 
                         "top10", "aboveAverage", "duplicateValues", "uniqueValues", "containsText"]
    if condition_type not in valid_conditions:
        return f"错误: 无效的条件类型 '{condition_type}'，可选值为: {', '.join(valid_conditions)}"
    
    # 验证格式类型（对于cellIs类型）
    if condition_type == "cellIs":
        valid_formats = ["lessThan", "lessThanOrEqual", "equal", "notEqual", "greaterThanOrEqual", 
                          "greaterThan", "between", "notBetween", "containsText", "notContainsText", 
                          "beginsWith", "endsWith"]
        if format_type not in valid_formats:
            return f"错误: 无效的格式类型 '{format_type}'，可选值为: {', '.join(valid_formats)}"
    
    try:
        # 导入必要的模块
        from openpyxl.styles import PatternFill, Font, Color
        from openpyxl.styles.differential import DifferentialStyle
        from openpyxl.formatting.rule import Rule, CellIsRule, FormulaRule, ColorScaleRule, DataBarRule, IconSetRule
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 创建样式
        font = None
        fill = None
        
        if color or text_color or bold or italic or underline:
            # 设置字体样式
            font_args = {}
            if text_color:
                font_args["color"] = text_color
            if bold:
                font_args["bold"] = bold
            if italic:
                font_args["italic"] = italic
            if underline:
                font_args["underline"] = underline
            
            if font_args:
                font = Font(**font_args)
            
            # 设置填充样式
            if color:
                try:
                    fill = PatternFill(start_color=color, end_color=color, fill_type=fill_type)
                except:
                    fill = PatternFill(start_color=color, end_color=color, patternType=fill_type)
        
        # 创建差异样式
        diff_style = DifferentialStyle(font=font, fill=fill)
        
        # 根据条件类型创建规则
        rule = None
        
        if condition_type == "cellIs":
            if format_type in ["between", "notBetween"] and condition_value2 is None:
                return f"错误: 格式类型 '{format_type}' 需要提供第二个条件值 (condition_value2)"
            
            rule = CellIsRule(
                operator=format_type,
                formula=[condition_value] if condition_value else [],
                formula2=[condition_value2] if condition_value2 else [],
                stopIfTrue=False,
                dxf=diff_style
            )
        
        elif condition_type == "expression":
            if not condition_value:
                return "错误: 表达式条件类型需要提供条件公式 (condition_value)"
            
            rule = FormulaRule(
                formula=[condition_value],
                stopIfTrue=False,
                dxf=diff_style
            )
        
        elif condition_type == "colorScale":
            # 创建色阶规则
            rule = ColorScaleRule(
                start_type="min",
                start_value=None,
                start_color="FFFFFF",  # 白色
                end_type="max",
                end_value=None,
                end_color=color or "FF0000"  # 红色（如果未提供颜色）
            )
        
        elif condition_type == "dataBar":
            # 创建数据条规则
            rule = DataBarRule(
                start_type="min",
                start_value=None,
                end_type="max",
                end_value=None,
                color=color or "638EC6",  # 蓝色（如果未提供颜色）
                showValue=True,
                minLength=None,
                maxLength=None
            )
        
        elif condition_type == "iconSet":
            # 创建图标集规则
            rule = IconSetRule(
                icon_style=format_type or "3Arrows",  # 使用format_type作为图标样式
                type="percent",
                values=[0, 33, 67]
            )
        
        elif condition_type == "top10":
            # 创建前10项规则
            try:
                rank = int(condition_value or 10)
                percent = format_type == "percent"
                bottom = format_type == "bottom"
                
                rule = Rule(
                    type="top10",
                    rank=rank,
                    percent=percent,
                    bottom=bottom,
                    dxf=diff_style
                )
            except ValueError:
                return f"错误: top10条件类型需要提供有效的数字作为条件值，当前值: '{condition_value}'"
        
        elif condition_type == "aboveAverage":
            # 创建高于平均值规则
            above = format_type != "below"
            
            rule = Rule(
                type="aboveAverage",
                aboveAverage=above,
                dxf=diff_style
            )
        
        elif condition_type in ["duplicateValues", "uniqueValues"]:
            # 创建重复值或唯一值规则
            rule = Rule(
                type=condition_type,
                dxf=diff_style
            )
        
        elif condition_type == "containsText":
            # 创建包含文本规则
            if not condition_value:
                return "错误: containsText条件类型需要提供文本条件值 (condition_value)"
            
            operator = format_type or "containsText"
            valid_text_ops = ["containsText", "notContainsText", "beginsWith", "endsWith"]
            
            if operator not in valid_text_ops:
                return f"错误: 无效的文本操作符 '{operator}'，可选值为: {', '.join(valid_text_ops)}"
            
            formula = None
            if operator == "containsText":
                formula = f'NOT(ISERROR(SEARCH("{condition_value}",A1)))'
            elif operator == "notContainsText":
                formula = f'ISERROR(SEARCH("{condition_value}",A1))'
            elif operator == "beginsWith":
                formula = f'LEFT(A1,{len(condition_value)})="{condition_value}"'
            elif operator == "endsWith":
                formula = f'RIGHT(A1,{len(condition_value)})="{condition_value}"'
            
            rule = FormulaRule(
                formula=[formula],
                stopIfTrue=False,
                dxf=diff_style
            )
        
        # 添加条件格式规则到工作表
        if rule:
            ws.conditional_formatting.add(cell_range, rule)
            
            # 保存工作簿
            wb.save(file_path)
            
            return f"成功在工作表 {sheet_name} 的区域 {cell_range} 添加了条件格式规则"
        else:
            return "错误: 无法创建条件格式规则，请检查参数"
    
    except Exception as e:
        return f"添加条件格式时出错: {str(e)}"

@mcp.tool()
def add_data_bar(
    file_path: str, 
    sheet_name: str, 
    cell_range: str, 
    color: str = "#638EC6",
    min_type: str = "min",
    max_type: str = "max",
    min_value: str = None,
    max_value: str = None,
    show_value: bool = True
) -> str:
    """
    在Excel工作簿中添加数据条条件格式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要应用数据条的单元格区域，如"A1:B10"
        color: 数据条颜色 (十六进制RGB格式，如"#638EC6"表示蓝色)
        min_type: 最小值类型，可选值: "min", "num", "percent", "percentile", "formula"
        max_type: 最大值类型，可选值: "max", "num", "percent", "percentile", "formula"
        min_value: 最小值（仅当min_type不为"min"时使用）
        max_value: 最大值（仅当max_type不为"max"时使用）
        show_value: 是否显示单元格值
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法添加数据条，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证最小值和最大值类型
    valid_types = ["min", "max", "num", "percent", "percentile", "formula"]
    if min_type not in valid_types:
        return f"错误: 无效的最小值类型 '{min_type}'，可选值为: {', '.join(valid_types)}"
    if max_type not in valid_types:
        return f"错误: 无效的最大值类型 '{max_type}'，可选值为: {', '.join(valid_types)}"
    
    # 检查最小值和最大值
    if min_type != "min" and min_value is None:
        return f"错误: 最小值类型为 '{min_type}' 时必须提供最小值 (min_value)"
    if max_type != "max" and max_value is None:
        return f"错误: 最大值类型为 '{max_type}' 时必须提供最大值 (max_value)"
    
    try:
        # 导入必要的模块
        from openpyxl.formatting.rule import DataBarRule
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 创建数据条规则
        rule = DataBarRule(
            start_type=min_type,
            start_value=min_value,
            end_type=max_type,
            end_value=max_value,
            color=color.replace('#', ''),  # 移除颜色中的#前缀
            showValue=show_value,
            minLength=None,
            maxLength=None
        )
        
        # 添加条件格式规则到工作表
        ws.conditional_formatting.add(cell_range, rule)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 的区域 {cell_range} 添加了数据条条件格式"
    
    except Exception as e:
        return f"添加数据条时出错: {str(e)}"

@mcp.tool()
def add_color_scale(
    file_path: str, 
    sheet_name: str, 
    cell_range: str, 
    min_color: str = "#FFFFFF",
    mid_color: str = None,
    max_color: str = "#FF0000",
    min_type: str = "min",
    mid_type: str = None,
    max_type: str = "max",
    min_value: str = None,
    mid_value: str = None,
    max_value: str = None
) -> str:
    """
    在Excel工作簿中添加色阶条件格式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要应用色阶的单元格区域，如"A1:B10"
        min_color: 最小值颜色 (十六进制RGB格式，默认为白色)
        mid_color: 中间值颜色 (十六进制RGB格式，可选)
        max_color: 最大值颜色 (十六进制RGB格式，默认为红色)
        min_type: 最小值类型，可选值: "min", "num", "percent", "percentile", "formula"
        mid_type: 中间值类型，可选值同上，如果提供则使用三色色阶
        max_type: 最大值类型，可选值同上
        min_value: 最小值（仅当min_type不为"min"时使用）
        mid_value: 中间值（仅当mid_type不为None时使用）
        max_value: 最大值（仅当max_type不为"max"时使用）
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法添加色阶，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    # 验证类型
    valid_types = ["min", "max", "num", "percent", "percentile", "formula"]
    if min_type not in valid_types:
        return f"错误: 无效的最小值类型 '{min_type}'，可选值为: {', '.join(valid_types)}"
    if mid_type is not None and mid_type not in valid_types:
        return f"错误: 无效的中间值类型 '{mid_type}'，可选值为: {', '.join(valid_types)}"
    if max_type not in valid_types:
        return f"错误: 无效的最大值类型 '{max_type}'，可选值为: {', '.join(valid_types)}"
    
    # 检查值
    if min_type != "min" and min_value is None:
        return f"错误: 最小值类型为 '{min_type}' 时必须提供最小值 (min_value)"
    if mid_type is not None and mid_value is None:
        return f"错误: 中间值类型为 '{mid_type}' 时必须提供中间值 (mid_value)"
    if max_type != "max" and max_value is None:
        return f"错误: 最大值类型为 '{max_type}' 时必须提供最大值 (max_value)"
    
    try:
        # 导入必要的模块
        from openpyxl.formatting.rule import ColorScaleRule
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 移除颜色中的#前缀
        min_color = min_color.replace('#', '')
        max_color = max_color.replace('#', '')
        if mid_color:
            mid_color = mid_color.replace('#', '')
        
        # 创建色阶规则
        if mid_color and mid_type:
            # 三色色阶
            rule = ColorScaleRule(
                start_type=min_type,
                start_value=min_value,
                start_color=min_color,
                mid_type=mid_type,
                mid_value=mid_value,
                mid_color=mid_color,
                end_type=max_type,
                end_value=max_value,
                end_color=max_color
            )
        else:
            # 双色色阶
            rule = ColorScaleRule(
                start_type=min_type,
                start_value=min_value,
                start_color=min_color,
                end_type=max_type,
                end_value=max_value,
                end_color=max_color
            )
        
        # 添加条件格式规则到工作表
        ws.conditional_formatting.add(cell_range, rule)
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 的区域 {cell_range} 添加了色阶条件格式"
    
    except Exception as e:
        return f"添加色阶时出错: {str(e)}"

@mcp.tool()
def clear_conditional_formatting(
    file_path: str, 
    sheet_name: str, 
    cell_range: str = None
) -> str:
    """
    清除Excel工作簿中的条件格式。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要清除条件格式的单元格区域，如"A1:B10"，如果不提供则清除整个工作表的条件格式
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法清除条件格式，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 导入必要的模块
        from openpyxl.utils.cell import range_boundaries
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        if cell_range:
            # 解析单元格范围
            try:
                min_col, min_row, max_col, max_row = range_boundaries(cell_range)
            except:
                return f"错误: 无效的单元格区域格式 '{cell_range}'"
            
            # 获取要删除的条件格式规则
            cf_to_remove = []
            
            for cf in ws.conditional_formatting:
                for range_str in cf.sqref.ranges:
                    cf_min_col, cf_min_row, cf_max_col, cf_max_row = range_boundaries(str(range_str))
                    
                    # 检查是否有重叠
                    if (cf_min_col <= max_col and cf_max_col >= min_col and
                        cf_min_row <= max_row and cf_max_row >= min_row):
                        cf_to_remove.append(cf)
                        break
            
            # 移除标记的条件格式规则
            for cf in cf_to_remove:
                ws.conditional_formatting.remove(cf)
            
            msg = f"成功清除了工作表 {sheet_name} 的区域 {cell_range} 中的条件格式"
        else:
            # 清除整个工作表的条件格式
            ws.conditional_formatting.clear()
            msg = f"成功清除了工作表 {sheet_name} 中的所有条件格式"
        
        # 保存工作簿
        wb.save(file_path)
        
        return msg
    
    except Exception as e:
        return f"清除条件格式时出错: {str(e)}"

@mcp.tool()
def batch_replace(
    file_path: str, 
    sheet_name: str, 
    cell_range: str, 
    find_text: str,
    replace_text: str,
    match_case: bool = False,
    match_entire_cell: bool = False
) -> str:
    """
    在Excel工作簿中批量替换文本。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        cell_range: 要进行替换的单元格区域，如"A1:B10"
        find_text: 要查找的文本
        replace_text: 替换为的文本
        match_case: 是否区分大小写
        match_entire_cell: 是否匹配整个单元格内容
    
    Returns:
        操作结果信息
    """
    # 检查是否安装了必要的库
    if not openpyxl_installed:
        return "错误: 无法执行批量替换，请先安装openpyxl库: pip install openpyxl"
    
    # 检查是否提供了完整路径
    if not os.path.isabs(file_path):
        # 从环境变量获取基础路径
        base_path = os.environ.get('OFFICE_EDIT_PATH')
        if not base_path:
            base_path = os.path.join(os.path.expanduser('~'), '桌面')
        
        # 构建完整路径
        file_path = os.path.join(base_path, file_path)
    
    # 确保文件存在
    if not os.path.exists(file_path):
        return f"错误: 文件 {file_path} 不存在"
    
    try:
        # 导入必要的模块
        from openpyxl.utils.cell import range_boundaries
        
        # 打开Excel工作簿
        wb = load_workbook(file_path)
        
        # 检查工作表名称是否存在
        if sheet_name not in wb.sheetnames:
            return f"错误: 工作表 '{sheet_name}' 不存在"
        
        # 获取指定工作表
        ws = wb[sheet_name]
        
        # 解析单元格范围
        try:
            min_col, min_row, max_col, max_row = range_boundaries(cell_range)
        except:
            return f"错误: 无效的单元格区域格式 '{cell_range}'"
        
        # 计数器
        replace_count = 0
        
        # 遍历指定区域的单元格
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                cell = ws.cell(row=row, column=col)
                
                # 检查单元格值是否为字符串
                if cell.value and isinstance(cell.value, str):
                    original_value = cell.value
                    
                    if match_entire_cell:
                        # 匹配整个单元格内容
                        if match_case:
                            # 区分大小写
                            if original_value == find_text:
                                cell.value = replace_text
                                replace_count += 1
                        else:
                            # 不区分大小写
                            if original_value.lower() == find_text.lower():
                                cell.value = replace_text
                                replace_count += 1
                    else:
                        # 匹配部分内容
                        if match_case:
                            # 区分大小写
                            if find_text in original_value:
                                cell.value = original_value.replace(find_text, replace_text)
                                replace_count += 1
                        else:
                            # 不区分大小写 - 使用正则表达式进行不区分大小写的替换
                            import re
                            new_value = re.sub(re.escape(find_text), replace_text, original_value, flags=re.IGNORECASE)
                            if new_value != original_value:
                                cell.value = new_value
                                replace_count += 1
        
        # 保存工作簿
        wb.save(file_path)
        
        return f"成功在工作表 {sheet_name} 的区域 {cell_range} 中替换了 {replace_count} 处文本"
    
    except Exception as e:
        return f"批量替换文本时出错: {str(e)}"

@mcp.tool()
def apply_sum(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    range_to_sum: str
) -> str:
    """
    在Excel工作簿中应用SUM函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        range_to_sum: 要求和的单元格范围，如"A1:D1"
    
    Returns:
        操作结果信息
    """
    try:
        formula = f"SUM({range_to_sum})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用SUM函数时出错: {str(e)}"

@mcp.tool()
def apply_average(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    range_to_average: str
) -> str:
    """
    在Excel工作簿中应用AVERAGE函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        range_to_average: 要求平均值的单元格范围，如"A1:D1"
    
    Returns:
        操作结果信息
    """
    try:
        formula = f"AVERAGE({range_to_average})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用AVERAGE函数时出错: {str(e)}"

@mcp.tool()
def apply_count(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    range_to_count: str
) -> str:
    """
    在Excel工作簿中应用COUNT函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        range_to_count: 要计数的单元格范围，如"A1:D1"
    
    Returns:
        操作结果信息
    """
    try:
        formula = f"COUNT({range_to_count})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用COUNT函数时出错: {str(e)}"

@mcp.tool()
def apply_max(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    range_to_max: str
) -> str:
    """
    在Excel工作簿中应用MAX函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        range_to_max: 要求最大值的单元格范围，如"A1:D1"
    
    Returns:
        操作结果信息
    """
    try:
        formula = f"MAX({range_to_max})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用MAX函数时出错: {str(e)}"

@mcp.tool()
def apply_min(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    range_to_min: str
) -> str:
    """
    在Excel工作簿中应用MIN函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        range_to_min: 要求最小值的单元格范围，如"A1:D1"
    
    Returns:
        操作结果信息
    """
    try:
        formula = f"MIN({range_to_min})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用MIN函数时出错: {str(e)}"

@mcp.tool()
def apply_countif(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    range_to_count: str,
    criteria: str
) -> str:
    """
    在Excel工作簿中应用COUNTIF函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        range_to_count: 要计数的单元格范围，如"A1:D1"
        criteria: 计数条件，如">5"或"\"Text\""
    
    Returns:
        操作结果信息
    """
    try:
        # 如果条件是文本且未添加引号，则添加引号
        if not (criteria.startswith('"') and criteria.endswith('"')) and not any(op in criteria for op in ['>', '<', '=']):
            criteria = f'"{criteria}"'
        
        formula = f"COUNTIF({range_to_count},{criteria})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用COUNTIF函数时出错: {str(e)}"

@mcp.tool()
def apply_sumif(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    criteria_range: str,
    criteria: str,
    sum_range: str = None
) -> str:
    """
    在Excel工作簿中应用SUMIF函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        criteria_range: 条件范围，如"A1:A10"
        criteria: 求和条件，如">5"或"\"Text\""
        sum_range: 要求和的单元格范围，如"B1:B10"，如不提供则使用criteria_range
    
    Returns:
        操作结果信息
    """
    try:
        # 如果条件是文本且未添加引号，则添加引号
        if not (criteria.startswith('"') and criteria.endswith('"')) and not any(op in criteria for op in ['>', '<', '=']):
            criteria = f'"{criteria}"'
        
        if sum_range:
            formula = f"SUMIF({criteria_range},{criteria},{sum_range})"
        else:
            formula = f"SUMIF({criteria_range},{criteria})"
        
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用SUMIF函数时出错: {str(e)}"

@mcp.tool()
def apply_vlookup(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    lookup_value: str,
    table_array: str,
    col_index_num: int,
    range_lookup: bool = False
) -> str:
    """
    在Excel工作簿中应用VLOOKUP函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        lookup_value: 要查找的值，可以是单元格引用(如"A1")或具体值
        table_array: 查找表范围，如"A1:C10"
        col_index_num: 返回值的列号(从1开始)
        range_lookup: 是否允许近似匹配，True为允许，False为精确匹配
    
    Returns:
        操作结果信息
    """
    try:
        # 如果lookup_value不是单元格引用且是文本，需要添加引号
        if not re.match(r'^[A-Za-z]+[0-9]+$', lookup_value) and not lookup_value.isdigit():
            lookup_value = f'"{lookup_value}"'
        
        formula = f"VLOOKUP({lookup_value},{table_array},{col_index_num},{str(range_lookup).upper()})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用VLOOKUP函数时出错: {str(e)}"

@mcp.tool()
def apply_hlookup(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    lookup_value: str,
    table_array: str,
    row_index_num: int,
    range_lookup: bool = False
) -> str:
    """
    在Excel工作簿中应用HLOOKUP函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        lookup_value: 要查找的值，可以是单元格引用(如"A1")或具体值
        table_array: 查找表范围，如"A1:J3"
        row_index_num: 返回值的行号(从1开始)
        range_lookup: 是否允许近似匹配，True为允许，False为精确匹配
    
    Returns:
        操作结果信息
    """
    try:
        # 如果lookup_value不是单元格引用且是文本，需要添加引号
        if not re.match(r'^[A-Za-z]+[0-9]+$', lookup_value) and not lookup_value.isdigit():
            lookup_value = f'"{lookup_value}"'
        
        formula = f"HLOOKUP({lookup_value},{table_array},{row_index_num},{str(range_lookup).upper()})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用HLOOKUP函数时出错: {str(e)}"

@mcp.tool()
def apply_if(
    file_path: str, 
    sheet_name: str, 
    target_cell: str, 
    logical_test: str,
    value_if_true: str,
    value_if_false: str
) -> str:
    """
    在Excel工作簿中应用IF函数。
    
    Args:
        file_path: Excel工作簿的完整路径或相对于输出目录的路径
        sheet_name: 工作表名称
        target_cell: 放置公式的目标单元格，如"E1"
        logical_test: 逻辑测试条件，如"A1>B1"
        value_if_true: 条件为真时的返回值
        value_if_false: 条件为假时的返回值
    
    Returns:
        操作结果信息
    """
    try:
        # 处理文本值，添加引号
        for value in [value_if_true, value_if_false]:
            if not re.match(r'^[A-Za-z]+[0-9]+$', value) and not value.isdigit() and not value.startswith('"'):
                if value == value_if_true:
                    value_if_true = f'"{value}"'
                if value == value_if_false:
                    value_if_false = f'"{value}"'
        
        formula = f"IF({logical_test},{value_if_true},{value_if_false})"
        result = apply_formula(file_path, sheet_name, target_cell, formula)
        return result
    except Exception as e:
        return f"应用IF函数时出错: {str(e)}"

if __name__ == "__main__":
    # 运行MCP服务器
    print("启动Excel MCP服务器...")
    mcp.run()
