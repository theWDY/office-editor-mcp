"""
测试MCP服务器是否正常工作的脚本
"""

import os
import sys
import subprocess
import time

def main():
    print("测试OFFICE EDITOR MCP服务器...")
    
    # 检查MCP SDK是否已安装
    try:
        import mcp
        print(f"✓ MCP SDK已安装")
    except ImportError:
        print("✗ MCP SDK未安装，请运行 'pip install mcp'")
        sys.exit(1)
    
    # 检查服务器脚本是否存在
    server_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "word_server.py")
    if os.path.exists(server_path):
        print(f"✓ 服务器脚本文件存在: {server_path}")
    else:
        print(f"✗ 服务器脚本文件不存在: {server_path}")
        sys.exit(1)
    
    # 尝试启动服务器进程
    print("正在启动服务器...")
    try:
        # 启动服务器但不等待它完成（因为服务器会一直运行）
        server_process = subprocess.Popen(
            [sys.executable, server_path],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True
        )
        
        # 等待一会儿让服务器启动
        time.sleep(2)
        
        # 检查进程是否还在运行
        if server_process.poll() is None:
            print("✓ 服务器启动成功")
            
            # 创建一个测试文件
            desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')
            test_filename = "mcp_test.txt"
            test_file_path = os.path.join(desktop_path, test_filename)
            
            print(f"测试在桌面创建文件: {test_filename}")
            try:
                with open(test_file_path, 'w') as f:
                    pass
                
                if os.path.exists(test_file_path):
                    print(f"✓ 成功创建测试文件: {test_file_path}")
                    # 清理测试文件
                    os.remove(test_file_path)
                    print(f"✓ 已删除测试文件")
                else:
                    print(f"✗ 无法创建测试文件: {test_file_path}")
            except Exception as e:
                print(f"✗ 文件操作失败: {str(e)}")
        else:
            stdout, stderr = server_process.communicate()
            print(f"✗ 服务器启动失败")
            print(f"标准输出: {stdout}")
            print(f"错误输出: {stderr}")
    except Exception as e:
        print(f"✗ 启动服务器时出错: {str(e)}")
    finally:
        # 关闭服务器进程
        if 'server_process' in locals() and server_process.poll() is None:
            server_process.terminate()
            print("✓ 已停止服务器进程")
    
    print("\n测试总结:")
    print("如果上面显示所有检查都通过了，那么您的MCP服务器应该可以正常工作。")
    print("接下来您可以按照README.md中的说明在Cursor中配置此服务器。")

if __name__ == "__main__":
    main() 