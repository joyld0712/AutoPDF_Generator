import tkinter as tk
import os
import sys
import json
from gui import App

def ensure_files_exist():
    # 获取应用程序的基础路径
    if getattr(sys, 'frozen', False):
        # 如果是打包后的应用
        base_path = os.path.dirname(sys.executable)
    else:
        # 如果是开发环境
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    
    # 确保address.txt存在
    address_file_path = os.path.join(base_path, 'address.txt')
    if not os.path.exists(address_file_path):
        # 创建默认的address.txt
        default_data = {
            "公司1": "公司1地址信息...",
            "银行信息": {
                "银行1": {
                    "account_name": "SHANG HAI DUO SI DAI WANG LUO KE JI YOU XIAN GONG SI",
                    "account_number": "798277695",
                    "bank_code": "016",
                    "branch_code": "478",
                    "swift_code": "DHBKHKHH",
                    "bank_name": "DBS Bank (Hong Kong) Limited",
                    "bank_address": "18th Floor, The Center, 99 Queen's Road Central, Central",
                    "city": "Hong Kong SAR"
                }
            }
        }
        with open(address_file_path, 'w', encoding='utf-8') as f:
            json.dump(default_data, f, ensure_ascii=False, indent=2)
        print(f"创建了默认的address.txt文件: {address_file_path}")
    
    # 确保templates目录存在
    templates_dir = os.path.join(base_path, 'templates')
    if not os.path.exists(templates_dir):
        os.makedirs(templates_dir)
        print(f"创建了templates目录: {templates_dir}")
    
    # 确保output目录存在
    output_dir = os.path.join(base_path, 'output')
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"创建了output目录: {output_dir}")

if __name__ == "__main__":
    ensure_files_exist()
    root = tk.Tk()
    app = App(root)
    root.mainloop()