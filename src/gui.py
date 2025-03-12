import re
import os
import sys
import json
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import DateEntry
from datetime import datetime
from config import Config
from processor import fill_word_template, convert_to_pdf, generate_invoice_pdf, log_submission

class App:
    # 在gui.py文件中修改读取address.txt的部分
    
    def __init__(self, root):
        self.root = root
        self.root.title("AutoPDF Generator")
        self.root.geometry('1024x768')
        self.config = Config()
        
        # 修改读取address.txt的方式，使用绝对路径
        try:
            # 获取应用程序的基础路径
            if getattr(sys, 'frozen', False):
                # 如果是打包后的应用
                base_path = os.path.dirname(sys.executable)
            else:
                # 如果是开发环境
                base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            
            address_file_path = os.path.join(base_path, 'address.txt')
            print(f"尝试读取地址文件: {address_file_path}")
            
            if not os.path.exists(address_file_path):
                # 如果文件不存在，创建一个空的address.txt文件
                with open(address_file_path, 'w', encoding='utf-8') as f:
                    json.dump({"公司1": "公司1地址信息...", "银行信息": {}}, f, ensure_ascii=False, indent=2)
                print(f"创建了新的address.txt文件: {address_file_path}")
            
            with open(address_file_path, 'r', encoding='utf-8') as f:
                address_data = json.load(f)
                # 分离公司信息和银行信息
                self.company_info = {k: v for k, v in address_data.items() if k != "银行信息"}
                self.bank_info = address_data.get("银行信息", {})
        except Exception as e:
            messagebox.showerror("错误", f"读取公司信息文件失败: {e}")
            print(f"读取地址文件出错: {e}")
            # 使用默认值
            self.company_info = {"默认公司": "默认地址信息"}
            self.bank_info = {}
        
        # 主容器使用Frame并居中
        main_frame = tk.Frame(root, padx=30, pady=30)
        main_frame.pack(expand=True, fill='both')
        
        # 模板选择区域
        template_frame = tk.LabelFrame(main_frame, text=" 选择模板 ", font=('微软雅黑',13), padx=15, pady=15)
        template_frame.grid(row=0, column=0, sticky='ew', pady=10)
        self.template_var = tk.StringVar(value="General_Agreement_Template.docx")
        templates = [
            "General_Agreement_Template.docx",
            "Invoice_AD_Template.docx",
            "Invoice_Promo_Template.docx"
        ]
        for t in templates:
            rb = tk.Radiobutton(template_frame, text=t, variable=self.template_var, value=t, command=self.update_form)
            rb.grid(row=len(template_frame.winfo_children()), column=0, sticky="w", pady=2)

        # 最近生成的文件列表
        recent_frame = tk.LabelFrame(main_frame, text=" 最近生成的文件 ", font=('微软雅黑',13), padx=15, pady=15)
        recent_frame.grid(row=0, column=1, sticky='nsew', pady=15, padx=15)
        self.recent_listbox = tk.Listbox(recent_frame, height=5, font=('微软雅黑',11))
        self.recent_listbox.pack(fill='both', expand=True)
        self.update_recent_files()

        # 表单框架（动态更新）
        self.form_frame = tk.LabelFrame(main_frame, text=" 填写信息 ", font=('微软雅黑',13), padx=15, pady=15)
        self.form_frame.grid(row=1, column=0, sticky='nsew', pady=10)
        # 按钮区域
        btn_frame = tk.Frame(main_frame)
        btn_frame.grid(row=2, column=0, sticky='e', pady=15)
        # 统一控件样式
        style = {'font': ('微软雅黑',11), 'padx':10, 'pady':8}
        
        # 带图标的按钮
        generate_btn = tk.Button(btn_frame, text=" 生成PDF", command=self.generate_pdf, **style,
                               bg='#4CAF50', fg='white', activebackground='#45a049',
                               compound='left')
        generate_btn.grid(row=0, column=0, padx=5)
        
        clear_btn = tk.Button(btn_frame, text=" 清除数据", command=self.clear_form, **style,
                            bg='#f44336', fg='white', activebackground='#d32f2f',
                            compound='left')
        clear_btn.grid(row=0, column=1, padx=5)
        
        # 添加悬停效果
        generate_btn.bind('<Enter>', lambda e: generate_btn.config(bg='#45a049'))
        generate_btn.bind('<Leave>', lambda e: generate_btn.config(bg='#4CAF50'))
        clear_btn.bind('<Enter>', lambda e: clear_btn.config(bg='#d32f2f'))
        clear_btn.bind('<Leave>', lambda e: clear_btn.config(bg='#f44336'))
        # 配置网格权重
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        self.entries = {}
        self.update_form()
        # 按钮
        

    def update_form(self):
        # 清除现有表单
        for widget in self.form_frame.winfo_children():
            widget.destroy()
        self.entries.clear()

        # 每个模板的字段
        template_fields = {
            "General_Agreement_Template.docx": ["business_name", "agreement_date", "seller_name", "seller_email", "store_link"],
            "Invoice_AD_Template.docx": ["invoice_my_address","invoice_address","invoice_no", "invoice_date","table_data"],
            "Invoice_Promo_Template.docx": ["invoice_my_address","invoice_address","invoice_no", "invoice_date", "table_data"]
        }
        selected_template = self.template_var.get()
        print(f"更新表单，选择的模板: {selected_template}")
        print(f"模板字段: {template_fields[selected_template]}")

        # 获取最近的输入数据
        recent_data = self.config.get_recent_inputs(selected_template)
        
        # 如果是发票模板，添加公司选择和银行信息选择下拉框
        row_offset = 0
        if selected_template in ["Invoice_AD_Template.docx", "Invoice_Promo_Template.docx"]:
            # 公司选择
            company_label = tk.Label(self.form_frame, text="选择公司:", 
                                   font=('微软雅黑',12), fg='#444444')
            company_label.grid(row=0, column=0, sticky="w", pady=5, padx=(0,10))
            
            self.company_var = tk.StringVar()
            company_select = ttk.Combobox(self.form_frame, width=40,
                                        font=('微软雅黑',12),
                                        textvariable=self.company_var)
            company_select['values'] = list(self.company_info.keys())
            company_select.grid(row=0, column=1, sticky="ew", pady=5, padx=(0,5))
            company_select.bind('<<ComboboxSelected>>', self.on_company_select)
            
            # 银行信息选择
            bank_label = tk.Label(self.form_frame, text="选择银行信息:", 
                                font=('微软雅黑',12), fg='#444444')
            bank_label.grid(row=1, column=0, sticky="w", pady=5, padx=(0,10))
            
            self.bank_var = tk.StringVar()
            self.bank_select = ttk.Combobox(self.form_frame, width=40,
                                     font=('微软雅黑',12),
                                     textvariable=self.bank_var)
            # 初始时不显示任何银行信息，等待用户选择公司后再更新
            self.bank_select['values'] = []
            self.bank_select.grid(row=1, column=1, sticky="ew", pady=5, padx=(0,5))
            
            row_offset = 2
        # 为General模板添加公司选择下拉框
        elif selected_template == "General_Agreement_Template.docx":
            company_label = tk.Label(self.form_frame, text="选择公司:", 
                                   font=('微软雅黑',12), fg='#444444')
            company_label.grid(row=0, column=0, sticky="w", pady=5, padx=(0,10))
            
            self.business_name_var = tk.StringVar()
            business_select = ttk.Combobox(self.form_frame, width=40,
                                        font=('微软雅黑',12),
                                        textvariable=self.business_name_var)
            # 设置公司选项
            business_select['values'] = ["上海公司: Shanghai Dosdai Network Tech Co.", 
                                       "香港公司: Number Seven Trading Limited"]
            business_select.grid(row=0, column=1, sticky="ew", pady=5, padx=(0,5))
            
            # 将business_name添加到entries字典中
            self.entries["business_name"] = business_select
            
            row_offset = 1
            
            # 从template_fields中移除business_name，因为我们已经处理了
            template_fields["General_Agreement_Template.docx"] = template_fields["General_Agreement_Template.docx"][1:]

        # 创建输入字段
        for i, field in enumerate(template_fields[selected_template], start=1):
            label = tk.Label(self.form_frame, text=field.replace("_", " ").title() + ":", 
                           font=('微软雅黑',12), fg='#444444')
            label.grid(row=i+row_offset, column=0, sticky="w", pady=5, padx=(0,10))
            
            if "date" in field:
                entry = DateEntry(self.form_frame, width=24, 
                                background='#0078D4', foreground='white', borderwidth=2)
            elif field == "table_data":
                # 添加多行文本输入区域用于表格数据
                entry = tk.Text(self.form_frame, height=5, width=34,
                              font=('微软雅黑',11), wrap=tk.WORD,
                              highlightbackground='#E0E0E0', highlightthickness=1)
                entry.grid(row=i+row_offset, column=1, sticky="ew", pady=5, padx=(0,5))
                
                # 将说明标签移动到文本框下方
                hint_label = tk.Label(self.form_frame, text="广告数据格式：Description,ASIN,Perday,Day;提报数据格式:Description,ASIN,Product,Amuount", 
                                   font=('微软雅黑',9), fg='#666666')
                hint_label.grid(row=i+row_offset+1, column=1, sticky="w", pady=(0,5), padx=(0,5))
                
                # 确保将entry添加到entries字典中
                self.entries[field] = entry
                print(f"创建Text控件用于字段: {field}")
                
                # 增加row_offset以适应新增的行
                row_offset += 1
                continue  # 跳过下面的通用grid设置，因为我们已经手动设置了
            elif field == "asin_list":
                entry = tk.Text(self.form_frame, height=5, width=34,
                              font=('微软雅黑',11), wrap=tk.WORD,
                              highlightbackground='#E0E0E0', highlightthickness=1)
            else:
                if field in ['seller_name', 'seller_email', 'store_link']:
                    entry = ttk.Combobox(self.form_frame, width=32,
                                        font=('微软雅黑',11))
                    entry['values'] = self.config.get_common_values(field)
                else:
                    entry = tk.Entry(self.form_frame, width=34,
                                   font=('微软雅黑',11),
                                   highlightbackground='#E0E0E0', highlightthickness=1)
            
            entry.grid(row=i+row_offset, column=1, sticky="ew", pady=5, padx=(0,5))
            
            # 设置最近使用的值
            if field in recent_data:
                if isinstance(entry, tk.Text):
                    entry.insert("1.0", recent_data[field])
                else:
                    entry.insert(0, recent_data[field])
            
            # 配置网格权重
            self.form_frame.rowconfigure(i+row_offset, weight=1)
            self.form_frame.columnconfigure(1, weight=1)
            # 添加输入框动画效果
            if not isinstance(entry, (ttk.Combobox, DateEntry)):
                entry.bind('<FocusIn>', lambda e: e.widget.config(highlightbackground='#0078D4'))
                entry.bind('<FocusOut>', lambda e: e.widget.config(highlightbackground='#E0E0E0'))
            self.entries[field] = entry

    def on_company_select(self, event=None):
        # 当选择公司时，自动填充地址信息
        selected_company = self.company_var.get()
        if selected_company in self.company_info and 'invoice_my_address' in self.entries:
            self.entries['invoice_my_address'].delete(0, tk.END)
            self.entries['invoice_my_address'].insert(0, self.company_info[selected_company])
            
            # 更新银行信息下拉框，只显示与所选公司相关的银行
            if hasattr(self, 'bank_select'):
                # 根据公司名称筛选银行信息
                company_prefix = ""
                if selected_company == "上海公司":
                    company_prefix = "上海-"
                elif selected_company == "香港公司":
                    company_prefix = "香港-"
                
                # 筛选以该公司前缀开头的银行信息
                filtered_banks = [bank for bank in self.bank_info.keys() 
                                 if bank.startswith(company_prefix)]
                
                print(f"为公司 '{selected_company}' 筛选银行信息: {filtered_banks}")
                
                # 更新下拉框的值
                self.bank_select['values'] = filtered_banks
                
                # 如果有匹配的银行，默认选择第一个
                if filtered_banks:
                    self.bank_var.set(filtered_banks[0])
                else:
                    self.bank_var.set("")

    def clear_form(self):
        for entry in self.entries.values():
            if isinstance(entry, tk.Entry) or isinstance(entry, DateEntry):
                entry.delete(0, tk.END)
            elif isinstance(entry, tk.Text):
                entry.delete("1.0", tk.END)

    def get_value(self, entry):
        """从不同类型的控件中获取值"""
        if isinstance(entry, tk.Text):
            value = entry.get("1.0", "end-1c")
            print(f"从Text控件获取值: '{value}'")
            return value
        elif isinstance(entry, DateEntry):
            date_value = entry.get_date().strftime("%Y-%m-%d")
            print(f"从DateEntry控件获取值: {date_value}")
            return date_value
        elif isinstance(entry, ttk.Combobox):
            value = entry.get()
            print(f"从Combobox控件获取值: {value}")
            return value
        else:
            value = entry.get()
            
            # 自动添加https://前缀并拆分多个网址
            field_name = [k for k, v in self.entries.items() if v == entry][0]
            if field_name == 'store_link':
                urls = [url.strip() for url in value.split(',')]
                validated_urls = []
                for url in urls:
                    if not url.startswith(('http://', 'https://')):
                        url = f'https://{url}'
                    validated_urls.append(url)
                value = ', '.join(validated_urls)
            
            print(f"从Entry控件获取值: {value}")
            return value

    def generate_pdf(self):
        selected_template = self.template_var.get()
        template_path = f"templates/{selected_template}"
        
        print("\n===== 开始生成PDF =====")
        print(f"选择的模板: {selected_template}")
        
        # 获取所有表单数据
        data = {}
        data['template_type'] = selected_template
        print("所有表单字段:")
        for field, entry in self.entries.items():
            print(f"字段名: {field}, 控件类型: {type(entry).__name__}")
            data[field] = self.get_value(entry)
            print(f"字段 '{field}' 的值: {data[field]}")
        
        # 处理business_name字段，提取实际的公司名称
        if 'business_name' in data and data['business_name']:
            # 从"上海公司: Shanghai Dosdai Network Tech Co."格式中提取实际名称
            if ":" in data['business_name']:
                data['business_name'] = data['business_name'].split(": ")[1]
            print(f"处理后的business_name: {data['business_name']}")
        
        # 验证输入
        if not all(data.values()):
            messagebox.showerror("错误", "请填写所有字段。")
            print("错误: 有字段未填写")
            return
        
        # 新增格式验证
        email_pattern = r'^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$'
        url_pattern = r'^https?://'
        
        if 'seller_email' in data and not re.match(email_pattern, data['seller_email']):
            messagebox.showerror("错误", "邮箱格式不正确，请检查后重新输入！")
            return
            
        if 'store_link' in data:
            urls = data['store_link'].split(',')
            for i, url in enumerate(urls, 1):
                if not re.match(url_pattern, url.strip()):
                    messagebox.showerror("错误", f"第{i}个网址格式不正确，必须以http://或https://开头！\n错误网址：{url}")
                    return

        # 更新配置
        self.config.update_recent_inputs(selected_template, data)
        for field in ['seller_name', 'seller_email', 'store_link']:
            if field in data:
                self.config.update_common_values(field, data[field])

        # 根据模板类型选择不同的PDF生成方式
        date_str = datetime.now().strftime("%Y%m%d")
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_dir = os.path.join("output", date_str)
        os.makedirs(output_dir, exist_ok=True)
        base_name = os.path.basename(selected_template).replace(".docx", "").replace("_Template", "")
        file_prefix = f"{base_name}_{timestamp}"
        output_pdf_path = os.path.join(output_dir, f"{file_prefix}.pdf")

        if selected_template in ["Invoice_AD_Template.docx", "Invoice_Promo_Template.docx"]:
            # 添加银行信息
            selected_bank = self.bank_var.get()
            
            # 如果没有选择银行信息，但选择了公司，则尝试自动选择匹配的银行
            if (not selected_bank or selected_bank not in self.bank_info) and self.company_var.get():
                selected_company = self.company_var.get()
                company_prefix = ""
                if selected_company == "上海公司":
                    company_prefix = "上海-"
                elif selected_company == "香港公司":
                    company_prefix = "香港-"
                
                # 尝试找到匹配的银行
                for bank_name in self.bank_info.keys():
                    if bank_name.startswith(company_prefix):
                        selected_bank = bank_name
                        print(f"自动选择银行信息: {selected_bank}")
                        break
            
            if selected_bank and selected_bank in self.bank_info:
                bank_info = self.bank_info[selected_bank]
                print(f"使用选择的银行信息: {selected_bank}")
            else:
                # 使用默认银行信息
                bank_info = {
                    "account_name": "SHANG HAI DUO SI DAI WANG LUO KE JI YOU XIAN GONG SI",
                    "account_number": "798277695",
                    "bank_code": "016",
                    "branch_code": "478",
                    "swift_code": "DHBKHKHH",
                    "bank_name": "DBS Bank (Hong Kong) Limited",
                    "bank_address": "18th Floor, The Center, 99 Queen's Road Central, Central",
                    "city": "Hong Kong SAR"
                }
                print("使用默认银行信息")
            
            data['bank_info'] = bank_info
            # 直接生成PDF
            generate_invoice_pdf(data, output_pdf_path)
        else:
            # 使用Word模板生成PDF
            filled_doc_path = fill_word_template(template_path, data)
            output_pdf_path = filled_doc_path.replace("_filled.docx", ".pdf")
            convert_to_pdf(filled_doc_path, output_pdf_path)
        
        # 记录操作
        log_submission(data, output_pdf_path)
        self.config.add_recent_file(output_pdf_path)
        self.update_recent_files()
        
        # 显示成功消息
        messagebox.showinfo("成功", f"PDF生成于: {output_pdf_path}")

    def update_recent_files(self):
        self.recent_listbox.delete(0, tk.END)
        recent_files = self.config.get_recent_files()
        for file in recent_files:
            # 从文件名中提取公司名称
            filename = os.path.basename(file)
            # 如果文件名包含连字符，提取连字符前的部分作为公司名称
            if '-' in filename:
                company_name = filename.split('-')[0]
            else:
                # 如果没有连字符，使用文件名（不含扩展名）
                company_name = os.path.splitext(filename)[0]
            self.recent_listbox.insert(tk.END, company_name)