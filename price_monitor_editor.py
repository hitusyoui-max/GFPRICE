import requests
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
class DocumentManager:
    def __init__(self):
        # 文档结构：{name: {'content': str, 'external_path': str or None}}
        self.documents = {}
        self.current_doc = None
        self.load_documents()
    
    def load_documents(self):
        """从文件加载文档列表"""
        if os.path.exists('documents.json'):
            try:
                with open('documents.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 兼容旧版本数据格式
                    for name, value in data.items():
                        if isinstance(value, str):
                            self.documents[name] = {'content': value, 'external_path': None}
                        else:
                            self.documents[name] = value
            except Exception as e:
                print(f"加载 documents.json 失败：{e}")
                self.documents = {}
    
    def save_documents(self):
        """保存文档列表到文件"""
        try:
            with open('documents.json', 'w', encoding='utf-8') as f:
                json.dump(self.documents, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"保存 documents.json 失败：{e}")
    
    def create_document(self, name, content="", external_path=None):
        """创建新文档"""
        if name in self.documents:
            return False, "文档名称已存在"
        self.documents[name] = {'content': content, 'external_path': external_path}
        self.save_documents()
        return True, f"文档「{name}」创建成功"
    
    def delete_document(self, name):
        """删除文档"""
        if name in self.documents:
            del self.documents[name]
            self.save_documents()
            return True, f"文档「{name}」删除成功"
        return False, "文档不存在"
    
    def update_document(self, name, content):
        """更新文档内容"""
        if name not in self.documents:
            return False, "文档不存在"
        
        doc_info = self.documents[name]
        doc_info['content'] = content
        
        # 如果是外部文件，同步保存到原文件
        if doc_info['external_path']:
            try:
                # 确保文件路径存在
                import os
                external_path = doc_info['external_path']
                
                # 写入文件（不自动换行，每行一个换行符）
                with open(external_path, 'w', encoding='utf-8', newline='') as f:
                    f.write(content)
                
                print(f"✅ 外部文件已同步保存：{external_path}")
                print(f"   内容长度：{len(content)} 字符")
            except PermissionError as e:
                error_msg = f"保存外部文件失败：文件被占用或无写入权限\n{str(e)}"
                print(f"❌ {error_msg}")
                return False, error_msg
            except Exception as e:
                error_msg = f"保存外部文件失败：{str(e)}"
                print(f"❌ {error_msg}")
                return False, error_msg
        
        self.save_documents()
        return True, f"文档「{name}」已保存"
    
    def reload_from_external(self, name):
        """从外部文件重新加载内容"""
        if name not in self.documents:
            return False, "文档不存在"
        
        doc_info = self.documents[name]
        external_path = doc_info['external_path']
        
        if not external_path:
            return False, "该文档不是外部关联文件，无需刷新"
        
        try:
            with open(external_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 更新内存中的内容
            doc_info['content'] = content
            self.save_documents()
            return True, f"已从外部文件重新加载内容"
        except Exception as e:
            return False, f"重新加载失败：{str(e)}"
    

    def sync_wps_doc(self):
        """同步WPS在线文档"""
        from tkinter import simpledialog, messagebox
        url = simpledialog.askstring("WPS在线文档", "请输入WPS表格分享链接:")
        if not url:
            return
        
        try:
            # 解析WPS分享链接，获取导出链接
            export_url = self.parse_wps_url(url)
            if not export_url:
                messagebox.showerror("错误", "无法解析WPS链接，请检查链接格式")
                return
            
            # 拉取数据
            data = self.fetch_wps_data(export_url)
            if not data:
                messagebox.showerror("错误", "拉取数据失败，请检查链接权限")
                return
            
            # 清空现有表格，导入数据
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            added = 0
            for row in data[1:]:  # 跳过表头
                if len(row) >= 9:
                    # 格式: [状态, ID, 商品名称, 链接, 最低价, 最高价, 监控间隔, 破价目标, 预期, 两轮循环]
                    while len(row) < 10:
                        row.append("否")  # 默认两轮循环为否
                    if not row[9] or str(row[9]).strip() == "":
                        row[9] = "否"
                    self.tree.insert('', 'end', values=row)
                    added += 1
            
            # 保存
            self.save_current_to_doc()
            messagebox.showinfo("同步成功", f"✅ 从WPS文档导入完成\n共导入 {added} 条商品数据")
            self.update_status(f"从WPS导入 {added} 条商品")
            self.update_tree_colors()
        
        except Exception as e:
            messagebox.showerror("异常", f"同步失败: {str(e)}")
    
    def parse_wps_url(self, share_url):
        """解析WPS分享链接，获取CSV导出链接"""
        share_url = share_url.strip()
        # 支持多种WPS链接格式
        # 格式1: https://docs.wps.cn/docs/s/xxxxxxxxxxx
        # 格式2: https://kdocs.cn/l/xxxxxxxxxxx
        # 格式3: https://www.wps.cn/doc/xxxxxxxxxxx
        
        # 如果已经是导出链接直接返回
        if share_url.endswith('.csv') or share_url.endswith('.xlsx'):
            return share_url
        
        # 提取fileid
        import re
        # kdocs.cn/l/abc123 格式
        match = re.search(r'kdocs\.cn\/l\/([a-zA-Z0-9]+)', share_url)
        if match:
            file_id = match.group(1)
            # 导出链接格式 (需要WPS开放API，这里使用公开分享获取CSV格式)
            return f"https://www.kdocs.cn/api/v1/file/{file_id}/export?format=csv"
        
        # docs.wps.cn/docs/s/ 格式
        match = re.search(r'docs\.wps\.cn\/docs\/s\/([a-zA-Z0-9]+)', share_url)
        if match:
            file_id = match.group(1)
            return f"https://docs.wps.cn/api/v1/file/{file_id}/export?format=csv"
        
        # 无法解析，用户直接输入导出链接
        return share_url
    
    def fetch_wps_data(self, url):
        """从WPS拉取CSV数据"""
        import requests
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
        }
        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()
        # 解析CSV
        content = response.text
        lines = [line.strip() for line in content.split('\n') if line.strip()]
        data = []
        for line in lines:
            # 简单CSV解析（处理逗号分隔）
            # 如果需要处理引号，需要用csv模块
            row = []
            in_quote = False
            current = ''
            for c in line:
                if c == '"':
                    in_quote = not in_quote
                elif c == ',' and not in_quote:
                    row.append(current.strip())
                    current = ''
                else:
                    current += c
            row.append(current.strip())
            # 清理空格
            row = [cell.strip().strip('"') for cell in row]
            data.append(row)
        return data
    
    def push_to_wps(self):
        """将数据推送回WPS文档"""
        from tkinter import messagebox
        messagebox.showinfo("提示", "WPS在线文档需要编辑权限才能推送\n\n请在WPS中开启「可编辑」分享权限，当前版本支持拉取数据完成\n推送功能需要配置WPS开放API，可后续升级")

    def rename_document(self, old_name, new_name):
        """重命名文档"""
        if old_name not in self.documents:
            return False, "原文档不存在"
        if new_name in self.documents:
            return False, "新名称已存在"
        doc_info = self.documents.pop(old_name)
        self.documents[new_name] = doc_info
        self.save_documents()
        return True, f"重命名成功：{old_name} → {new_name}"
    
    def get_document(self, name):
        """获取文档内容"""
        doc_info = self.documents.get(name)
        if doc_info:
            return doc_info['content']
        return None
    
    def get_external_path(self, name):
        """获取文档关联的外部文件路径"""
        doc_info = self.documents.get(name)
        if doc_info:
            return doc_info['external_path']
        return None
    
    def is_external_doc(self, name):
        """判断是否为外部文档"""
        doc_info = self.documents.get(name)
        if doc_info:
            return doc_info['external_path'] is not None
        return False
    
    def get_all_names(self):
        """获取所有文档名称"""
        return list(self.documents.keys())
class LinkManagerDialog:
    """链接管理器对话框"""
    def __init__(self, parent, initial_links=""):
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("竞品链接管理 🔗")
        self.dialog.geometry("500x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.resizable(False, False)
        
        # 居中
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx()+150, parent.winfo_rooty()+100))
        
        self.result = None
        
        # 初始化链接列表
        self.links = []
        if initial_links:
            # 按英文逗号分割，并去除空白
            self.links = [link.strip() for link in initial_links.split(',') if link.strip()]
        
        self.create_widgets()
    
    def create_widgets(self):
        """创建界面元素"""
        main_frame = ttk.Frame(self.dialog, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 标题
        ttk.Label(main_frame, text="竞品链接管理", font=('Microsoft YaHei', 12, 'bold')).grid(
            row=0, column=0, columnspan=3, pady=(0, 15), sticky=tk.W)
        
        # 链接列表框架
        list_frame = ttk.LabelFrame(main_frame, text="当前链接列表", padding=10)
        list_frame.grid(row=1, column=0, columnspan=3, sticky=tk.NSEW, pady=(0, 15))
        list_frame.grid_columnconfigure(0, weight=1)
        
        # 链接列表框
        self.link_listbox = tk.Listbox(list_frame, height=8, font=('Microsoft YaHei', 10))
        self.link_listbox.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.link_listbox.yview)
        self.link_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 初始化列表
        self.update_listbox()
        
        # 链接输入区域
        ttk.Label(main_frame, text="添加新链接:", font=('Microsoft YaHei', 10)).grid(
            row=2, column=0, sticky=tk.W, pady=(0, 5))
        
        self.link_entry = ttk.Entry(main_frame, width=40, font=('Microsoft YaHei', 10))
        self.link_entry.grid(row=3, column=0, sticky=tk.W+tk.E, pady=(0, 10))
        self.link_entry.bind('<Return>', lambda e: self.add_link())
        
        # 添加按钮
        add_btn = ttk.Button(main_frame, text="➕ 添加", command=self.add_link, width=10)
        add_btn.grid(row=3, column=1, padx=(10, 0), pady=(0, 10))
        
        # 删除按钮
        del_btn = ttk.Button(main_frame, text="🗑️ 删除选中", command=self.delete_selected, width=12)
        del_btn.grid(row=3, column=2, padx=(10, 0), pady=(0, 10))
        
        # 按钮区域
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=4, column=0, columnspan=3, pady=15)
        
        ttk.Button(btn_frame, text="✅ 确定", command=self.confirm, width=12).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ 取消", command=self.cancel, width=12).pack(side=tk.LEFT, padx=5)
        
        # 提示文本
        ttk.Label(main_frame, text="提示：链接会自动用英文逗号 (,) 分隔保存", 
                 foreground='gray', font=('Microsoft YaHei', 9)).grid(
                 row=5, column=0, columnspan=3, pady=(10, 0))
        
        # 配置网格权重
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)
    
    def update_listbox(self):
        """更新链接列表显示"""
        self.link_listbox.delete(0, tk.END)
        for i, link in enumerate(self.links, 1):
            # 显示前 50 个字符，超出显示...
            display_text = link[:50] + "..." if len(link) > 50 else link
            self.link_listbox.insert(tk.END, f"{i}. {display_text}")
    
    def add_link(self):
        """添加链接"""
        link = self.link_entry.get().strip()
        if not link:
            messagebox.showwarning("提示", "请输入链接地址")
            return
        
        # 检查是否已存在
        if link in self.links:
            messagebox.showinfo("提示", "该链接已存在")
            return
        
        self.links.append(link)
        self.update_listbox()
        self.link_entry.delete(0, tk.END)
        self.link_entry.focus()
    
    def delete_selected(self):
        """删除选中的链接"""
        selection = self.link_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的链接")
            return
        
        index = selection[0]
        if 0 <= index < len(self.links):
            del self.links[index]
            self.update_listbox()
    
    def confirm(self):
        """确认并返回结果"""
        if not self.links:
            self.result = ""
        else:
            self.result = ", ".join(self.links)  # 用英文逗号 + 空格分隔
        self.dialog.destroy()
    
    def cancel(self):
        """取消"""
        self.result = None
        self.dialog.destroy()
    
    def get_result(self):
        """获取结果"""
        return self.result
class PriceMonitorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("闲鱼竞价系统商品库 📊")
        self.root.geometry("1550x700")  # 进一步加宽窗口
        
        self.doc_manager = DocumentManager()
        self.current_row = None
        
        # 创建界面布局
        self.create_widgets()
        
        # 更新文档列表
        self.update_doc_list()
    
    def create_widgets(self):
        """创建界面元素"""
        # 分割窗口：左侧文档列表，右侧监控列表
        self.paned_window = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.paned_window.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # 左侧面板 - 文档管理
        self.left_frame = ttk.Frame(self.paned_window, width=220)
        self.paned_window.add(self.left_frame, weight=1)
        
        # 右侧面板 - 商品监控列表
        self.right_frame = ttk.Frame(self.paned_window)
        self.paned_window.add(self.right_frame, weight=6)
        
        # ============== 左侧 ==============
        # 左侧标题
        ttk.Label(self.left_frame, text="📚 监控文档列表", font=('Microsoft YaHei', 12, 'bold')).pack(pady=10)
        
        # 文档列表框
        self.doc_listbox = tk.Listbox(self.left_frame, font=('Microsoft YaHei', 10), height=15)
        self.doc_listbox.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.doc_listbox.bind('<<ListboxSelect>>', self.on_doc_select)
        
        # 左侧按钮区域
        self.left_buttons = ttk.Frame(self.left_frame)
        self.left_buttons.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(self.left_buttons, text="➕ 新建文档", command=self.new_document).pack(fill=tk.X, pady=2)
        ttk.Button(self.left_buttons, text="📝 重命名", command=self.rename_document).pack(fill=tk.X, pady=2)
        ttk.Button(self.left_buttons, text="🗑️ 删除文档", command=self.delete_document).pack(fill=tk.X, pady=2)
        ttk.Separator(self.left_buttons, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        ttk.Button(self.left_buttons, text="📁 关联外部 txt", command=self.open_external_txt).pack(fill=tk.X, pady=2)
        ttk.Button(self.left_buttons, text="🔄 刷新内容", command=self.refresh_current_doc).pack(fill=tk.X, pady=2)
        ttk.Button(self.left_buttons, text="💾 保存全部", command=self.save_current_doc).pack(fill=tk.X, pady=2)
        
        # 格式说明
        ttk.Separator(self.left_buttons, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=5)
        format_text = ("📋 新格式说明：\n"
                       "我的商品ID|商品名称|竞品链接|最低价|最高价|监控间隔|破价目标值|预期值\n\n"
                       "✅ 首页表格已显示破价目标值和预期值\n"
                       "最后两个字段是非必填项\n"
                       "不符合格式的内容会自动屏蔽\n"
                       "竞品链接：多个用逗号分隔")
        ttk.Label(self.left_frame, text=format_text, wraplength=200, 
                 foreground='blue', font=('Microsoft YaHei', 8)).pack(pady=5)
        
        # ============== 右侧 ==============
        # 顶部控制区
        self.control_frame = ttk.Frame(self.right_frame)
        self.control_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.current_doc_label = ttk.Label(self.control_frame, text="当前文档：未选择", 
                                          font=('Microsoft YaHei', 11, 'bold'))
        self.current_doc_label.pack(side=tk.LEFT)
        
        ttk.Separator(self.control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=15, fill=tk.Y)
        
        ttk.Button(self.control_frame, text="➕ 添加商品", command=self.add_product_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="✏️ 编辑商品", command=self.edit_product_dialog).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="✅ 开启监控", command=self.enable_monitor).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="⏸️ 暂停监控", command=self.disable_monitor).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="🗑️ 删除商品", command=self.delete_product).pack(side=tk.LEFT, padx=5)
        ttk.Button(self.control_frame, text="🗑️ 删除商品", command=self.delete_product).pack(side=tk.LEFT, padx=5)

        
        # 分隔符
        
        # 分隔符
        ttk.Separator(self.control_frame, orient=tk.VERTICAL).pack(side=tk.LEFT, padx=10, fill=tk.Y)
        
        # 两轮循环按钮
        self.two_rounds_enable_btn = ttk.Button(self.control_frame, text="🔄 开启循环", command=self.enable_two_rounds)
        self.two_rounds_enable_btn.pack(side=tk.LEFT, padx=2)
        
        self.two_rounds_disable_btn = ttk.Button(self.control_frame, text="🔴 关闭循环", command=self.disable_two_rounds)
        self.two_rounds_disable_btn.pack(side=tk.LEFT, padx=2)
        
        
        # 商品列表表格 - 明确显示所有列，包括破价目标值和预期值
        columns = ("status", "id", "name", "links", "min_price", "max_price", "interval", "target_price", "expected", "two_rounds")
        self.tree = ttk.Treeview(self.right_frame, columns=columns, show="headings", height=20)
        
        # 【修复】明确设置所有列标题，确保破价目标值和预期值显示
        self.tree.heading("status", text="状态")
        self.tree.heading("id", text="我的商品 ID")
        self.tree.heading("name", text="商品名称")
        self.tree.heading("links", text="竞品链接")
        self.tree.heading("min_price", text="最低价")
        self.tree.heading("max_price", text="最高价")
        self.tree.heading("interval", text="监控间隔(秒)")
        self.tree.heading("target_price", text="破价目标值")  # 明确显示
        self.tree.heading("expected", text="预期值")        # 明确显示
        self.tree.heading("two_rounds", text="两轮循环")  # 两轮循环
        
        # 【修复】设置合理的列宽，确保所有列都能看到
        self.tree.column("status", width=60, anchor=tk.CENTER)
        self.tree.column("id", width=90, anchor=tk.CENTER)
        self.tree.column("name", width=140)
        self.tree.column("links", width=240)
        self.tree.column("min_price", width=65, anchor=tk.CENTER)
        self.tree.column("max_price", width=65, anchor=tk.CENTER)
        self.tree.column("interval", width=75, anchor=tk.CENTER)
        self.tree.column("target_price", width=85, anchor=tk.CENTER)  # 确保显示
        self.tree.column("expected", width=80, anchor=tk.CENTER)    # 确保显示
        self.tree.column("two_rounds", width=80, anchor=tk.CENTER)  # 两轮循环
        
        # 添加滚动条
        self.tree_scroll = ttk.Scrollbar(self.right_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscroll=self.tree_scroll.set)
        
        # 添加水平滚动条
        self.tree_scroll_x = ttk.Scrollbar(self.right_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(xscroll=self.tree_scroll_x.set)
        
        # 布局
        self.tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree_scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        
        # 绑定选中事件和双击编辑
        self.tree.bind('<<TreeviewSelect>>', self.on_product_select)
        self.tree.bind('<Double-1>', self.on_double_click_edit)  # 双击编辑
        
        # 状态栏
        self.status_bar = ttk.Label(self.root, text="就绪 | 双击商品行可编辑 | ✅ 已显示破价目标值、预期值和两轮循环", relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X, side=tk.BOTTOM)
    
    def update_doc_list(self):
        """更新文档列表显示"""
        self.doc_listbox.delete(0, tk.END)
        for name in self.doc_manager.get_all_names():
            # 标记外部文档
            if self.doc_manager.is_external_doc(name):
                self.doc_listbox.insert(tk.END, f"📄 {name}")
            else:
                self.doc_listbox.insert(tk.END, f"📝 {name}")
    
    def on_doc_select(self, event):
        """文档选择事件"""
        selection = self.doc_listbox.curselection()
        if not selection:
            return
        
        # 加载选中的文档
        index = selection[0]
        # 去掉图标前缀获取实际名称
        full_text = self.doc_listbox.get(index)
        doc_name = full_text[2:] if len(full_text) > 2 and full_text[0] in ['📄', '📝'] else full_text
        
        content = self.doc_manager.get_document(doc_name)
        if content is None:
            return
        
        self.doc_manager.current_doc = doc_name
        
        # 显示是否为外部文件
        ext_path = self.doc_manager.get_external_path(doc_name)
        if ext_path:
            self.current_doc_label.config(text=f"当前文档：{doc_name} 📄 (外部文件)")
            self.update_status(f"已打开外部文件：{ext_path}")
        else:
            self.current_doc_label.config(text=f"当前文档：{doc_name}")
            self.update_status(f"已打开文档：{doc_name}")
        
        # 解析并加载商品列表（过滤不符合格式的内容）
        self.load_products_to_tree(content)
    
    def refresh_current_doc(self):
        """刷新当前文档内容（从外部文件重新加载）"""
        if self.doc_manager.current_doc is None:
            messagebox.showwarning("提示", "请先选择一个文档")
            return
        
        doc_name = self.doc_manager.current_doc
        
        # 检查是否为外部文件
        if not self.doc_manager.is_external_doc(doc_name):
            messagebox.showinfo("提示", "当前文档是内部文档，无需刷新\n如需重新加载，请重新选择该文档")
            return
        
        # 重新加载内容
        success, msg = self.doc_manager.reload_from_external(doc_name)
        if success:
            # 重新加载到界面
            content = self.doc_manager.get_document(doc_name)
            self.load_products_to_tree(content)
            self.update_status(f"已刷新文档「{doc_name}」的内容")
            messagebox.showinfo("成功", msg)
        else:
            messagebox.showerror("错误", msg)
    
    def parse_line(self, line):
        """
        解析一行数据，新格式：
        我的商品 ID|商品名称 | 竞品链接 | 最低价 | 最高价 | 监控间隔 | 破价目标值 | 预期值
        最后两个字段为非必填项
        返回：(是否启用，[id, name, links, min_price, max_price, interval, target_price, expected])
        不符合要求的格式返回 None（会被屏蔽不显示）
        """
        line = line.strip()
        if not line:
            return None  # 空行屏蔽
        
        # 检查是否被注释（暂停监控）
        enabled = not line.startswith('#')
        if not enabled:
            line = line[1:].strip()  # 去掉注释符号
        
        # 用 | 分割
        parts = line.split('|')
        
        # 去掉前后空格
        parts = [p.strip() for p in parts]
        
        # 最少需要 6 个字段（前 6 个必填，最后两个可选）
        if len(parts) < 6:
            return None  # 字段不足，屏蔽不显示
        
        # 只取前 9 个字段，多余忽略
        parts = parts[:9]
        
        # 补足空字段到 9 个
        while len(parts) < 9:
            parts.append('')
        
        # ID 和名称必须有内容，否则屏蔽
        if not parts[0] or not parts[1]:
            return None
        
        return enabled, parts
    
    def load_products_to_tree(self, content):
        """将内容加载到树形列表，过滤不符合格式的内容"""
        # 清空现有项
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 逐行解析
        lines = content.splitlines()
        skipped = 0
        loaded = 0
        
        for line in lines:
            result = self.parse_line(line)
            if result is None:
                skipped += 1
                continue
            
            enabled, parts = result
            loaded += 1
            status = "✅ 监控中" if enabled else "⏸️ 已暂停"
            values = [status] + parts  # parts 已经包含 target 和 expected
            # 设置标签颜色
            tags = ("enabled" if enabled else "disabled",)
            self.tree.insert("", tk.END, values=values, tags=tags)
        
        # 设置颜色
        self.tree.tag_configure("enabled", background="#f0f9f0")
        self.tree.tag_configure("disabled", background="#f0f0f0", foreground="#888888")
        
        if skipped > 0:
            self.update_status(f"加载完成：共 {loaded} 条有效商品，跳过 {skipped} 条不符合格式的记录 | ✅ 已显示破价目标值、预期值和两轮循环")
        else:
            self.update_status(f"加载完成：共 {loaded} 条有效商品 | ✅ 已显示破价目标值、预期值和两轮循环")
    
    def get_all_lines_content(self):
        """从树形列表获取所有行内容，组装成 txt 内容"""
        content = []
        for item in self.tree.get_children():
            values = self.tree.item(item)['values']
            # values: [status, id, name, links, min, max, interval, target, expected]
            # 【修复】强制转换为字符串，防止整数导致 join 失败
            id_ = str(values[1])
            name = str(values[2])
            links = str(values[3])
            min_price = str(values[4])
            max_price = str(values[5])
            interval = str(values[6])
            target = str(values[7]) if len(values) > 7 else ''
            expected = str(values[8]) if len(values) > 8 else ''
            two_rounds = str(values[9]) if len(values) > 9 else ''
            
            # 组装格式：我的商品 ID|商品名称 | 竞品链接 | 最低价 | 最高价 | 监控间隔 | 破价目标值 | 预期值 | 两轮循环
            line_parts = [id_, name, links, min_price, max_price, interval]
            
            # 只有非空才添加最后三个字段（兼容旧格式）
            if target:
                line_parts.append(target)
                if expected:
                    line_parts.append(expected)
                    if two_rounds:
                        line_parts.append(two_rounds)
            
            line = "|".join(line_parts)
            
            # 检查状态
            if "已暂停" in values[0]:
                line = "# " + line
            
            content.append(line)
        
        return '\n'.join(content)
    
    def on_product_select(self, event):
        """商品选中事件"""
        selection = self.tree.selection()
        if selection:
            self.current_row = selection[0]
        else:
            self.current_row = None
    
    def on_double_click_edit(self, event):
        """双击编辑商品"""
        self.edit_product_dialog()
    
    def add_product_dialog(self):
        """打开添加商品对话框"""
        if self.doc_manager.current_doc is None:
            messagebox.showwarning("提示", "请先选择一个监控文档")
            return
        
        self._show_product_dialog("添加新商品监控", None)
    
    def edit_product_dialog(self):
        """打开编辑商品对话框"""
        if not self.current_row:
            messagebox.showwarning("提示", "请先选择要编辑的商品\n(可双击商品行快速编辑)")
            return
        
        # 获取当前选中商品的数据
        values = self.tree.item(self.current_row)['values']
        self._show_product_dialog("编辑商品监控", values)
    
    def _show_product_dialog(self, title, existing_values):
        """显示商品添加/编辑对话框"""
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry("640x650")  # 加高对话框以容纳所有字段（含两轮循环）
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # 居中
        dialog.geometry("+%d+%d" % (self.root.winfo_rootx()+70, self.root.winfo_rooty()+20))
        
        form_frame = ttk.Frame(dialog, padding=20)
        form_frame.pack(fill=tk.BOTH, expand=True)
        
        # 我的商品 ID
        ttk.Label(form_frame, text="我的商品 ID:", font=('Microsoft YaHei', 10)).grid(row=0, column=0, sticky=tk.W, pady=6)

        # WPS同步按钮
        self.wps_sync_btn = ttk.Button(control_frame, text="📥 同步WPS", command=self.sync_wps_doc)
        self.wps_sync_btn.grid(row=0, column=1, padx=5, pady=5)
        self.push_wps_btn = ttk.Button(control_frame, text="📤 推送WPS", command=self.push_to_wps)
        self.push_wps_btn.grid(row=0, column=2, padx=5, pady=5)

        id_var = tk.StringVar()
        id_entry = ttk.Entry(form_frame, textvariable=id_var, width=50)
        id_entry.grid(row=0, column=1, sticky=tk.W, pady=6)
        
        # 商品名称
        ttk.Label(form_frame, text="商品名称:", font=('Microsoft YaHei', 10)).grid(row=1, column=0, sticky=tk.W, pady=6)
        name_var = tk.StringVar()
        name_entry = ttk.Entry(form_frame, textvariable=name_var, width=50)
        name_entry.grid(row=1, column=1, sticky=tk.W, pady=6)
        
        # 竞品链接（多个用逗号分隔）
        ttk.Label(form_frame, text="竞品链接:", font=('Microsoft YaHei', 10)).grid(row=2, column=0, sticky=tk.W, pady=6)
        
        # 链接输入框和按钮的框架
        links_frame = ttk.Frame(form_frame)
        links_frame.grid(row=2, column=1, sticky=tk.W+tk.E, pady=6)
        
        links_var = tk.StringVar()
        links_entry = ttk.Entry(links_frame, textvariable=links_var, width=43)
        links_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # 链接管理按钮
        manage_btn = ttk.Button(links_frame, text="🔗 管理", command=lambda: self.manage_links(links_var), width=8)
        manage_btn.pack(side=tk.LEFT, padx=(5, 0))
        
        # 最低价
        ttk.Label(form_frame, text="最低价:", font=('Microsoft YaHei', 10)).grid(row=3, column=0, sticky=tk.W, pady=6)
        min_var = tk.StringVar()
        min_entry = ttk.Entry(form_frame, textvariable=min_var, width=50)
        min_entry.grid(row=3, column=1, sticky=tk.W, pady=6)
        
        # 最高价
        ttk.Label(form_frame, text="最高价:", font=('Microsoft YaHei', 10)).grid(row=4, column=0, sticky=tk.W, pady=6)
        max_var = tk.StringVar()
        max_entry = ttk.Entry(form_frame, textvariable=max_var, width=50)
        max_entry.grid(row=4, column=1, sticky=tk.W, pady=6)
        
        # 监控间隔
        ttk.Label(form_frame, text="监控间隔 (秒):", font=('Microsoft YaHei', 10)).grid(row=5, column=0, sticky=tk.W, pady=6)
        interval_var = tk.StringVar(value="3600")
        interval_entry = ttk.Entry(form_frame, textvariable=interval_var, width=50)
        interval_entry.grid(row=5, column=1, sticky=tk.W, pady=6)
        
        # 【修复】破价目标值（非必填）- 明确显示
        ttk.Label(form_frame, text="破价目标值:\n(非必填)", font=('Microsoft YaHei', 10)).grid(row=6, column=0, sticky=tk.W, pady=6)
        target_var = tk.StringVar()
        target_entry = ttk.Entry(form_frame, textvariable=target_var, width=50)
        target_entry.grid(row=6, column=1, sticky=tk.W, pady=6)
        
        # 【修复】预期值（非必填）- 明确显示
        ttk.Label(form_frame, text="预期值:\n(非必填)", font=('Microsoft YaHei', 10)).grid(row=7, column=0, sticky=tk.W, pady=6)
        expected_var = tk.StringVar()
        expected_entry = ttk.Entry(form_frame, textvariable=expected_var, width=50)
        expected_entry.grid(row=7, column=1, sticky=tk.W, pady=6)
        
        # 两轮循环（非必填）
        ttk.Label(form_frame, text="两轮循环:\n(非必填)", font=('Microsoft YaHei', 10)).grid(row=8, column=0, sticky=tk.W, pady=6)
        two_rounds_var = tk.StringVar()
        two_rounds_entry = ttk.Entry(form_frame, textvariable=two_rounds_var, width=50)
        two_rounds_entry.grid(row=8, column=1, sticky=tk.W, pady=6)
        
        # 格式提示
        ttk.Label(form_frame, text="新格式：我的商品 ID|商品名称 | 竞品链接 | 最低价 | 最高价 | 监控间隔 | 破价目标值 | 预期值 | 两轮循环\n"
                 "最后三个字段为非必填，不符合格式内容会自动屏蔽\n✅ 对话框已包含所有字段输入框", 
                 foreground='gray', font=('Microsoft YaHei', 8)).grid(row=9, column=0, columnspan=2, pady=8)
        
        # 如果是编辑模式，填充现有数据
        is_edit_mode = existing_values is not None
        if is_edit_mode:
            # existing_values: [status, id, name, links, min, max, interval, target, expected]
            id_var.set(str(existing_values[1]))
            name_var.set(str(existing_values[2]))
            links_var.set(str(existing_values[3]))
            min_var.set(str(existing_values[4]))
            max_var.set(str(existing_values[5]))
            if len(existing_values) >= 7:
                interval_var.set(str(existing_values[6]))
            if len(existing_values) >= 8:
                target_var.set(str(existing_values[7]))  # 加载破价目标值
            if len(existing_values) >= 9:
                expected_var.set(str(existing_values[8]))  # 加载预期值
            if len(existing_values) >= 10:
                two_rounds_var.set(str(existing_values[9]))  # 加载两轮循环
        
        def confirm():
            """确认添加/编辑"""
            try:
                id_ = id_var.get().strip()
                name = name_var.get().strip()
                links = links_var.get().strip()
                min_p = min_var.get().strip()
                max_p = max_var.get().strip()
                interval = interval_var.get().strip()
                target = target_var.get().strip()  # 获取破价目标值
                expected = expected_var.get().strip()  # 获取预期值
                two_rounds = two_rounds_var.get().strip()  # 获取两轮循环
                
                # 前 6 项必填
                if not id_ or not name:
                    messagebox.showwarning("提示", "商品 ID 和商品名称不能为空")
                    return
                
                # 获取当前状态（编辑时保持原状态，添加时默认启用）
                if is_edit_mode:
                    status = existing_values[0]
                    tags = ("enabled" if "监控中" in status else "disabled",)
                else:
                    status = "✅ 监控中"
                    tags = ("enabled",)
                
                # 如果两轮循环有内容，但破价或预期为空，自动填 0（无论全局开关是否开启）
                if two_rounds:
                    if not target:
                        target = '0'
                    if not expected:
                        expected = '0'
                
                # 组装所有值（包括最后三个可选字段，都显示在表格）
                values = [status, id_, name, links, min_p, max_p, interval, target, expected, two_rounds]
                
                if is_edit_mode:
                    # 编辑模式：更新现有行
                    self.tree.item(self.current_row, values=values, tags=tags)
                    action_text = "编辑"
                else:
                    # 添加模式：插入新行
                    self.tree.insert("", tk.END, values=values, tags=tags)
                    action_text = "添加"
                
                # 保存到文档
                success, msg = self.save_current_to_doc()
                if success:
                    messagebox.showinfo("成功", f"商品「{name}」{action_text}成功\n✅ 已同步保存到文档\n✅ 破价目标值、预期值和两轮循环已保存")
                    dialog.destroy()
                    self.update_status(f"已{action_text}商品：{name}")
                else:
                    messagebox.showerror("保存失败", f"数据已添加到列表，但保存失败：\n{msg}\n请检查文件权限或磁盘空间")
                    
            except Exception as e:
                messagebox.showerror("错误", f"处理数据时出错：{str(e)}")
        
        def cancel():
            dialog.destroy()
        
        # 按钮
        btn_frame = ttk.Frame(form_frame)
        btn_frame.grid(row=10, column=0, columnspan=2, pady=12)
        
        btn_text = "✅ 保存修改" if is_edit_mode else "✅ 添加"
        ttk.Button(btn_frame, text=btn_text, command=confirm, width=15).pack(side=tk.LEFT, padx=10)
        ttk.Button(btn_frame, text="❌ 取消", command=cancel, width=15).pack(side=tk.LEFT, padx=10)
        
        # 绑定回车键确认
        id_entry.bind('<Return>', lambda e: confirm())
        name_entry.bind('<Return>', lambda e: confirm())
    
    def manage_links(self, links_var):
        """打开链接管理器"""
        # 获取当前链接
        current_links = links_var.get()
        
        # 打开链接管理器对话框
        link_dialog = LinkManagerDialog(self.root, current_links)
        self.root.wait_window(link_dialog.dialog)
        
        # 获取结果
        result = link_dialog.get_result()
        if result is not None:  # 注意：空字符串也是有效结果
            links_var.set(result)
    
    def enable_monitor(self):
        """开启选中商品监控：去除开头#"""
        if not self.current_row:
            messagebox.showwarning("提示", "请先选择要开启监控的商品")
            return
        
        item = self.current_row
        values = list(self.tree.item(item)['values'])
        
        # 如果已经是监控中，不需要操作
        if "监控中" in values[0]:
            self.update_status("该商品已在监控中，无需操作")
            return
        
        # 设置状态为开启
        values[0] = "✅ 监控中"
        self.tree.item(item, values=values, tags=("enabled",))
        success, msg = self.save_current_to_doc()
        if success:
            self.update_status("已开启商品监控并保存")
        else:
            messagebox.showerror("保存失败", msg)
    
    def disable_monitor(self):
        """暂停选中商品监控：添加开头#"""
        if not self.current_row:
            messagebox.showwarning("提示", "请先选择要暂停监控的商品")
            return
        
        item = self.current_row
        values = list(self.tree.item(item)['values'])
        
        # 如果已经是暂停状态，不需要操作
        if "已暂停" in values[0]:
            self.update_status("该商品已暂停监控，无需操作")
            return
        
        # 设置状态为暂停
        values[0] = "⏸️ 已暂停"
        self.tree.item(item, values=values, tags=("disabled",))
        success, msg = self.save_current_to_doc()
        if success:
            self.update_status("已暂停商品监控并保存")
        else:
            messagebox.showerror("保存失败", msg)




    def enable_two_rounds(self):
        """对选中商品开启两轮循环"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要开启两轮循环的商品（可按住 Ctrl 多选）")
            return
        
        modified = 0
        filled = 0
        for item in selection:
            values = list(self.tree.item(item)['values'])
            # values: [status, id, name, links, min, max, interval, target, expected, two_rounds]
            if len(values) >= 10:
                # 将两轮循环设为"是"
                values[9] = "是"
                # 如果破价目标值为空，填 0
                target = str(values[7]).strip()
                if not target:
                    values[7] = '0'
                    filled += 1
                # 如果预期值为空，填 0
                expected = str(values[8]).strip()
                if not expected:
                    values[8] = '0'
                    filled += 1
                # 更新行
                self.tree.item(item, values=values)
                modified += 1
        
        # 保存到文档
        if modified > 0:
            self.save_current_to_doc()
            if filled > 0:
                messagebox.showinfo("开启循环完成", f"✅ 已开启两轮循环\n{modified} 个选中商品已设为「是」\n自动为 {filled} 个空值填充 0")
                self.update_status(f"已开启 {modified} 个商品的两轮循环，自动填充 {filled} 个空值")
            else:
                messagebox.showinfo("开启循环完成", f"✅ 已开启两轮循环\n{modified} 个选中商品已设为「是」")
                self.update_status(f"已开启 {modified} 个商品的两轮循环")
        else:
            messagebox.showwarning("提示", "没有选中任何商品")


    def disable_two_rounds(self):
        """对选中商品关闭两轮循环"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要关闭两轮循环的商品（可按住 Ctrl 多选）")
            return
        
        modified = 0
        for item in selection:
            values = list(self.tree.item(item)['values'])
            # values: [status, id, name, links, min, max, interval, target, expected, two_rounds]
            if len(values) >= 10:
                # 将两轮循环设为"否"
                values[9] = "否"
                # 更新行
                self.tree.item(item, values=values)
                modified += 1
        
        # 保存到文档
        if modified > 0:
            self.save_current_to_doc()
            messagebox.showinfo("关闭循环完成", f"🔴 已关闭两轮循环\n{modified} 个选中商品已设为「否」")
            self.update_status(f"已关闭 {modified} 个商品的两轮循环")
        else:
            messagebox.showwarning("提示", "没有选中任何商品")

    def delete_product(self):
        """删除选中商品"""
        if not self.current_row:
            messagebox.showwarning("提示", "请先选择要删除的商品")
            return
        
        # 获取商品名称用于确认提示
        values = self.tree.item(self.current_row)['values']
        product_name = values[2] if len(values) > 2 else "未知商品"
        
        if not messagebox.askyesno("确认删除", f"确定要删除商品「{product_name}」吗？\n此操作不可恢复"):
            return
        
        self.tree.delete(self.current_row)
        self.current_row = None
        success, msg = self.save_current_to_doc()
        if success:
            self.update_status("已删除选中商品并保存")
        else:
            messagebox.showerror("保存失败", msg)
    
    def save_current_to_doc(self):
        """保存当前商品列表到文档"""
        if self.doc_manager.current_doc is None:
            return False, "未选择文档"
        
        try:
            content = self.get_all_lines_content()
            success, msg = self.doc_manager.update_document(self.doc_manager.current_doc, content)
            return success, msg
        except Exception as e:
            return False, f"保存过程出错：{str(e)}"
    
    def save_current_doc(self):
        """手动保存当前文档"""
        if self.doc_manager.current_doc is None:
            messagebox.showwarning("提示", "请先选择一个文档")
            return
        
        success, msg = self.save_current_to_doc()
        if success:
            messagebox.showinfo("成功", f"✅ {msg}\n外部文件已同步更新！")
        else:
            messagebox.showerror("错误", msg)
    
    def new_document(self):
        """新建文档"""
        dialog = tk.Toplevel(self.root)
        dialog.title("新建监控文档")
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="请输入文档名称：", font=('Microsoft YaHei', 10)).pack(pady=15)
        name_var = tk.StringVar()
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        
        def confirm():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("提示", "文档名称不能为空")
                return
            
            success, msg = self.doc_manager.create_document(name, "")
            if success:
                self.update_doc_list()
                messagebox.showinfo("成功", msg)
                dialog.destroy()
            else:
                messagebox.showerror("错误", msg)
        
        ttk.Button(dialog, text="确定", command=confirm).pack(pady=20)
    
    def rename_document(self):
        """重命名文档"""
        selection = self.doc_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要重命名的文档")
            return
        
        full_text = self.doc_listbox.get(selection[0])
        old_name = full_text[2:] if len(full_text) > 2 and full_text[0] in ['📄', '📝'] else full_text
        
        dialog = tk.Toplevel(self.root)
        dialog.title("重命名文档")
        dialog.geometry("350x150")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"原名称：{old_name}\n请输入新名称：").pack(pady=15)
        name_var = tk.StringVar(value=old_name)
        entry = ttk.Entry(dialog, textvariable=name_var, width=30)
        entry.pack(pady=5)
        entry.focus()
        
        def confirm():
            new_name = name_var.get().strip()
            if not new_name:
                messagebox.showwarning("提示", "新名称不能为空")
                return
            
            success, msg = self.doc_manager.rename_document(old_name, new_name)
            if success:
                self.update_doc_list()
                if self.doc_manager.current_doc == old_name:
                    self.doc_manager.current_doc = new_name
                    ext_path = self.doc_manager.get_external_path(new_name)
                    if ext_path:
                        self.current_doc_label.config(text=f"当前文档：{new_name} 📄 (外部文件)")
                    else:
                        self.current_doc_label.config(text=f"当前文档：{new_name}")
                messagebox.showinfo("成功", msg)
                dialog.destroy()
            else:
                messagebox.showerror("错误", msg)
        
        ttk.Button(dialog, text="确定", command=confirm).pack(pady=20)
    
    def delete_document(self):
        """删除文档"""
        selection = self.doc_listbox.curselection()
        if not selection:
            messagebox.showwarning("提示", "请先选择要删除的文档")
            return
        
        full_text = self.doc_listbox.get(selection[0])
        doc_name = full_text[2:] if len(full_text) > 2 and full_text[0] in ['📄', '📝'] else full_text
        
        if not messagebox.askyesno("确认删除", f"确定要删除文档「{doc_name}」吗？\n此操作不会删除关联的外部 txt 文件"):
            return
        
        success, msg = self.doc_manager.delete_document(doc_name)
        if success:
            self.update_doc_list()
            if self.doc_manager.current_doc == doc_name:
                self.doc_manager.current_doc = None
                self.current_doc_label.config(text="当前文档：未选择")
                # 清空商品列表
                for item in self.tree.get_children():
                    self.tree.delete(item)
            messagebox.showinfo("成功", msg)
            self.update_status(msg)
        else:
            messagebox.showerror("错误", msg)
    
    def open_external_txt(self):
        """打开外部 txt 文件进行关联"""
        file_path = filedialog.askopenfilename(
            title="选择要关联的 txt 文件",
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")]
        )
        if not file_path:
            return
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            # 使用文件名作为默认文档名
            default_name = os.path.splitext(os.path.basename(file_path))[0]
            
            dialog = tk.Toplevel(self.root)
            dialog.title("关联外部文件")
            dialog.geometry("520x240")  # 增加高度，确保按钮显示完整
            dialog.transient(self.root)
            dialog.grab_set()
            
            ttk.Label(dialog, text=f"已选择文件：{os.path.basename(file_path)}\n请输入文档名称：", 
                     font=('Microsoft YaHei', 10)).pack(pady=15)
            name_var = tk.StringVar(value=default_name)
            entry = ttk.Entry(dialog, textvariable=name_var, width=42)
            entry.pack(pady=5)
            entry.focus()
            
            ttk.Label(dialog, text=f"⚠️ 保存时会直接修改原文件：\n{file_path}", 
                     foreground='orange', font=('Microsoft YaHei', 9)).pack(pady=10)
            
            def confirm():
                name = name_var.get().strip()
                if not name:
                    messagebox.showwarning("提示", "文档名称不能为空")
                    return
                
                success, msg = self.doc_manager.create_document(name, content, file_path)
                if success:
                    self.update_doc_list()
                    messagebox.showinfo("成功", f"文件关联成功，保存会直接修改原文件")
                    dialog.destroy()
                else:
                    messagebox.showerror("错误", msg)
            
            ttk.Button(dialog, text="确定关联", command=confirm).pack(pady=20)
            
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败：{str(e)}")
    
    def update_status(self, message):
        """更新状态栏"""
        self.status_bar.config(text=f"{message} | 双击商品行可编辑 | ✅ 已显示破价目标值、预期值和两轮循环")
def main():
    root = tk.Tk()
    app = PriceMonitorApp(root)
    
    # 设置窗口最小大小
    root.minsize(1250, 550)
    
    # 居中显示窗口
    window_width = 1550
    window_height = 700
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x = (screen_width // 2) - (window_width // 2)
    y = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x}+{y}")
    
    root.mainloop()
if __name__ == "__main__":
    main()