"""
PDF 到 Word 转换器 - 图形用户界面版本
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
from pathlib import Path
import threading
from pdf_to_word_converter import PDFToWordConverter


class PDFToWordGUI:
    """PDF 到 Word 转换器图形界面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 到 Word 转换器")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # 设置样式
        style = ttk.Style()
        style.theme_use('clam')
        
        # 创建转换器实例
        self.converter = PDFToWordConverter()
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        """创建所有界面组件"""
        
        # 主容器
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)
        
        # 标题
        title_label = ttk.Label(
            main_frame, 
            text="📄 PDF 到 Word 转换器",
            font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, pady=10)
        
        # 创建选项卡
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        
        # 各个功能选项卡
        self.create_convert_tab()
        self.create_search_tab()
        self.create_highlight_tab()
        self.create_replace_tab()
        self.create_add_text_tab()
        self.create_info_tab()
        
        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(
            main_frame,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(5, 0))
        
    def create_convert_tab(self):
        """创建PDF转Word选项卡"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="PDF 转 Word")
        
        # PDF文件选择
        ttk.Label(tab, text="PDF 文件:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        pdf_frame = ttk.Frame(tab)
        pdf_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        pdf_frame.columnconfigure(0, weight=1)
        
        self.pdf_path_var = tk.StringVar()
        ttk.Entry(pdf_frame, textvariable=self.pdf_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(pdf_frame, text="浏览...", command=self.browse_pdf).grid(
            row=0, column=1
        )
        
        # Word输出文件
        ttk.Label(tab, text="输出 Word 文件:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        word_frame = ttk.Frame(tab)
        word_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        word_frame.columnconfigure(0, weight=1)
        
        self.word_path_var = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.word_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(word_frame, text="浏览...", command=self.browse_word_save).grid(
            row=0, column=1
        )
        
        ttk.Label(tab, text="(留空则自动生成文件名)", font=("Arial", 8)).grid(
            row=4, column=0, sticky=tk.W, pady=(0, 10)
        )
        
        # 转换按钮
        convert_btn = ttk.Button(
            tab,
            text="🔄 开始转换",
            command=self.convert_pdf,
            style="Accent.TButton"
        )
        convert_btn.grid(row=5, column=0, pady=20)
        
        # 进度信息
        self.convert_output = scrolledtext.ScrolledText(
            tab, height=15, width=70, wrap=tk.WORD
        )
        self.convert_output.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        tab.rowconfigure(6, weight=1)
        
    def create_search_tab(self):
        """创建关键词搜索选项卡"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="搜索关键词")
        
        # Word文件选择
        ttk.Label(tab, text="Word 文件:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        word_frame = ttk.Frame(tab)
        word_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        word_frame.columnconfigure(0, weight=1)
        
        self.search_word_path_var = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.search_word_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(word_frame, text="浏览...", command=self.browse_word_search).grid(
            row=0, column=1
        )
        
        # 关键词输入
        ttk.Label(tab, text="关键词:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        self.search_keyword_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.search_keyword_var, width=40).grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        
        # 搜索按钮
        ttk.Button(tab, text="🔍 搜索", command=self.search_keyword).grid(
            row=4, column=0, pady=20, sticky=tk.W
        )
        
        # 搜索结果
        ttk.Label(tab, text="搜索结果:", font=("Arial", 10, "bold")).grid(
            row=5, column=0, sticky=tk.W, pady=5
        )
        
        self.search_output = scrolledtext.ScrolledText(
            tab, height=20, width=70, wrap=tk.WORD
        )
        self.search_output.grid(row=6, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        tab.rowconfigure(6, weight=1)
        
    def create_highlight_tab(self):
        """创建高亮关键词选项卡"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="高亮关键词")
        
        # Word文件选择
        ttk.Label(tab, text="Word 文件:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        word_frame = ttk.Frame(tab)
        word_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        word_frame.columnconfigure(0, weight=1)
        
        self.highlight_word_path_var = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.highlight_word_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(word_frame, text="浏览...", command=self.browse_word_highlight).grid(
            row=0, column=1
        )
        
        # 关键词输入
        ttk.Label(tab, text="要高亮的关键词:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        self.highlight_keyword_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.highlight_keyword_var, width=40).grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        
        # 输出文件
        ttk.Label(tab, text="输出文件:", font=("Arial", 10, "bold")).grid(
            row=4, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        output_frame = ttk.Frame(tab)
        output_frame.grid(row=5, column=0, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        self.highlight_output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.highlight_output_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(output_frame, text="浏览...", command=self.browse_word_highlight_output).grid(
            row=0, column=1
        )
        
        ttk.Label(tab, text="(留空则自动生成文件名)", font=("Arial", 8)).grid(
            row=6, column=0, sticky=tk.W, pady=(0, 10)
        )
        
        # 高亮按钮
        ttk.Button(tab, text="✨ 高亮关键词", command=self.highlight_keyword).grid(
            row=7, column=0, pady=20, sticky=tk.W
        )
        
        # 输出信息
        self.highlight_output_text = scrolledtext.ScrolledText(
            tab, height=12, width=70, wrap=tk.WORD
        )
        self.highlight_output_text.grid(row=8, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        tab.rowconfigure(8, weight=1)
        
    def create_replace_tab(self):
        """创建文本替换选项卡"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="替换文本")
        
        # Word文件选择
        ttk.Label(tab, text="Word 文件:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        word_frame = ttk.Frame(tab)
        word_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        word_frame.columnconfigure(0, weight=1)
        
        self.replace_word_path_var = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.replace_word_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(word_frame, text="浏览...", command=self.browse_word_replace).grid(
            row=0, column=1
        )
        
        # 旧文本
        ttk.Label(tab, text="要替换的文本:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        self.old_text_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.old_text_var, width=50).grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        
        # 新文本
        ttk.Label(tab, text="新文本:", font=("Arial", 10, "bold")).grid(
            row=4, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        self.new_text_var = tk.StringVar()
        ttk.Entry(tab, textvariable=self.new_text_var, width=50).grid(
            row=5, column=0, sticky=tk.W, pady=5
        )
        
        # 输出文件
        ttk.Label(tab, text="输出文件:", font=("Arial", 10, "bold")).grid(
            row=6, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        output_frame = ttk.Frame(tab)
        output_frame.grid(row=7, column=0, sticky=(tk.W, tk.E), pady=5)
        output_frame.columnconfigure(0, weight=1)
        
        self.replace_output_var = tk.StringVar()
        ttk.Entry(output_frame, textvariable=self.replace_output_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(output_frame, text="浏览...", command=self.browse_word_replace_output).grid(
            row=0, column=1
        )
        
        ttk.Label(tab, text="(留空则自动生成文件名)", font=("Arial", 8)).grid(
            row=8, column=0, sticky=tk.W, pady=(0, 10)
        )
        
        # 替换按钮
        ttk.Button(tab, text="🔄 替换文本", command=self.replace_text).grid(
            row=9, column=0, pady=20, sticky=tk.W
        )
        
        # 输出信息
        self.replace_output_text = scrolledtext.ScrolledText(
            tab, height=10, width=70, wrap=tk.WORD
        )
        self.replace_output_text.grid(row=10, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        tab.rowconfigure(10, weight=1)
        
    def create_add_text_tab(self):
        """创建添加文本选项卡"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="添加文本")
        
        # Word文件选择
        ttk.Label(tab, text="Word 文件:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        word_frame = ttk.Frame(tab)
        word_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        word_frame.columnconfigure(0, weight=1)
        
        self.add_word_path_var = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.add_word_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(word_frame, text="浏览...", command=self.browse_word_add).grid(
            row=0, column=1
        )
        
        # 要添加的文本
        ttk.Label(tab, text="要添加的文本:", font=("Arial", 10, "bold")).grid(
            row=2, column=0, sticky=tk.W, pady=(20, 5)
        )
        
        self.add_text = scrolledtext.ScrolledText(tab, height=10, width=70, wrap=tk.WORD)
        self.add_text.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 添加按钮
        ttk.Button(tab, text="➕ 添加文本", command=self.add_text_to_doc).grid(
            row=4, column=0, pady=20, sticky=tk.W
        )
        
        # 输出信息
        self.add_output = scrolledtext.ScrolledText(
            tab, height=8, width=70, wrap=tk.WORD
        )
        self.add_output.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        tab.rowconfigure(5, weight=1)
        
    def create_info_tab(self):
        """创建文档信息选项卡"""
        tab = ttk.Frame(self.notebook, padding="20")
        self.notebook.add(tab, text="文档信息")
        
        # Word文件选择
        ttk.Label(tab, text="Word 文件:", font=("Arial", 10, "bold")).grid(
            row=0, column=0, sticky=tk.W, pady=5
        )
        
        word_frame = ttk.Frame(tab)
        word_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        word_frame.columnconfigure(0, weight=1)
        
        self.info_word_path_var = tk.StringVar()
        ttk.Entry(word_frame, textvariable=self.info_word_path_var, width=60).grid(
            row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10)
        )
        ttk.Button(word_frame, text="浏览...", command=self.browse_word_info).grid(
            row=0, column=1
        )
        
        # 查看按钮
        ttk.Button(tab, text="📊 查看文档信息", command=self.show_doc_info).grid(
            row=2, column=0, pady=20, sticky=tk.W
        )
        
        # 信息显示
        self.info_output = scrolledtext.ScrolledText(
            tab, height=20, width=70, wrap=tk.WORD, font=("Arial", 11)
        )
        self.info_output.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=10)
        tab.rowconfigure(3, weight=1)
        
    # 文件浏览对话框方法
    def browse_pdf(self):
        filename = filedialog.askopenfilename(
            title="选择 PDF 文件",
            filetypes=[("PDF 文件", "*.pdf"), ("所有文件", "*.*")]
        )
        if filename:
            self.pdf_path_var.set(filename)
            
    def browse_word_save(self):
        filename = filedialog.asksaveasfilename(
            title="保存 Word 文件",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.word_path_var.set(filename)
            
    def browse_word_search(self):
        filename = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.search_word_path_var.set(filename)
            
    def browse_word_highlight(self):
        filename = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.highlight_word_path_var.set(filename)
            
    def browse_word_highlight_output(self):
        filename = filedialog.asksaveasfilename(
            title="保存高亮后的文件",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.highlight_output_var.set(filename)
            
    def browse_word_replace(self):
        filename = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.replace_word_path_var.set(filename)
            
    def browse_word_replace_output(self):
        filename = filedialog.asksaveasfilename(
            title="保存替换后的文件",
            defaultextension=".docx",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.replace_output_var.set(filename)
            
    def browse_word_add(self):
        filename = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.add_word_path_var.set(filename)
            
    def browse_word_info(self):
        filename = filedialog.askopenfilename(
            title="选择 Word 文件",
            filetypes=[("Word 文件", "*.docx"), ("所有文件", "*.*")]
        )
        if filename:
            self.info_word_path_var.set(filename)
    
    # 功能实现方法
    def convert_pdf(self):
        """PDF转Word"""
        pdf_path = self.pdf_path_var.get().strip()
        word_path = self.word_path_var.get().strip() or None
        
        if not pdf_path:
            messagebox.showwarning("警告", "请选择 PDF 文件！")
            return
            
        if not os.path.exists(pdf_path):
            messagebox.showerror("错误", f"文件不存在: {pdf_path}")
            return
        
        self.convert_output.delete(1.0, tk.END)
        self.convert_output.insert(tk.END, f"正在转换 {pdf_path}...\n\n")
        self.status_var.set("正在转换...")
        
        def convert_thread():
            try:
                result = self.converter.convert_pdf_to_word(pdf_path, word_path)
                self.root.after(0, lambda: self.convert_output.insert(
                    tk.END, f"✓ 转换成功！\n\n文件保存在:\n{result}\n"
                ))
                self.root.after(0, lambda: self.status_var.set("转换完成"))
                self.root.after(0, lambda: messagebox.showinfo("成功", f"转换完成！\n\n文件保存在:\n{result}"))
            except Exception as e:
                self.root.after(0, lambda: self.convert_output.insert(
                    tk.END, f"✗ 转换失败:\n{str(e)}\n"
                ))
                self.root.after(0, lambda: self.status_var.set("转换失败"))
                self.root.after(0, lambda: messagebox.showerror("错误", f"转换失败:\n{str(e)}"))
        
        thread = threading.Thread(target=convert_thread)
        thread.daemon = True
        thread.start()
        
    def search_keyword(self):
        """搜索关键词"""
        word_path = self.search_word_path_var.get().strip()
        keyword = self.search_keyword_var.get().strip()
        
        if not word_path:
            messagebox.showwarning("警告", "请选择 Word 文件！")
            return
            
        if not keyword:
            messagebox.showwarning("警告", "请输入要搜索的关键词！")
            return
            
        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"文件不存在: {word_path}")
            return
        
        self.search_output.delete(1.0, tk.END)
        self.status_var.set("正在搜索...")
        
        try:
            results = self.converter.search_keyword(word_path, keyword)
            
            if results:
                self.search_output.insert(tk.END, f"找到 {len(results)} 处包含 '{keyword}' 的内容：\n\n")
                self.search_output.insert(tk.END, "=" * 70 + "\n\n")
                
                for idx, (para_idx, text) in enumerate(results, 1):
                    self.search_output.insert(tk.END, f"【结果 {idx}】段落 {para_idx}:\n")
                    self.search_output.insert(tk.END, f"{text}\n\n")
                    self.search_output.insert(tk.END, "-" * 70 + "\n\n")
                
                self.status_var.set(f"搜索完成 - 找到 {len(results)} 处结果")
            else:
                self.search_output.insert(tk.END, f"未找到关键词 '{keyword}'")
                self.status_var.set("未找到结果")
                
        except Exception as e:
            self.search_output.insert(tk.END, f"✗ 搜索失败:\n{str(e)}")
            self.status_var.set("搜索失败")
            messagebox.showerror("错误", f"搜索失败:\n{str(e)}")
            
    def highlight_keyword(self):
        """高亮关键词"""
        word_path = self.highlight_word_path_var.get().strip()
        keyword = self.highlight_keyword_var.get().strip()
        output_path = self.highlight_output_var.get().strip() or None
        
        if not word_path:
            messagebox.showwarning("警告", "请选择 Word 文件！")
            return
            
        if not keyword:
            messagebox.showwarning("警告", "请输入要高亮的关键词！")
            return
            
        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"文件不存在: {word_path}")
            return
        
        self.highlight_output_text.delete(1.0, tk.END)
        self.highlight_output_text.insert(tk.END, f"正在高亮关键词 '{keyword}'...\n\n")
        self.status_var.set("正在处理...")
        
        try:
            result = self.converter.highlight_keyword(word_path, keyword, output_path)
            self.highlight_output_text.insert(tk.END, f"✓ 高亮完成！\n\n文件保存在:\n{result}\n")
            self.status_var.set("高亮完成")
            messagebox.showinfo("成功", f"高亮完成！\n\n文件保存在:\n{result}")
        except Exception as e:
            self.highlight_output_text.insert(tk.END, f"✗ 高亮失败:\n{str(e)}\n")
            self.status_var.set("高亮失败")
            messagebox.showerror("错误", f"高亮失败:\n{str(e)}")
            
    def replace_text(self):
        """替换文本"""
        word_path = self.replace_word_path_var.get().strip()
        old_text = self.old_text_var.get().strip()
        new_text = self.new_text_var.get().strip()
        output_path = self.replace_output_var.get().strip() or None
        
        if not word_path:
            messagebox.showwarning("警告", "请选择 Word 文件！")
            return
            
        if not old_text:
            messagebox.showwarning("警告", "请输入要替换的文本！")
            return
            
        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"文件不存在: {word_path}")
            return
        
        self.replace_output_text.delete(1.0, tk.END)
        self.replace_output_text.insert(tk.END, f"正在替换文本...\n\n")
        self.status_var.set("正在处理...")
        
        try:
            result = self.converter.replace_text(word_path, old_text, new_text, output_path)
            self.replace_output_text.insert(tk.END, f"✓ 替换完成！\n\n文件保存在:\n{result}\n")
            self.status_var.set("替换完成")
            messagebox.showinfo("成功", f"替换完成！\n\n文件保存在:\n{result}")
        except Exception as e:
            self.replace_output_text.insert(tk.END, f"✗ 替换失败:\n{str(e)}\n")
            self.status_var.set("替换失败")
            messagebox.showerror("错误", f"替换失败:\n{str(e)}")
            
    def add_text_to_doc(self):
        """添加文本到文档"""
        word_path = self.add_word_path_var.get().strip()
        text = self.add_text.get(1.0, tk.END).strip()
        
        if not word_path:
            messagebox.showwarning("警告", "请选择 Word 文件！")
            return
            
        if not text:
            messagebox.showwarning("警告", "请输入要添加的文本！")
            return
            
        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"文件不存在: {word_path}")
            return
        
        self.add_output.delete(1.0, tk.END)
        self.add_output.insert(tk.END, f"正在添加文本...\n\n")
        self.status_var.set("正在处理...")
        
        try:
            result = self.converter.add_text_to_document(word_path, text)
            self.add_output.insert(tk.END, f"✓ 添加完成！\n\n文件已更新:\n{result}\n")
            self.status_var.set("添加完成")
            messagebox.showinfo("成功", f"添加完成！\n\n文件已更新:\n{result}")
        except Exception as e:
            self.add_output.insert(tk.END, f"✗ 添加失败:\n{str(e)}\n")
            self.status_var.set("添加失败")
            messagebox.showerror("错误", f"添加失败:\n{str(e)}")
            
    def show_doc_info(self):
        """显示文档信息"""
        word_path = self.info_word_path_var.get().strip()
        
        if not word_path:
            messagebox.showwarning("警告", "请选择 Word 文件！")
            return
            
        if not os.path.exists(word_path):
            messagebox.showerror("错误", f"文件不存在: {word_path}")
            return
        
        self.info_output.delete(1.0, tk.END)
        self.status_var.set("正在获取信息...")
        
        try:
            info = self.converter.get_document_info(word_path)
            
            self.info_output.insert(tk.END, "=" * 70 + "\n")
            self.info_output.insert(tk.END, "文档信息\n")
            self.info_output.insert(tk.END, "=" * 70 + "\n\n")
            
            self.info_output.insert(tk.END, f"文件路径: {word_path}\n\n")
            
            for key, value in info.items():
                self.info_output.insert(tk.END, f"• {key}: {value:,}\n")
            
            self.info_output.insert(tk.END, "\n" + "=" * 70 + "\n")
            
            self.status_var.set("信息获取完成")
        except Exception as e:
            self.info_output.insert(tk.END, f"✗ 获取信息失败:\n{str(e)}")
            self.status_var.set("获取信息失败")
            messagebox.showerror("错误", f"获取信息失败:\n{str(e)}")


def main():
    """启动GUI应用"""
    root = tk.Tk()
    app = PDFToWordGUI(root)
    
    # 设置窗口图标（如果有的话）
    try:
        root.iconbitmap('icon.ico')
    except:
        pass
    
    # 居中显示窗口
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f'{width}x{height}+{x}+{y}')
    
    root.mainloop()


if __name__ == "__main__":
    main()
