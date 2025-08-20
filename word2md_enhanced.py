# -*- coding: utf-8 -*-
"""
Enhanced Word to Markdown Converter
支持 .docx/.doc/.rtf 格式，完整保留图片、公式、表格
优化的图形界面和格式处理
"""

import os
import sys
import json
import shutil
import tempfile
import subprocess
import threading
import traceback
import re
import html
from pathlib import Path
from typing import Optional, Tuple, List

import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext

# 配置文件路径
CONFIG_FILE = os.path.join(os.path.expanduser("~"), ".word2md_enhanced_config.json")

def ensure_dependencies():
    """确保所有必要的依赖包都已安装"""
    required_packages = [
        ("docx", "python-docx"),
        ("mammoth", "mammoth"),
        ("pypandoc", "pypandoc"),
        ("PIL", "Pillow"),
        ("lxml", "lxml")
    ]
    
    missing_packages = []
    for module_name, pip_name in required_packages:
        try:
            __import__(module_name)
        except ImportError:
            missing_packages.append(pip_name)
    
    if missing_packages:
        print(f"正在安装缺失的依赖包: {', '.join(missing_packages)}")
        try:
            subprocess.check_call([
                sys.executable, "-m", "pip", "install"
            ] + missing_packages)
            print("依赖包安装完成")
        except subprocess.CalledProcessError as e:
            print(f"依赖包安装失败: {e}")
            return False
    return True

# 确保依赖包
if not ensure_dependencies():
    sys.exit(1)

# 导入依赖包
from docx import Document
import mammoth
from PIL import Image
import io
from lxml import etree
import pypandoc

class EnhancedWordToMarkdownConverter:
    def __init__(self):
        self.output_dir = None
        self.image_dir = None
        self.image_counter = 0
        self.pandoc_available = self._check_pandoc()
        
    def _check_pandoc(self) -> bool:
        """检查Pandoc是否可用"""
        try:
            pypandoc.get_pandoc_version()
            return True
        except OSError:
            try:
                pypandoc.download_pandoc()
                return True
            except Exception:
                return False
    
    def extract_images_from_docx(self, docx_path: str, output_folder: str) -> List[str]:
        """从docx文件中提取图片"""
        doc = Document(docx_path)
        self.image_counter = 0
        image_paths = []
        
        images_dir = Path(output_folder) / "images"
        images_dir.mkdir(exist_ok=True)
        
        for rel in doc.part.rels.values():
            if "image" in rel.reltype:
                try:
                    image_data = rel.target_part.blob
                    ext = rel.target_part.partname.split('.')[-1].lower()
                    if ext not in ['png', 'jpg', 'jpeg', 'gif', 'bmp', 'tiff', 'webp']:
                        ext = 'png'
                    
                    image_name = f"image_{self.image_counter:03d}.{ext}"
                    image_path = images_dir / image_name
                    
                    with open(image_path, 'wb') as f:
                        f.write(image_data)
                    
                    image_paths.append(str(image_path))
                    self.image_counter += 1
                except Exception as e:
                    print(f"提取图片时出错: {e}")
                    continue
                
        return image_paths
    
    def process_math_equations(self, content: str) -> str:
        """处理数学公式，将常见符号转换为LaTeX格式"""
        math_symbols = {
            'α': r'\alpha', 'β': r'\beta', 'γ': r'\gamma', 'δ': r'\delta',
            'ε': r'\varepsilon', 'θ': r'\theta', 'λ': r'\lambda', 'μ': r'\mu',
            'π': r'\pi', 'σ': r'\sigma', 'φ': r'\varphi', 'ω': r'\omega',
            '∑': r'\sum', '∫': r'\int', '∞': r'\infty', '∂': r'\partial',
            '±': r'\pm', '×': r'\times', '÷': r'\div', '≤': r'\leq',
            '≥': r'\geq', '≠': r'\neq', '≈': r'\approx', '√': r'\sqrt',
            '²': r'^2', '³': r'^3', '¹': r'^1', '₀': r'_0', '₁': r'_1',
            '₂': r'_2', '₃': r'_3', '₄': r'_4'
        }
        
        for symbol, latex in math_symbols.items():
            if symbol in content:
                pattern = re.compile(re.escape(symbol))
                matches = list(pattern.finditer(content))
                
                for match in reversed(matches):
                    start, end = match.span()
                    before = content[:start]
                    after = content[end:]
                    
                    in_math = (before.count('$') % 2 == 1)
                    
                    if not in_math:
                        replacement = f"${latex}$"
                    else:
                        replacement = latex
                    
                    content = content[:start] + replacement + content[end:]
        
        return content
    
    def convert_with_pandoc(self, docx_path: str, output_path: str) -> Tuple[str, List[str]]:
        """使用Pandoc进行转换"""
        if not self.pandoc_available:
            raise RuntimeError("Pandoc不可用")
        
        self.output_dir = Path(output_path).parent
        base_name = Path(docx_path).stem
        media_dir = self.output_dir / f"{base_name}_media"
        
        extra_args = [
            "--wrap=none",
            f"--extract-media={media_dir}",
            "--standalone"
        ]
        
        markdown_content = pypandoc.convert_file(
            docx_path,
            'gfm+tex_math_dollars+pipe_tables',
            extra_args=extra_args
        )
        
        markdown_content = self.process_math_equations(markdown_content)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        return markdown_content, []
    
    def convert_with_mammoth(self, docx_path: str, output_path: str) -> Tuple[str, List[str]]:
        """使用Mammoth进行转换"""
        self.output_dir = Path(output_path).parent
        self.image_dir = self.output_dir / "images"
        self.image_dir.mkdir(exist_ok=True)
        
        image_paths = self.extract_images_from_docx(docx_path, self.output_dir)
        
        def convert_image(image):
            try:
                with image.open() as image_bytes:
                    image_data = image_bytes.read()
                
                ext = 'png'
                if hasattr(image, 'content_type'):
                    if 'jpeg' in image.content_type:
                        ext = 'jpg'
                    elif 'png' in image.content_type:
                        ext = 'png'
                    elif 'gif' in image.content_type:
                        ext = 'gif'
                
                image_name = f"image_{self.image_counter:03d}.{ext}"
                image_path = self.image_dir / image_name
                
                with open(image_path, 'wb') as f:
                    f.write(image_data)
                
                self.image_counter += 1
                
                return {"src": f"images/{image_name}"}
            except Exception as e:
                print(f"处理图片时出错: {e}")
                return {"src": ""}
        
        with open(docx_path, "rb") as docx_file:
            result = mammoth.convert_to_html(
                docx_file,
                convert_image=mammoth.images.img_element(convert_image)
            )
        
        html_content = result.value
        html_content = self.process_math_equations(html_content)
        markdown_content = self.html_to_markdown(html_content)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
        
        return markdown_content, result.messages
    
    def html_to_markdown(self, html_content: str) -> str:
        """将HTML转换为Markdown"""
        # 转换标题
        for i in range(6, 0, -1):
            html_content = re.sub(
                f'<h{i}[^>]*>(.*?)</h{i}>',
                lambda m: f"{'#' * i} {m.group(1).strip()}\n\n",
                html_content,
                flags=re.DOTALL | re.IGNORECASE
            )
        
        # 转换段落
        html_content = re.sub(
            r'<p[^>]*>(.*?)</p>',
            lambda m: f"{m.group(1).strip()}\n\n" if m.group(1).strip() else "",
            html_content,
            flags=re.DOTALL | re.IGNORECASE
        )
        
        # 转换格式
        html_content = re.sub(r'<(b|strong)[^>]*>(.*?)</\1>', r'**\2**', html_content, flags=re.IGNORECASE | re.DOTALL)
        html_content = re.sub(r'<(i|em)[^>]*>(.*?)</\1>', r'*\2*', html_content, flags=re.IGNORECASE | re.DOTALL)
        html_content = re.sub(r'<a[^>]*href="([^"]*)"[^>]*>(.*?)</a>', r'[\2](\1)', html_content, flags=re.IGNORECASE | re.DOTALL)
        
        # 转换图片
        html_content = re.sub(
            r'<img[^>]*src="([^"]*)"[^>]*(?:alt="([^"]*)")?[^>]*>',
            lambda m: f'![{m.group(2) if m.group(2) else "图片"}]({m.group(1)})\n\n',
            html_content,
            flags=re.IGNORECASE
        )
        
        # 转换列表
        def convert_ul(match):
            items = re.findall(r'<li[^>]*>(.*?)</li>', match.group(1), re.DOTALL | re.IGNORECASE)
            result = []
            for item in items:
                item_text = re.sub(r'<[^>]+>', '', item.strip())
                if item_text:
                    result.append(f"- {item_text}")
            return '\n'.join(result) + '\n\n' if result else ''
        
        html_content = re.sub(r'<ul[^>]*>(.*?)</ul>', convert_ul, html_content, flags=re.DOTALL | re.IGNORECASE)
        
        def convert_ol(match):
            items = re.findall(r'<li[^>]*>(.*?)</li>', match.group(1), re.DOTALL | re.IGNORECASE)
            result = []
            for i, item in enumerate(items, 1):
                item_text = re.sub(r'<[^>]+>', '', item.strip())
                if item_text:
                    result.append(f"{i}. {item_text}")
            return '\n'.join(result) + '\n\n' if result else ''
        
        html_content = re.sub(r'<ol[^>]*>(.*?)</ol>', convert_ol, html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # 转换表格
        def convert_table(match):
            table_html = match.group(0)
            rows = re.findall(r'<tr[^>]*>(.*?)</tr>', table_html, re.DOTALL | re.IGNORECASE)
            
            if not rows:
                return ''
            
            markdown_table = []
            for i, row in enumerate(rows):
                cells = re.findall(r'<t[hd][^>]*>(.*?)</t[hd]>', row, re.DOTALL | re.IGNORECASE)
                processed_cells = []
                
                for cell in cells:
                    cell_text = re.sub(r'<[^>]+>', ' ', cell.strip())
                    cell_text = re.sub(r'\s+', ' ', cell_text).strip()
                    cell_text = cell_text.replace('|', '\\|')
                    processed_cells.append(cell_text)
                
                if processed_cells:
                    markdown_table.append('| ' + ' | '.join(processed_cells) + ' |')
                    
                    if i == 0:
                        separator = '| ' + ' | '.join(['---'] * len(processed_cells)) + ' |'
                        markdown_table.append(separator)
            
            return '\n'.join(markdown_table) + '\n\n' if markdown_table else ''
        
        html_content = re.sub(r'<table[^>]*>.*?</table>', convert_table, html_content, flags=re.DOTALL | re.IGNORECASE)
        
        # 清理
        html_content = re.sub(r'<br[^>]*>', '\n', html_content, flags=re.IGNORECASE)
        html_content = re.sub(r'<[^>]+>', '', html_content)
        html_content = html.unescape(html_content)
        html_content = re.sub(r'\n{3,}', '\n\n', html_content)
        
        return html_content.strip()
    
    def convert(self, input_path: str, output_path: str, use_pandoc: bool = True) -> Tuple[str, List[str]]:
        """主转换方法"""
        if not os.path.exists(input_path):
            raise FileNotFoundError(f"输入文件不存在: {input_path}")
        
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        
        if use_pandoc and self.pandoc_available:
            try:
                return self.convert_with_pandoc(input_path, output_path)
            except Exception as e:
                print(f"Pandoc转换失败，切换到Mammoth: {e}")
                return self.convert_with_mammoth(input_path, output_path)
        else:
            return self.convert_with_mammoth(input_path, output_path)


class Word2MDConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Word2MD - Word转Markdown转换器")
        self.root.geometry("1000x750")
        
        # 设置图标
        try:
            self.root.iconbitmap("word2md_icon.ico")
        except:
            pass
        
        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.use_pandoc_var = tk.BooleanVar(value=True)
        self.batch_mode_var = tk.BooleanVar(value=False)
        self.batch_files = []
        
        self.converter = EnhancedWordToMarkdownConverter()
        self.converted_content = None
        
        self.create_widgets()
        
    def create_widgets(self):
        """创建界面组件"""
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)
        
        # 转换模式选择
        mode_frame = ttk.LabelFrame(main_frame, text="转换模式", padding="10")
        mode_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Radiobutton(
            mode_frame, text="单文件转换", variable=self.batch_mode_var, 
            value=False, command=self.toggle_mode
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        ttk.Radiobutton(
            mode_frame, text="批量转换", variable=self.batch_mode_var, 
            value=True, command=self.toggle_mode
        ).pack(side=tk.LEFT)
        
        # 文件选择框架
        self.file_frame = ttk.LabelFrame(main_frame, text="文件选择", padding="10")
        self.file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.file_frame.columnconfigure(1, weight=1)
        
        # 单文件模式组件
        self.single_frame = ttk.Frame(self.file_frame)
        self.single_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
        self.single_frame.columnconfigure(1, weight=1)
        
        ttk.Label(self.single_frame, text="Word文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_entry = ttk.Entry(self.single_frame, textvariable=self.input_path_var, width=60)
        self.input_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 5))
        ttk.Button(self.single_frame, text="浏览...", command=self.browse_input_file).grid(row=0, column=2)
        
        ttk.Label(self.single_frame, text="输出文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.output_entry = ttk.Entry(self.single_frame, textvariable=self.output_path_var, width=60)
        self.output_entry.grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 5))
        ttk.Button(self.single_frame, text="选择...", command=self.browse_output_file).grid(row=1, column=2)
        
        # 批量模式组件
        self.batch_frame = ttk.Frame(self.file_frame)
        self.batch_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.batch_frame.columnconfigure(0, weight=1)
        self.batch_frame.rowconfigure(1, weight=1)
        
        # 批量文件选择按钮
        batch_btn_frame = ttk.Frame(self.batch_frame)
        batch_btn_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(batch_btn_frame, text="添加文件", command=self.add_batch_files).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(batch_btn_frame, text="添加文件夹", command=self.add_batch_folder).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(batch_btn_frame, text="清空列表", command=self.clear_batch_files).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Label(batch_btn_frame, text="输出目录:").pack(side=tk.LEFT, padx=(20, 5))
        
        self.batch_output_var = tk.StringVar()
        batch_output_entry = ttk.Entry(batch_btn_frame, textvariable=self.batch_output_var, width=30)
        batch_output_entry.pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(batch_btn_frame, text="选择", command=self.browse_batch_output).pack(side=tk.LEFT)
        
        # 批量文件列表
        list_frame = ttk.Frame(self.batch_frame)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # 创建Treeview显示文件列表
        columns = ('文件名', '路径', '状态')
        self.file_tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=8)
        
        for col in columns:
            self.file_tree.heading(col, text=col)
            if col == '文件名':
                self.file_tree.column(col, width=200)
            elif col == '路径':
                self.file_tree.column(col, width=400)
            else:
                self.file_tree.column(col, width=100)
        
        # 滚动条
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.file_tree.yview)
        self.file_tree.configure(yscrollcommand=scrollbar.set)
        
        self.file_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # 右键菜单
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="移除选中文件", command=self.remove_selected_files)
        self.file_tree.bind("<Button-3>", self.show_context_menu)
        
        # 初始隐藏批量模式
        self.batch_frame.grid_remove()
        
        # 选项和按钮
        options_frame = ttk.LabelFrame(main_frame, text="转换选项", padding="10")
        options_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        pandoc_status = "可用" if self.converter.pandoc_available else "不可用"
        ttk.Checkbutton(
            options_frame, 
            text=f"使用Pandoc ({pandoc_status})", 
            variable=self.use_pandoc_var,
            state=tk.NORMAL if self.converter.pandoc_available else tk.DISABLED
        ).pack(side=tk.LEFT, padx=(0, 20))
        
        # 转换按钮
        button_frame = ttk.Frame(options_frame)
        button_frame.pack(side=tk.RIGHT)
        
        self.convert_btn = ttk.Button(
            button_frame, 
            text="开始转换", 
            command=self.start_conversion
        )
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(
            button_frame, 
            text="清空日志", 
            command=self.clear_log
        ).pack(side=tk.LEFT)
        
        # 日志区域
        log_frame = ttk.LabelFrame(main_frame, text="转换日志", padding="5")
        log_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame, 
            wrap=tk.WORD, 
            width=80, 
            height=15,
            font=('Consolas', 9)
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 进度条和状态
        status_frame = ttk.Frame(main_frame)
        status_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)
        status_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(status_frame, mode='determinate')
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        self.status_label = ttk.Label(status_frame, text="就绪")
        self.status_label.grid(row=0, column=1)
        
    def log(self, message):
        """添加日志信息"""
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()
        
    def browse_input_file(self):
        """浏览输入文件"""
        filename = filedialog.askopenfilename(
            title="选择Word文档",
            filetypes=[
                ("Word文档", "*.docx;*.doc"),
                ("DOCX文件", "*.docx"),
                ("DOC文件", "*.doc"),
                ("所有文件", "*.*")
            ]
        )
        if filename:
            self.input_path_var.set(filename)
            output_path = Path(filename).with_suffix('.md')
            self.output_path_var.set(str(output_path))
    
    def toggle_mode(self):
        """切换转换模式"""
        if self.batch_mode_var.get():
            # 批量模式
            self.single_frame.grid_remove()
            self.batch_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S))
            self.file_frame.rowconfigure(1, weight=1)
        else:
            # 单文件模式
            self.batch_frame.grid_remove()
            self.single_frame.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E))
            self.file_frame.rowconfigure(1, weight=0)
    
    def add_batch_files(self):
        """添加批量文件"""
        filenames = filedialog.askopenfilenames(
            title="选择Word文档",
            filetypes=[
                ("Word文档", "*.docx;*.doc"),
                ("DOCX文件", "*.docx"),
                ("DOC文件", "*.doc"),
                ("所有文件", "*.*")
            ]
        )
        
        for filename in filenames:
            if filename not in [f['path'] for f in self.batch_files]:
                file_info = {
                    'name': os.path.basename(filename),
                    'path': filename,
                    'status': '待转换'
                }
                self.batch_files.append(file_info)
                
                # 添加到列表显示
                self.file_tree.insert('', 'end', values=(file_info['name'], file_info['path'], file_info['status']))
        
        self.log(f"已添加 {len(filenames)} 个文件")
    
    def add_batch_folder(self):
        """添加文件夹中的所有Word文件"""
        folder = filedialog.askdirectory(title="选择包含Word文档的文件夹")
        if not folder:
            return
        
        word_extensions = ['.docx', '.doc']
        added_count = 0
        
        for root, dirs, files in os.walk(folder):
            for file in files:
                if any(file.lower().endswith(ext) for ext in word_extensions):
                    filepath = os.path.join(root, file)
                    if filepath not in [f['path'] for f in self.batch_files]:
                        file_info = {
                            'name': file,
                            'path': filepath,
                            'status': '待转换'
                        }
                        self.batch_files.append(file_info)
                        self.file_tree.insert('', 'end', values=(file_info['name'], file_info['path'], file_info['status']))
                        added_count += 1
        
        self.log(f"从文件夹中添加了 {added_count} 个Word文件")
    
    def clear_batch_files(self):
        """清空批量文件列表"""
        self.batch_files.clear()
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)
        self.log("已清空文件列表")
    
    def browse_batch_output(self):
        """选择批量输出目录"""
        folder = filedialog.askdirectory(title="选择输出目录")
        if folder:
            self.batch_output_var.set(folder)
    
    def show_context_menu(self, event):
        """显示右键菜单"""
        try:
            self.context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.context_menu.grab_release()
    
    def remove_selected_files(self):
        """移除选中的文件"""
        selected_items = self.file_tree.selection()
        if not selected_items:
            return
        
        for item in selected_items:
            values = self.file_tree.item(item, 'values')
            if values:
                file_path = values[1]
                # 从列表中移除
                self.batch_files = [f for f in self.batch_files if f['path'] != file_path]
                self.file_tree.delete(item)
        
        self.log(f"已移除 {len(selected_items)} 个文件")
    
    def clear_log(self):
        """清空日志"""
        self.log_text.delete(1.0, tk.END)
            
    def browse_output_file(self):
        """浏览输出文件"""
        filename = filedialog.asksaveasfilename(
            title="保存Markdown文件",
            defaultextension=".md",
            filetypes=[
                ("Markdown文件", "*.md"),
                ("所有文件", "*.*")
            ]
        )
        if filename:
            self.output_path_var.set(filename)
            
    def start_conversion(self):
        """开始转换"""
        if self.batch_mode_var.get():
            self.start_batch_conversion()
        else:
            self.start_single_conversion()
    
    def start_single_conversion(self):
        """开始单文件转换"""
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        
        if not input_path:
            messagebox.showerror("错误", "请选择输入文件！")
            return
            
        if not output_path:
            messagebox.showerror("错误", "请指定输出文件！")
            return
            
        if not os.path.exists(input_path):
            messagebox.showerror("错误", "输入文件不存在！")
            return
            
        self.convert_btn.config(state=tk.DISABLED)
        self.progress_bar.config(mode='indeterminate')
        self.progress_bar.start(10)
        self.status_label.config(text="转换中...")
        
        self.log("开始单文件转换...")
        
        thread = threading.Thread(
            target=self._convert_single_thread,
            args=(input_path, output_path)
        )
        thread.start()
    
    def start_batch_conversion(self):
        """开始批量转换"""
        if not self.batch_files:
            messagebox.showerror("错误", "请先添加要转换的文件！")
            return
        
        output_dir = self.batch_output_var.get().strip()
        if not output_dir:
            messagebox.showerror("错误", "请选择输出目录！")
            return
        
        if not os.path.exists(output_dir):
            try:
                os.makedirs(output_dir)
            except Exception as e:
                messagebox.showerror("错误", f"无法创建输出目录：{e}")
                return
        
        self.convert_btn.config(state=tk.DISABLED)
        self.progress_bar.config(mode='determinate')
        self.progress_bar['maximum'] = len(self.batch_files)
        self.progress_bar['value'] = 0
        self.status_label.config(text=f"批量转换: 0/{len(self.batch_files)}")
        
        self.log(f"开始批量转换 {len(self.batch_files)} 个文件...")
        
        thread = threading.Thread(
            target=self._convert_batch_thread,
            args=(output_dir,)
        )
        thread.start()
        
    def _convert_single_thread(self, input_path, output_path):
        """单文件转换线程"""
        try:
            content, messages = self.converter.convert(
                input_path, 
                output_path, 
                use_pandoc=self.use_pandoc_var.get()
            )
            
            self.root.after(0, self._single_conversion_complete, content, messages)
            
        except Exception as e:
            self.root.after(0, self._conversion_error, str(e))
    
    def _convert_batch_thread(self, output_dir):
        """批量转换线程"""
        success_count = 0
        error_count = 0
        
        for i, file_info in enumerate(self.batch_files):
            try:
                input_path = file_info['path']
                filename = Path(input_path).stem
                output_path = os.path.join(output_dir, f"{filename}.md")
                
                # 更新状态
                self.root.after(0, self._update_file_status, i, '转换中')
                self.root.after(0, self._update_progress, i + 1, f"批量转换: {i + 1}/{len(self.batch_files)}")
                
                # 执行转换
                content, messages = self.converter.convert(
                    input_path,
                    output_path,
                    use_pandoc=self.use_pandoc_var.get()
                )
                
                # 更新成功状态
                self.root.after(0, self._update_file_status, i, '转换成功')
                self.root.after(0, self.log, f"✓ {file_info['name']} -> {filename}.md")
                success_count += 1
                
            except Exception as e:
                # 更新失败状态
                self.root.after(0, self._update_file_status, i, '转换失败')
                self.root.after(0, self.log, f"✗ {file_info['name']}: {str(e)}")
                error_count += 1
        
        # 批量转换完成
        self.root.after(0, self._batch_conversion_complete, success_count, error_count, output_dir)
            
    def _update_file_status(self, file_index, status):
        """更新文件状态"""
        if file_index < len(self.batch_files):
            self.batch_files[file_index]['status'] = status
            
            # 更新Treeview显示
            items = self.file_tree.get_children()
            if file_index < len(items):
                item = items[file_index]
                values = list(self.file_tree.item(item, 'values'))
                values[2] = status
                self.file_tree.item(item, values=values)
    
    def _update_progress(self, current, status_text):
        """更新进度条"""
        self.progress_bar['value'] = current
        self.status_label.config(text=status_text)
    
    def _single_conversion_complete(self, content, messages):
        """单文件转换完成"""
        self.progress_bar.stop()
        self.convert_btn.config(state=tk.NORMAL)
        self.status_label.config(text="转换完成")
        
        self.converted_content = content
        
        self.log("转换完成！")
        if messages:
            self.log(f"警告信息: {len(messages)} 条")
            for msg in messages[:5]:  # 只显示前5条
                self.log(f"  - {msg}")
        
        messagebox.showinfo("成功", "Word文档已成功转换为Markdown格式！")
        
        # 询问是否打开文件夹
        if messagebox.askyesno("打开文件夹", "是否打开输出文件所在的文件夹？"):
            output_dir = os.path.dirname(self.output_path_var.get())
            try:
                os.startfile(output_dir)
            except:
                subprocess.Popen(['explorer', output_dir])
    
    def _batch_conversion_complete(self, success_count, error_count, output_dir):
        """批量转换完成"""
        self.convert_btn.config(state=tk.NORMAL)
        self.status_label.config(text=f"批量转换完成: 成功{success_count}个, 失败{error_count}个")
        
        total = success_count + error_count
        self.log(f"批量转换完成！成功: {success_count}/{total}, 失败: {error_count}/{total}")
        
        if success_count > 0:
            messagebox.showinfo(
                "批量转换完成", 
                f"批量转换完成！\n成功: {success_count} 个\n失败: {error_count} 个\n\n输出目录: {output_dir}"
            )
            
            # 询问是否打开输出目录
            if messagebox.askyesno("打开文件夹", "是否打开输出目录？"):
                try:
                    os.startfile(output_dir)
                except:
                    subprocess.Popen(['explorer', output_dir])
        else:
            messagebox.showerror("转换失败", "所有文件转换失败，请检查日志信息。")
        
    def _conversion_error(self, error_msg):
        """转换错误"""
        self.progress_bar.stop()
        self.convert_btn.config(state=tk.NORMAL)
        self.status_label.config(text="转换失败")
        
        self.log(f"转换失败: {error_msg}")
        messagebox.showerror("转换失败", f"转换过程中出现错误：\n{error_msg}")


def main():
    root = tk.Tk()
    app = Word2MDConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
