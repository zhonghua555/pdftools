import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import threading
import tempfile
from pypdf import PdfWriter
from PIL import Image, ImageTk
import fitz  # PyMuPDF


class ThumbnailViewer:
    """通用缩略图查看器，支持懒加载、grid布局、多选/单选、滚动加载"""
    def __init__(self, parent, doc, page_indices, max_columns=5, thumbnail_size=120,
                 select_mode='single', on_select=None):
        """
        parent: 父容器（通常是Frame）
        doc: fitz.Document 对象
        page_indices: 需要显示的页面索引列表（按顺序）
        max_columns: 每行最大列数
        thumbnail_size: 缩略图最大尺寸（宽高）
        select_mode: 'single' 或 'multiple'
        on_select: 选中状态变化时的回调，参数为 (选中索引列表)
        """
        self.parent = parent
        self.doc = doc
        self.page_indices = page_indices  # 原始文档中的页号列表
        self.max_columns = max_columns
        self.thumb_size = thumbnail_size
        self.select_mode = select_mode
        self.on_select = on_select

        # 存储每页的控件信息
        self.thumb_items = []  # 每个元素: {'frame', 'photo', 'label', 'page_idx', 'selected'}
        self.selected_indices = set()  # 存储选中的显示顺序索引

        # 懒加载控制
        self.loaded_count = 0          # 已加载的缩略图数量
        self.loading = False           # 是否正在加载中
        self.pending_load = 0          # 待加载数量（用于滚动时触发）
        self.batch_size = 20           # 每次加载的页数

        # 创建滚动区域
        self.canvas = tk.Canvas(parent, bg='#f0f0f0', highlightthickness=0)
        self.v_scrollbar = ttk.Scrollbar(parent, orient='vertical', command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.v_scrollbar.set)
        self.v_scrollbar.pack(side='right', fill='y')
        self.canvas.pack(side='left', fill='both', expand=True)

        # 内部容器
        self.inner_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor='nw')
        self.inner_frame.bind('<Configure>', self._on_inner_configure)

        # 绑定滚动事件
        self.canvas.bind('<Configure>', self._on_canvas_configure)
        self.canvas.bind('<MouseWheel>', self._on_mousewheel)

        # 绑定滚动条事件以实现懒加载
        self.canvas.bind('<Button-1>', lambda e: None)  # 捕获焦点
        self.canvas.bind('<ButtonRelease-1>', self._check_scroll)
        self.v_scrollbar.bind('<ButtonRelease-1>', self._check_scroll)

        # 初始化第一屏
        self.load_more()

    def _on_inner_configure(self, event):
        """内部Frame大小变化时更新Canvas滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))

    def _on_canvas_configure(self, event):
        """Canvas大小变化时重新布局内部子控件"""
        self._relayout()

    def _relayout(self):
        """重新计算grid布局（窗口大小改变时调用）"""
        if not self.thumb_items:
            return
        canvas_width = self.canvas.winfo_width()
        if canvas_width <= 1:
            return
        # 计算每行最大列数
        thumb_width = self.thumb_size + 10  # 加间距
        cols = max(1, canvas_width // thumb_width)
        for idx, item in enumerate(self.thumb_items):
            row = idx // cols
            col = idx % cols
            item['frame'].grid(row=row, column=col, padx=5, pady=5, sticky='nw')

    def _on_mousewheel(self, event):
        """鼠标滚轮滚动，触发懒加载检查"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        self._check_scroll()

    def _check_scroll(self, event=None):
        """检查滚动位置，接近底部时加载更多"""
        if self.loading:
            return
        # 获取可视区域和总高度
        bbox = self.canvas.bbox('all')
        if not bbox:
            return
        total_height = bbox[3] - bbox[1]
        visible_top = self.canvas.canvasy(0)
        visible_bottom = visible_top + self.canvas.winfo_height()
        # 当可视区域底部超过总高度的80%时加载更多
        if visible_bottom >= 0.8 * total_height:
            self.load_more()

    def load_more(self):
        """加载下一批缩略图"""
        if self.loading:
            return
        remaining = len(self.page_indices) - self.loaded_count
        if remaining <= 0:
            return
        to_load = min(self.batch_size, remaining)
        if to_load == 0:
            return

        self.loading = True
        # 启动后台线程生成缩略图
        threading.Thread(target=self._generate_thumbnails_batch, args=(to_load,), daemon=True).start()

    def _generate_thumbnails_batch(self, count):
        """在后台线程中生成一批缩略图，每生成一个就通过after添加到UI"""
        start = self.loaded_count
        end = min(start + count, len(self.page_indices))
        for i in range(start, end):
            page_idx = self.page_indices[i]
            try:
                # 生成缩略图
                img = self._render_thumbnail(page_idx)
                photo = ImageTk.PhotoImage(img)
                # 在主线程中添加UI
                self.parent.after(0, self._add_thumbnail, i, page_idx, photo)
            except Exception as e:
                print(f"生成缩略图失败 page {page_idx}: {e}")
        # 所有缩略图生成完成
        self.parent.after(0, self._loading_finished)

    def _render_thumbnail(self, page_idx):
        """渲染单页缩略图"""
        page = self.doc.load_page(page_idx)
        zoom = self.thumb_size / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        img.thumbnail((self.thumb_size, self.thumb_size), Image.Resampling.LANCZOS)
        return img

    def _add_thumbnail(self, order_idx, page_idx, photo):
        """添加一个缩略图到UI（在主线程中执行）"""
        # 创建Frame
        frame = ttk.Frame(self.inner_frame, relief='ridge', borderwidth=1)
        # 图片标签
        img_label = ttk.Label(frame, image=photo)
        img_label.image = photo
        img_label.pack()
        # 文本标签
        info_label = ttk.Label(frame, text=f"第 {page_idx+1} 页")
        info_label.pack()

        # 绑定点击事件
        def on_click(e, idx=order_idx):
            self._toggle_selection(idx)
        frame.bind('<Button-1>', on_click)
        img_label.bind('<Button-1>', on_click)
        info_label.bind('<Button-1>', on_click)

        # 保存项目信息
        self.thumb_items.append({
            'frame': frame,
            'photo': photo,
            'page_idx': page_idx,
            'selected': False
        })

        # 重新布局
        self._relayout()
        # 更新滚动区域
        self.inner_frame.update_idletasks()
        self.canvas.configure(scrollregion=self.canvas.bbox('all'))
        # 记录已加载数量
        self.loaded_count = order_idx + 1

    def _loading_finished(self):
        """一批缩略图加载完成"""
        self.loading = False
        # 如果还没加载完所有页，但滚动到底部会再次触发
        if self.loaded_count < len(self.page_indices):
            self._check_scroll()

    def _toggle_selection(self, idx):
        """切换选中状态"""
        if self.select_mode == 'single':
            # 清除其他选中
            for i, item in enumerate(self.thumb_items):
                if i == idx:
                    new_state = not item['selected']
                    item['selected'] = new_state
                    if new_state:
                        self.selected_indices = {i}
                    else:
                        self.selected_indices.clear()
                else:
                    if item['selected']:
                        item['selected'] = False
                        self._update_frame_style(item['frame'], False)
            if idx in self.selected_indices:
                self._update_frame_style(self.thumb_items[idx]['frame'], True)
            else:
                self._update_frame_style(self.thumb_items[idx]['frame'], False)
        else:  # multiple
            item = self.thumb_items[idx]
            new_state = not item['selected']
            item['selected'] = new_state
            if new_state:
                self.selected_indices.add(idx)
            else:
                self.selected_indices.discard(idx)
            self._update_frame_style(item['frame'], new_state)

        if self.on_select:
            # 回调传递选中的原始页面索引列表
            selected_page_indices = [self.thumb_items[i]['page_idx'] for i in sorted(self.selected_indices)]
            self.on_select(selected_page_indices)

    def _update_frame_style(self, frame, selected):
        """更新Frame的样式表示选中"""
        if selected:
            frame.config(style='Selected.TFrame')
        else:
            frame.config(style='TFrame')

    def get_selected_page_indices(self):
        """获取当前选中的原始页面索引列表"""
        return [self.thumb_items[i]['page_idx'] for i in sorted(self.selected_indices)]

    def select_all(self):
        """全选（仅多选模式有效）"""
        if self.select_mode != 'multiple':
            return
        for i, item in enumerate(self.thumb_items):
            if not item['selected']:
                item['selected'] = True
                self.selected_indices.add(i)
                self._update_frame_style(item['frame'], True)
        if self.on_select:
            self.on_select(self.get_selected_page_indices())

    def select_none(self):
        """取消全选"""
        for i, item in enumerate(self.thumb_items):
            if item['selected']:
                item['selected'] = False
                self.selected_indices.discard(i)
                self._update_frame_style(item['frame'], False)
        if self.on_select:
            self.on_select([])

    def invert_selection(self):
        """反选（仅多选模式）"""
        if self.select_mode != 'multiple':
            return
        for i, item in enumerate(self.thumb_items):
            item['selected'] = not item['selected']
            if item['selected']:
                self.selected_indices.add(i)
            else:
                self.selected_indices.discard(i)
            self._update_frame_style(item['frame'], item['selected'])
        if self.on_select:
            self.on_select(self.get_selected_page_indices())

    def destroy(self):
        """销毁所有控件"""
        for item in self.thumb_items:
            item['frame'].destroy()
        self.inner_frame.destroy()
        self.canvas.destroy()
        self.v_scrollbar.destroy()


class PDFToolbox:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 工具箱")
        try:
            self.root.iconbitmap('app32.ico')
        except:
            pass
        self.root.geometry("1024x600")
        self.root.resizable(True, True)

        # 样式
        self.style = ttk.Style()
        self.style.theme_use('clam')
        self.style.configure('TNotebook.Tab', padding=[12, 4], font=('Segoe UI', 10))
        self.style.configure('TLabel', font=('Segoe UI', 9))
        self.style.configure('TButton', font=('Segoe UI', 9))
        self.style.configure('TEntry', font=('Segoe UI', 9))
        self.style.configure('Selected.TFrame', background='lightblue')

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=10)

        # 状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = ttk.Label(root, textvariable=self.status_var, relief='sunken', anchor='w')
        status_bar.pack(side='bottom', fill='x', padx=10, pady=5)

        # 创建各功能页面
        self.create_merge_tab()
        self.create_split_tab()
        self.create_pdf2img_tab()
        self.create_img2pdf_tab()
        self.create_rotate_tab()
        self.create_sort_tab()
        self.create_menu()

        # 存储各模块的缩略图查看器，以便在关闭文档时清理
        self.viewers = []

    def create_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="帮助", menu=help_menu)
        help_menu.add_command(label="关于", command=self.show_about)

    def show_about(self):
        about_text = (
            "PDF 工具箱\n\n"
            "版本：2.0\n"
            "功能：合并、拆分、旋转、排序、PDF↔图片\n\n"
            "© 2026 马建旗 24461001@qq.com\n"
            "本工具使用 python 构建。\n"
            "遵循 MIT 协议，欢迎学习交流。"
        )
        messagebox.showinfo("关于 PDF 工具箱", about_text)

    # ---------- 通用方法 ----------
    def get_absolute_path(self, path, create_dir=False):
        if not path:
            return ""
        abs_path = os.path.abspath(path)
        dirname = os.path.dirname(abs_path)
        if create_dir and dirname and not os.path.exists(dirname):
            os.makedirs(dirname, exist_ok=True)
        return abs_path

    def validate_output_path(self, path, create_dir=True):
        abs_path = self.get_absolute_path(path, create_dir)
        if not abs_path:
            raise ValueError("输出路径不能为空")
        return abs_path

    def run_thread(self, target, args=(), on_success=None, on_error=None, on_finally=None, status_msg=None):
        """在后台线程中运行函数，并可选更新状态栏"""
        def wrapper():
            if status_msg:
                self.root.after(0, lambda: self.status_var.set(status_msg))
            try:
                result = target(*args)
                if on_success:
                    self.root.after(0, on_success, result)
            except Exception as e:
                if on_error:
                    self.root.after(0, on_error, e)
                else:
                    self.root.after(0, lambda: messagebox.showerror("错误", str(e)))
            finally:
                if on_finally:
                    self.root.after(0, on_finally)
                if status_msg:
                    self.root.after(0, lambda: self.status_var.set("就绪"))
        threading.Thread(target=wrapper, daemon=True).start()

    def select_file(self, var, title="选择文件", filetypes=[("PDF files", "*.pdf")]):
        f = filedialog.askopenfilename(title=title, filetypes=filetypes)
        if f:
            var.set(f)
            return f
        return None

    def open_file_with_default_app(self, file_path):
        import subprocess
        try:
            subprocess.run(['start', file_path], shell=True)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开文件：{e}")

    def show_loading(self, message="正在加载，请稍候..."):
        self.loading_win = tk.Toplevel(self.root)
        self.loading_win.title("请稍候")
        self.loading_win.geometry("300x100")
        self.loading_win.transient(self.root)
        self.loading_win.grab_set()
        ttk.Label(self.loading_win, text=message).pack(pady=20)
        self.loading_win.update()

    def hide_loading(self):
        if hasattr(self, 'loading_win') and self.loading_win:
            self.loading_win.destroy()
            self.loading_win = None

    # ---------- 1. 合并 PDF ----------
    def create_merge_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="合并 PDF")
        ttk.Label(frame, text="待合并 PDF 文件（按顺序）:").pack(anchor='w', padx=5, pady=5)

        # 使用简化的列表方式，不使用缩略图懒加载（因为合并的文件数量通常较少）
        self.merge_files = []  # 存储文件路径
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.merge_listbox = tk.Listbox(list_frame)
        scroll = ttk.Scrollbar(list_frame, orient='vertical', command=self.merge_listbox.yview)
        self.merge_listbox.configure(yscrollcommand=scroll.set)
        self.merge_listbox.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame, text="添加 PDF", command=self.add_merge_file).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_merge_file).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="上移", command=self.move_merge_up).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="下移", command=self.move_merge_down).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="清空列表", command=self.clear_merge_list).pack(side='left', padx=5)

        self.merge_btn = ttk.Button(frame, text="开始合并", command=self.start_merge)
        self.merge_btn.pack(pady=10)

    def add_merge_file(self):
        files = filedialog.askopenfilenames(title="选择 PDF 文件", filetypes=[("PDF files", "*.pdf")])
        for f in files:
            if f not in self.merge_files:
                self.merge_files.append(f)
                self.merge_listbox.insert(tk.END, f"{os.path.basename(f)}  ({self.get_pdf_page_count(f)}页)")

    def get_pdf_page_count(self, pdf_path):
        try:
            doc = fitz.open(pdf_path)
            count = len(doc)
            doc.close()
            return count
        except:
            return 0

    def remove_merge_file(self):
        sel = self.merge_listbox.curselection()
        if sel:
            idx = sel[0]
            self.merge_listbox.delete(idx)
            del self.merge_files[idx]

    def move_merge_up(self):
        sel = self.merge_listbox.curselection()
        if sel and sel[0] > 0:
            idx = sel[0]
            self.merge_files[idx], self.merge_files[idx-1] = self.merge_files[idx-1], self.merge_files[idx]
            self.refresh_merge_listbox()
            self.merge_listbox.selection_set(idx-1)

    def move_merge_down(self):
        sel = self.merge_listbox.curselection()
        if sel and sel[0] < len(self.merge_files)-1:
            idx = sel[0]
            self.merge_files[idx], self.merge_files[idx+1] = self.merge_files[idx+1], self.merge_files[idx]
            self.refresh_merge_listbox()
            self.merge_listbox.selection_set(idx+1)

    def refresh_merge_listbox(self):
        self.merge_listbox.delete(0, tk.END)
        for f in self.merge_files:
            self.merge_listbox.insert(tk.END, f"{os.path.basename(f)}  ({self.get_pdf_page_count(f)}页)")

    def clear_merge_list(self):
        self.merge_files.clear()
        self.merge_listbox.delete(0, tk.END)

    def start_merge(self):
        if len(self.merge_files) < 2:
            messagebox.showerror("错误", "请至少添加两个 PDF 文件")
            return
        out = filedialog.asksaveasfilename(title="保存合并后的PDF", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not out:
            return
        try:
            out_abs = self.validate_output_path(out, create_dir=True)
        except Exception as e:
            messagebox.showerror("错误", f"输出路径无效：{e}")
            return

        self.merge_btn.config(state='disabled', text='合并中...')
        self.run_thread(
            target=self.merge_pdfs,
            args=(self.merge_files, out_abs),
            on_success=lambda r: messagebox.showinfo("完成", f"合并成功！\n输出文件：{r}"),
            on_error=lambda e: messagebox.showerror("错误", f"合并失败：{e}"),
            on_finally=lambda: self.merge_btn.config(state='normal', text='开始合并'),
            status_msg="正在合并 PDF，请稍候..."
        )

    @staticmethod
    def merge_pdfs(files, out):
        writer = PdfWriter()
        for f in files:
            writer.append(f)
        with open(out, 'wb') as f:
            writer.write(f)
        return out

    # ---------- 2. 拆分/删除 ----------
    def create_split_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="拆分/删除")

        top_frame = ttk.Frame(frame)
        top_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(top_frame, text="选择 PDF 文件:").pack(side='left')
        self.split_file = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.split_file, width=50).pack(side='left', padx=5)
        ttk.Button(top_frame, text="浏览", command=self.load_split_file).pack(side='left')

        self.split_page_count = tk.StringVar(value="未选择文件")
        ttk.Label(top_frame, textvariable=self.split_page_count, foreground="gray").pack(side='left', padx=10)

        # 缩略图区域容器
        self.split_viewer_frame = ttk.Frame(frame)
        self.split_viewer_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.split_viewer = None
        self.split_doc = None

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame, text="全选", command=self.split_select_all).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="取消全选", command=self.split_select_none).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="删除选中页面", command=self.delete_selected_pages).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="保存原文件", command=self.save_original_split).pack(side='left', padx=5)

    def load_split_file(self):
        file = self.select_file(self.split_file, "选择 PDF 文件")
        if not file:
            return
        if self.split_doc:
            self.split_doc.close()
            if self.split_viewer:
                self.split_viewer.destroy()
        try:
            self.show_loading("正在打开PDF并生成缩略图...")
            self.split_doc = fitz.open(file)
            total = len(self.split_doc)
            self.split_page_count.set(f"总页数: {total}")
            # 创建缩略图查看器（多选模式）
            self.split_viewer = ThumbnailViewer(
                parent=self.split_viewer_frame,
                doc=self.split_doc,
                page_indices=list(range(total)),
                max_columns=5,
                thumbnail_size=120,
                select_mode='multiple'
            )
            self.hide_loading()
        except Exception as e:
            self.hide_loading()
            messagebox.showerror("错误", f"无法打开 PDF：{e}")

    def split_select_all(self):
        if self.split_viewer:
            self.split_viewer.select_all()

    def split_select_none(self):
        if self.split_viewer:
            self.split_viewer.select_none()

    def delete_selected_pages(self):
        if not self.split_doc or not self.split_viewer:
            messagebox.showerror("错误", "未加载 PDF")
            return
        selected = self.split_viewer.get_selected_page_indices()
        if not selected:
            messagebox.showerror("错误", "请至少选中一个页面")
            return
        if not messagebox.askyesno("确认删除", f"即将删除 {len(selected)} 页，是否继续？"):
            return
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output:
            return
        try:
            out_abs = self.validate_output_path(output, create_dir=True)
            self.run_thread(
                target=self.save_without_pages,
                args=(self.split_doc, selected, out_abs),
                on_success=lambda r: messagebox.showinfo("完成", f"已删除 {len(selected)} 页，保存到：{r}"),
                on_error=lambda e: messagebox.showerror("错误", f"保存失败：{e}"),
                status_msg="正在删除页面并保存..."
            )
        except Exception as e:
            messagebox.showerror("错误", f"输出路径无效：{e}")

    def save_without_pages(self, doc, pages_to_remove, output_path):
        new_doc = fitz.open()
        total = len(doc)
        for i in range(total):
            if i not in pages_to_remove:
                new_doc.insert_pdf(doc, from_page=i, to_page=i)
        new_doc.save(output_path)
        new_doc.close()
        return output_path

    def save_original_split(self):
        if not self.split_doc:
            messagebox.showerror("错误", "未加载 PDF")
            return
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output:
            return
        try:
            out_abs = self.validate_output_path(output, create_dir=True)
            self.run_thread(
                target=self.save_copy,
                args=(self.split_doc, out_abs),
                on_success=lambda r: messagebox.showinfo("完成", f"原文件已保存到：{r}"),
                on_error=lambda e: messagebox.showerror("错误", f"保存失败：{e}"),
                status_msg="正在保存副本..."
            )
        except Exception as e:
            messagebox.showerror("错误", f"输出路径无效：{e}")

    def save_copy(self, doc, output_path):
        new_doc = fitz.open()
        new_doc.insert_pdf(doc)
        new_doc.save(output_path)
        new_doc.close()
        return output_path

    # ---------- 3. PDF 转图片 ----------
    def create_pdf2img_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="PDF 转图片")

        top_frame = ttk.Frame(frame)
        top_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(top_frame, text="选择 PDF 文件:").pack(side='left')
        self.pdf2img_file = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.pdf2img_file, width=50).pack(side='left', padx=5)
        ttk.Button(top_frame, text="浏览", command=self.load_pdf2img_file).pack(side='left')

        self.pdf2img_page_count = tk.StringVar(value="未选择文件")
        ttk.Label(top_frame, textvariable=self.pdf2img_page_count, foreground="gray").pack(side='left', padx=10)

        self.pdf2img_viewer_frame = ttk.Frame(frame)
        self.pdf2img_viewer_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.pdf2img_viewer = None
        self.pdf2img_doc = None

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame, text="全选", command=self.pdf2img_select_all).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="取消全选", command=self.pdf2img_select_none).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="反选", command=self.pdf2img_invert_selection).pack(side='left', padx=5)
        self.pdf2img_btn = ttk.Button(btn_frame, text="导出页面", command=self.export_pages_with_settings)
        self.pdf2img_btn.pack(side='left', padx=20)

    def load_pdf2img_file(self):
        file = self.select_file(self.pdf2img_file, "选择 PDF 文件")
        if not file:
            return
        if self.pdf2img_doc:
            self.pdf2img_doc.close()
            if self.pdf2img_viewer:
                self.pdf2img_viewer.destroy()
        try:
            self.show_loading("正在打开PDF并生成缩略图...")
            self.pdf2img_doc = fitz.open(file)
            total = len(self.pdf2img_doc)
            self.pdf2img_page_count.set(f"总页数: {total}")
            self.pdf2img_viewer = ThumbnailViewer(
                parent=self.pdf2img_viewer_frame,
                doc=self.pdf2img_doc,
                page_indices=list(range(total)),
                max_columns=5,
                thumbnail_size=120,
                select_mode='multiple'
            )
            # 默认全选
            self.pdf2img_viewer.select_all()
            self.hide_loading()
        except Exception as e:
            self.hide_loading()
            messagebox.showerror("错误", f"无法打开 PDF：{e}")

    def pdf2img_select_all(self):
        if self.pdf2img_viewer:
            self.pdf2img_viewer.select_all()

    def pdf2img_select_none(self):
        if self.pdf2img_viewer:
            self.pdf2img_viewer.select_none()

    def pdf2img_invert_selection(self):
        if self.pdf2img_viewer:
            self.pdf2img_viewer.invert_selection()

    def export_pages_with_settings(self):
        if not self.pdf2img_doc or not self.pdf2img_viewer:
            messagebox.showerror("错误", "未加载 PDF")
            return
        selected = self.pdf2img_viewer.get_selected_page_indices()
        if not selected:
            messagebox.showerror("错误", "请至少选中一个页面")
            return

        # 弹出设置对话框
        settings_win = tk.Toplevel(self.root)
        settings_win.title("导出设置")
        settings_win.geometry("300x200")
        settings_win.resizable(False, False)
        settings_win.transient(self.root)
        settings_win.grab_set()

        ttk.Label(settings_win, text="图片格式:").pack(pady=5)
        fmt_var = tk.StringVar(value="png")
        fmt_combo = ttk.Combobox(settings_win, textvariable=fmt_var, values=['png', 'jpg', 'jpeg'], state='readonly')
        fmt_combo.pack()

        ttk.Label(settings_win, text="DPI (72-600):").pack(pady=5)
        dpi_var = tk.StringVar(value="200")
        dpi_entry = ttk.Entry(settings_win, textvariable=dpi_var, width=10)
        dpi_entry.pack()

        def on_confirm():
            fmt = fmt_var.get().lower()
            try:
                dpi_val = int(dpi_var.get())
                if dpi_val < 72 or dpi_val > 600:
                    raise ValueError
            except:
                messagebox.showerror("错误", "DPI 必须是 72-600 之间的整数")
                return
            settings_win.destroy()
            output_dir = filedialog.askdirectory(title="选择输出文件夹（可新建文件夹）")
            if not output_dir:
                return
            try:
                abs_dir = self.get_absolute_path(output_dir, create_dir=True)
            except Exception as e:
                messagebox.showerror("错误", f"输出文件夹无效：{e}")
                return
            self.pdf2img_btn.config(state='disabled', text='导出中...')
            self.run_thread(
                target=self.convert_selected_pages_to_images,
                args=(self.pdf2img_doc, selected, abs_dir, fmt, dpi_val),
                on_success=lambda c: messagebox.showinfo("完成", f"导出成功！共 {c} 张图片"),
                on_error=lambda e: messagebox.showerror("错误", f"导出失败：{e}"),
                on_finally=lambda: self.pdf2img_btn.config(state='normal', text='导出页面'),
                status_msg="正在导出图片，请稍候..."
            )

        ttk.Button(settings_win, text="确定", command=on_confirm).pack(pady=10)

    def convert_selected_pages_to_images(self, doc, page_indices, output_dir, fmt, dpi):
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        os.makedirs(output_dir, exist_ok=True)
        for i, page_idx in enumerate(page_indices):
            page = doc.load_page(page_idx)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            ext = fmt if fmt != 'jpg' else 'jpeg'
            save_path = os.path.join(output_dir, f'page_{page_idx+1:03d}.{fmt}')
            if fmt == 'png':
                pix.save(save_path)
            else:
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                img.save(save_path, quality=95)
        return len(page_indices)

    # ---------- 4. 图片转 PDF ----------
    def create_img2pdf_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="图片转 PDF")

        ttk.Label(frame, text="图片文件列表（按顺序）:").pack(anchor='w', padx=5, pady=5)

        # 使用Listbox简化，因为图片数量通常不多
        self.img_files = []
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.img_listbox = tk.Listbox(list_frame)
        scroll = ttk.Scrollbar(list_frame, orient='vertical', command=self.img_listbox.yview)
        self.img_listbox.configure(yscrollcommand=scroll.set)
        self.img_listbox.pack(side='left', fill='both', expand=True)
        scroll.pack(side='right', fill='y')

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame, text="添加图片", command=self.add_images).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="移除选中", command=self.remove_img).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="上移", command=self.move_img_up).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="下移", command=self.move_img_down).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="置顶", command=self.move_img_top).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="置底", command=self.move_img_bottom).pack(side='left', padx=5)

        self.img2pdf_btn = ttk.Button(frame, text="开始转换", command=self.start_img2pdf)
        self.img2pdf_btn.pack(pady=10)

    def add_images(self):
        files = filedialog.askopenfilenames(title="选择图片", filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp")])
        for f in files:
            if f not in self.img_files:
                self.img_files.append(f)
                self.img_listbox.insert(tk.END, os.path.basename(f))

    def remove_img(self):
        sel = self.img_listbox.curselection()
        if sel:
            idx = sel[0]
            self.img_listbox.delete(idx)
            del self.img_files[idx]

    def move_img_up(self):
        sel = self.img_listbox.curselection()
        if sel and sel[0] > 0:
            idx = sel[0]
            self.img_files[idx], self.img_files[idx-1] = self.img_files[idx-1], self.img_files[idx]
            self.refresh_img_listbox()
            self.img_listbox.selection_set(idx-1)

    def move_img_down(self):
        sel = self.img_listbox.curselection()
        if sel and sel[0] < len(self.img_files)-1:
            idx = sel[0]
            self.img_files[idx], self.img_files[idx+1] = self.img_files[idx+1], self.img_files[idx]
            self.refresh_img_listbox()
            self.img_listbox.selection_set(idx+1)

    def move_img_top(self):
        sel = self.img_listbox.curselection()
        if sel and sel[0] > 0:
            idx = sel[0]
            item = self.img_files.pop(idx)
            self.img_files.insert(0, item)
            self.refresh_img_listbox()
            self.img_listbox.selection_set(0)

    def move_img_bottom(self):
        sel = self.img_listbox.curselection()
        if sel and sel[0] < len(self.img_files)-1:
            idx = sel[0]
            item = self.img_files.pop(idx)
            self.img_files.append(item)
            self.refresh_img_listbox()
            self.img_listbox.selection_set(len(self.img_files)-1)

    def refresh_img_listbox(self):
        self.img_listbox.delete(0, tk.END)
        for f in self.img_files:
            self.img_listbox.insert(tk.END, os.path.basename(f))

    def start_img2pdf(self):
        if not self.img_files:
            messagebox.showerror("错误", "请至少添加一张图片")
            return
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output:
            return
        try:
            out_abs = self.validate_output_path(output, create_dir=True)
        except Exception as e:
            messagebox.showerror("错误", f"输出路径无效：{e}")
            return

        self.img2pdf_btn.config(state='disabled', text='转换中...')
        self.run_thread(
            target=self.convert_images_to_pdf,
            args=(self.img_files, out_abs),
            on_success=lambda c: messagebox.showinfo("完成", f"转换成功！共 {c} 张图片"),
            on_error=lambda e: messagebox.showerror("错误", f"转换失败：{e}"),
            on_finally=lambda: self.img2pdf_btn.config(state='normal', text='开始转换'),
            status_msg="正在转换图片为 PDF..."
        )

    @staticmethod
    def convert_images_to_pdf(images, output):
        writer = PdfWriter()
        for img_path in images:
            img = Image.open(img_path)
            if img.mode != 'RGB':
                img = img.convert('RGB')
            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as tmp:
                img.save(tmp, format='PDF')
                tmp_path = tmp.name
            writer.append(tmp_path)
            os.unlink(tmp_path)
        with open(output, 'wb') as f:
            writer.write(f)
        return len(images)

    # ---------- 5. 页面旋转 ----------
    def create_rotate_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="页面旋转")

        top_frame = ttk.Frame(frame)
        top_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(top_frame, text="选择 PDF 文件:").pack(side='left')
        self.rotate_file = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.rotate_file, width=50).pack(side='left', padx=5)
        ttk.Button(top_frame, text="浏览", command=self.load_rotate_file).pack(side='left')

        self.rotate_page_count = tk.StringVar(value="未选择文件")
        ttk.Label(top_frame, textvariable=self.rotate_page_count, foreground="gray").pack(side='left', padx=10)

        control_frame = ttk.Frame(frame)
        control_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(control_frame, text="旋转角度:").pack(side='left')
        self.rotate_angle = tk.StringVar(value="90")
        ttk.Radiobutton(control_frame, text="顺时针 90°", variable=self.rotate_angle, value="90").pack(side='left', padx=5)
        ttk.Radiobutton(control_frame, text="逆时针 90°", variable=self.rotate_angle, value="270").pack(side='left', padx=5)
        ttk.Radiobutton(control_frame, text="180°", variable=self.rotate_angle, value="180").pack(side='left', padx=5)
        ttk.Button(control_frame, text="预览旋转", command=self.preview_rotate).pack(side='left', padx=10)
        ttk.Button(control_frame, text="重置旋转", command=self.reset_rotate).pack(side='left', padx=5)
        ttk.Button(control_frame, text="保存旋转", command=self.save_rotate).pack(side='left', padx=5)

        self.rotate_viewer_frame = ttk.Frame(frame)
        self.rotate_viewer_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.rotate_viewer = None
        self.rotate_doc = None
        self.original_rotations = {}  # 原始旋转角度

    def load_rotate_file(self):
        file = self.select_file(self.rotate_file, "选择 PDF 文件")
        if not file:
            return
        if self.rotate_doc:
            self.rotate_doc.close()
            if self.rotate_viewer:
                self.rotate_viewer.destroy()
        try:
            self.show_loading("正在打开PDF并生成缩略图...")
            self.rotate_doc = fitz.open(file)
            total = len(self.rotate_doc)
            self.rotate_page_count.set(f"总页数: {total}")
            # 保存原始旋转
            for i in range(total):
                self.original_rotations[i] = self.rotate_doc.load_page(i).rotation
            self.rotate_viewer = ThumbnailViewer(
                parent=self.rotate_viewer_frame,
                doc=self.rotate_doc,
                page_indices=list(range(total)),
                max_columns=5,
                thumbnail_size=120,
                select_mode='multiple'
            )
            self.hide_loading()
        except Exception as e:
            self.hide_loading()
            messagebox.showerror("错误", f"无法打开 PDF：{e}")

    def preview_rotate(self):
        if not self.rotate_doc or not self.rotate_viewer:
            messagebox.showerror("错误", "未加载 PDF")
            return
        selected = self.rotate_viewer.get_selected_page_indices()
        if not selected:
            messagebox.showinfo("提示", "请先选择要旋转的页面")
            return
        angle = int(self.rotate_angle.get())
        # 创建临时文档预览
        preview_doc = fitz.open()
        for i in range(len(self.rotate_doc)):
            page = self.rotate_doc.load_page(i)
            if i in selected:
                page.set_rotation((page.rotation + angle) % 360)
            preview_doc.insert_pdf(self.rotate_doc, from_page=i, to_page=i)
            # 恢复原文档中的旋转
            page.set_rotation(self.original_rotations[i])
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, "rotate_preview.pdf")
        preview_doc.save(temp_path)
        preview_doc.close()
        self.open_file_with_default_app(temp_path)

    def reset_rotate(self):
        if not self.rotate_doc:
            messagebox.showerror("错误", "未加载 PDF")
            return
        for i in range(len(self.rotate_doc)):
            page = self.rotate_doc.load_page(i)
            page.set_rotation(self.original_rotations[i])
        messagebox.showinfo("提示", "已重置所有旋转操作")
        # 刷新缩略图（需要重新生成，因为旋转角度影响显示）
        if self.rotate_viewer:
            self.rotate_viewer.destroy()
        self.rotate_viewer = ThumbnailViewer(
            parent=self.rotate_viewer_frame,
            doc=self.rotate_doc,
            page_indices=list(range(len(self.rotate_doc))),
            max_columns=5,
            thumbnail_size=120,
            select_mode='multiple'
        )

    def save_rotate(self):
        if not self.rotate_doc or not self.rotate_viewer:
            messagebox.showerror("错误", "未加载 PDF")
            return
        selected = self.rotate_viewer.get_selected_page_indices()
        if not selected:
            if not messagebox.askyesno("确认", "没有选中任何页面，是否直接保存原文件？"):
                return
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output:
            return
        try:
            out_abs = self.validate_output_path(output, create_dir=True)
            self.run_thread(
                target=self.save_rotated_pdf,
                args=(self.rotate_doc, selected, int(self.rotate_angle.get()), out_abs),
                on_success=lambda r: messagebox.showinfo("完成", f"旋转结果已保存到：{r}"),
                on_error=lambda e: messagebox.showerror("错误", f"保存失败：{e}"),
                status_msg="正在保存旋转后的 PDF..."
            )
        except Exception as e:
            messagebox.showerror("错误", f"输出路径无效：{e}")

    def save_rotated_pdf(self, doc, selected_indices, angle, output_path):
        new_doc = fitz.open()
        for i in range(len(doc)):
            page = doc.load_page(i)
            if i in selected_indices:
                page.set_rotation((page.rotation + angle) % 360)
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
            # 恢复原文档中的旋转
            page.set_rotation(self.original_rotations[i])
        new_doc.save(output_path)
        new_doc.close()
        return output_path

    # ---------- 6. 页面排序 ----------
    def create_sort_tab(self):
        frame = ttk.Frame(self.notebook)
        self.notebook.add(frame, text="页面排序")

        top_frame = ttk.Frame(frame)
        top_frame.pack(fill='x', padx=5, pady=5)
        ttk.Label(top_frame, text="选择 PDF 文件:").pack(side='left')
        self.sort_file = tk.StringVar()
        ttk.Entry(top_frame, textvariable=self.sort_file, width=50).pack(side='left', padx=5)
        ttk.Button(top_frame, text="浏览", command=self.load_sort_file).pack(side='left')

        self.sort_page_count = tk.StringVar(value="未选择文件")
        ttk.Label(top_frame, textvariable=self.sort_page_count, foreground="gray").pack(side='left', padx=10)

        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill='x', padx=5, pady=5)
        ttk.Button(btn_frame, text="上移", command=self.sort_move_up).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="下移", command=self.sort_move_down).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="置顶", command=self.sort_move_top).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="置底", command=self.sort_move_bottom).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="保存排序", command=self.save_sorted_pdf).pack(side='left', padx=20)

        self.sort_viewer_frame = ttk.Frame(frame)
        self.sort_viewer_frame.pack(fill='both', expand=True, padx=5, pady=5)
        self.sort_viewer = None
        self.sort_doc = None
        self.sort_order = []  # 当前排序后的原始页面索引列表

    def load_sort_file(self):
        file = self.select_file(self.sort_file, "选择 PDF 文件")
        if not file:
            return
        if self.sort_doc:
            self.sort_doc.close()
            if self.sort_viewer:
                self.sort_viewer.destroy()
        try:
            self.show_loading("正在打开PDF并生成缩略图...")
            self.sort_doc = fitz.open(file)
            total = len(self.sort_doc)
            self.sort_page_count.set(f"总页数: {total}")
            self.sort_order = list(range(total))
            self.sort_viewer = ThumbnailViewer(
                parent=self.sort_viewer_frame,
                doc=self.sort_doc,
                page_indices=self.sort_order,
                max_columns=5,
                thumbnail_size=120,
                select_mode='single',
                on_select=self.on_sort_select
            )
            self.hide_loading()
        except Exception as e:
            self.hide_loading()
            messagebox.showerror("错误", f"无法打开 PDF：{e}")

    def on_sort_select(self, selected_indices):
        """排序页面选中的回调，记录选中的显示位置索引（原始顺序）"""
        # 由于ThumbnailViewer返回的是原始页面索引，我们需要根据当前sort_order找到显示顺序中的位置
        # 这里不直接使用，因为排序需要知道用户点击的缩略图在显示列表中的索引。
        # 更好的做法：在ThumbnailViewer中增加一个方法获取当前选中的显示顺序索引。
        # 为了简单，我们在排序模块中不使用选中回调，而是直接获取当前选中项。
        pass

    def get_selected_sort_index(self):
        """获取当前选中的缩略图在显示顺序中的索引"""
        if not self.sort_viewer:
            return None
        selected = self.sort_viewer.get_selected_page_indices()
        if not selected:
            return None
        # 因为排序模式下，ThumbnailViewer显示的page_indices是当前sort_order，所以选中的原始页号就是显示列表中的对应页号
        # 但需要找到它在当前sort_order中的位置
        page_idx = selected[0]
        try:
            pos = self.sort_order.index(page_idx)
            return pos
        except ValueError:
            return None

    def sort_move_up(self):
        idx = self.get_selected_sort_index()
        if idx is None:
            messagebox.showinfo("提示", "请先点击选择要移动的页面")
            return
        if idx > 0:
            self.sort_order[idx], self.sort_order[idx-1] = self.sort_order[idx-1], self.sort_order[idx]
            self.refresh_sort_viewer()

    def sort_move_down(self):
        idx = self.get_selected_sort_index()
        if idx is None:
            messagebox.showinfo("提示", "请先点击选择要移动的页面")
            return
        if idx < len(self.sort_order)-1:
            self.sort_order[idx], self.sort_order[idx+1] = self.sort_order[idx+1], self.sort_order[idx]
            self.refresh_sort_viewer()

    def sort_move_top(self):
        idx = self.get_selected_sort_index()
        if idx is None:
            messagebox.showinfo("提示", "请先点击选择要移动的页面")
            return
        if idx > 0:
            item = self.sort_order.pop(idx)
            self.sort_order.insert(0, item)
            self.refresh_sort_viewer()

    def sort_move_bottom(self):
        idx = self.get_selected_sort_index()
        if idx is None:
            messagebox.showinfo("提示", "请先点击选择要移动的页面")
            return
        if idx < len(self.sort_order)-1:
            item = self.sort_order.pop(idx)
            self.sort_order.append(item)
            self.refresh_sort_viewer()

    def refresh_sort_viewer(self):
        """刷新排序视图"""
        if self.sort_viewer:
            self.sort_viewer.destroy()
        self.sort_viewer = ThumbnailViewer(
            parent=self.sort_viewer_frame,
            doc=self.sort_doc,
            page_indices=self.sort_order,
            max_columns=5,
            thumbnail_size=120,
            select_mode='single'
        )

    def save_sorted_pdf(self):
        if not self.sort_doc:
            messagebox.showerror("错误", "未加载 PDF")
            return
        output = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output:
            return
        try:
            out_abs = self.validate_output_path(output, create_dir=True)
            self.run_thread(
                target=self.save_sorted,
                args=(self.sort_doc, self.sort_order, out_abs),
                on_success=lambda r: messagebox.showinfo("完成", f"排序结果已保存到：{r}"),
                on_error=lambda e: messagebox.showerror("错误", f"保存失败：{e}"),
                status_msg="正在保存排序后的 PDF..."
            )
        except Exception as e:
            messagebox.showerror("错误", f"输出路径无效：{e}")

    def save_sorted(self, doc, order, output_path):
        new_doc = fitz.open()
        for page_idx in order:
            new_doc.insert_pdf(doc, from_page=page_idx, to_page=page_idx)
        new_doc.save(output_path)
        new_doc.close()
        return output_path


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFToolbox(root)
    root.mainloop()