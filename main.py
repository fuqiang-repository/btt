"""TemplateForge — Word 模板批量生成工具 GUI 入口"""

import os
import platform
import subprocess
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from core.data_loader import load_data, get_columns, load_appendix_data
from core.engine import find_placeholders, generate_single, generate_documents
from core.excel_splitter import split_excel, split_excel_by_rows, get_total_rows


class TemplateForgeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("TemplateForge — 文档批量处理工具")
        self.root.geometry("820x820")
        self.root.minsize(720, 700)

        # ── Word 生成状态 ──
        self.template_path = tk.StringVar()
        self.data_path = tk.StringVar()
        self.appendix_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.filename_template = tk.StringVar(value="{序号}_{经销商名称}")
        self.appendix_filter_col = tk.StringVar()
        self.appendix_match_col = tk.StringVar()
        self.is_generating = False
        self.preview_row = tk.IntVar(value=1)

        self.data_columns: list[str] = []
        self.appendix_columns: list[str] = []
        self.data_rows: list[dict] = []

        # ── Excel 拆分状态 ──
        self.split_excel_path = tk.StringVar()
        self.split_output_dir = tk.StringVar()
        self.split_column = tk.StringVar()
        self.split_filename_template = tk.StringVar(value="{分组值}")
        self.split_header_row = tk.IntVar(value=1)
        self.split_mode = tk.StringVar(value="按分组列拆分")
        self.split_rows_per_file = tk.IntVar(value=100)
        self.is_splitting = False

        self.split_columns: list[str] = []
        self.split_max_row: int = 1

        self._build_ui()

    # ─────────────────── UI 构建 ───────────────────

    def _build_ui(self):
        # 顶部标题
        header = ttk.Frame(self.root, padding=(16, 12, 16, 4))
        header.pack(fill="x")
        font_name = "PingFang SC" if platform.system() == "Darwin" else "Microsoft YaHei"
        ttk.Label(
            header, text="📄 TemplateForge — 文档批量处理工具",
            font=(font_name, 18, "bold"),
        ).pack(side="left")

        # Notebook 标签页
        self.notebook = ttk.Notebook(self.root, padding=(12, 8))
        self.notebook.pack(fill="both", expand=True, padx=8, pady=(0, 4))

        # Word 生成标签页
        self.word_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.word_tab, text="📝 Word 模板生成")
        self._build_word_tab()

        # Excel 拆分标签页
        self.split_tab = ttk.Frame(self.notebook, padding=8)
        self.notebook.add(self.split_tab, text="📊 Excel 拆分")
        self._build_split_tab()

        # ── 共用底部区域（进度 + 日志）──
        bottom = ttk.Frame(self.root, padding=(16, 4, 16, 12))
        bottom.pack(fill="both", expand=False)

        self.progress = ttk.Progressbar(bottom, mode="determinate", length=400)
        self.progress.pack(fill="x", pady=(0, 4))
        self.status_label = ttk.Label(bottom, text="就绪", foreground="gray")
        self.status_label.pack(fill="x", pady=(0, 8))

        log_frame = ttk.LabelFrame(bottom, text="操作日志", padding=6)
        log_frame.pack(fill="both", expand=True)

        mono_font = "Menlo" if platform.system() == "Darwin" else "Consolas"
        self.log_text = tk.Text(
            log_frame, height=6, state="disabled", wrap="word", font=(mono_font, 10)
        )
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    # ─────────────────── Word 生成标签页 ───────────────────

    def _build_word_tab(self):
        main = self.word_tab

        # ── 文件选择区 ──
        file_frame = ttk.LabelFrame(main, text="文件选择", padding=10)
        file_frame.pack(fill="x", pady=(0, 8))

        self._add_file_row(file_frame, "模板文件 (.docx):", self.template_path,
                           self._on_template_selected, filetypes=[("Word 文档", "*.docx")])
        self._add_file_row(file_frame, "数据文件 (.xlsx/.csv):", self.data_path,
                           self._on_data_selected, filetypes=[("数据文件", "*.xlsx *.csv")])
        self._add_file_row(file_frame, "附录文件 (.xlsx) [可选]:", self.appendix_path,
                           self._on_appendix_selected, filetypes=[("Excel 文件", "*.xlsx")])
        self._add_dir_row(file_frame, "输出目录:", self.output_dir)

        # ── 配置区 ──
        config_frame = ttk.LabelFrame(main, text="生成配置", padding=10)
        config_frame.pack(fill="x", pady=(0, 8))

        row0 = ttk.Frame(config_frame)
        row0.pack(fill="x", pady=2)
        ttk.Label(row0, text="文件名模板:").pack(side="left")
        ttk.Entry(row0, textvariable=self.filename_template, width=40).pack(side="left", padx=(8, 4))
        ttk.Label(row0, text="支持 {序号}、{列名} 变量", foreground="gray").pack(side="left")

        row1 = ttk.Frame(config_frame)
        row1.pack(fill="x", pady=2)
        ttk.Label(row1, text="附录筛选列:").pack(side="left")
        self.appendix_filter_combo = ttk.Combobox(
            row1, textvariable=self.appendix_filter_col, width=20, state="readonly"
        )
        self.appendix_filter_combo.pack(side="left", padx=(8, 16))
        ttk.Label(row1, text="匹配数据列:").pack(side="left")
        self.appendix_match_combo = ttk.Combobox(
            row1, textvariable=self.appendix_match_col, width=20, state="readonly"
        )
        self.appendix_match_combo.pack(side="left", padx=(8, 4))

        # ── 数据预览区 ──
        preview_frame = ttk.LabelFrame(main, text="数据预览（前 5 行）", padding=6)
        preview_frame.pack(fill="both", expand=True, pady=(0, 8))

        cols_frame = ttk.Frame(preview_frame)
        cols_frame.pack(fill="x", pady=(0, 4))
        ttk.Label(cols_frame, text="检测到的占位符: ").pack(side="left")
        self.placeholders_label = ttk.Label(
            cols_frame, text="（请先选择模板和数据文件）", foreground="gray"
        )
        self.placeholders_label.pack(side="left")

        # 预览表格
        tree_frame = ttk.Frame(preview_frame)
        tree_frame.pack(fill="both", expand=True)

        self.preview_tree = ttk.Treeview(tree_frame, show="headings", height=5)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.preview_tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.preview_tree.xview)
        self.preview_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # ── 操作区 ──
        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=(0, 8))

        self.preview_btn = ttk.Button(action_frame, text="👁 预览", command=self._on_preview)
        self.preview_btn.pack(side="left", padx=(0, 4))
        ttk.Label(action_frame, text="第").pack(side="left")
        self.preview_spin = ttk.Spinbox(
            action_frame, textvariable=self.preview_row, width=4, from_=1, to=9999
        )
        self.preview_spin.pack(side="left", padx=(2, 2))
        ttk.Label(action_frame, text="行").pack(side="left", padx=(0, 8))

        ttk.Separator(action_frame, orient="vertical").pack(side="left", fill="y", padx=4)

        self.generate_btn = ttk.Button(action_frame, text="🚀 开始生成", command=self._on_generate)
        self.generate_btn.pack(side="left", padx=(0, 8))

        self.open_dir_btn = ttk.Button(
            action_frame, text="📂 打开输出目录", command=self._open_output_dir, state="disabled"
        )
        self.open_dir_btn.pack(side="left")

    # ─────────────────── Excel 拆分标签页 ───────────────────

    def _build_split_tab(self):
        main = self.split_tab

        # ── 文件选择区 ──
        file_frame = ttk.LabelFrame(main, text="文件选择", padding=10)
        file_frame.pack(fill="x", pady=(0, 8))

        self._add_file_row(
            file_frame, "Excel 文件 (.xlsx):", self.split_excel_path,
            self._on_split_excel_selected, filetypes=[("Excel 文件", "*.xlsx")]
        )
        self._add_dir_row(file_frame, "输出目录:", self.split_output_dir)

        # ── 配置区 ──
        config_frame = ttk.LabelFrame(main, text="拆分配置", padding=10)
        config_frame.pack(fill="x", pady=(0, 8))

        # 表头所在行
        row_header = ttk.Frame(config_frame)
        row_header.pack(fill="x", pady=2)
        ttk.Label(row_header, text="表头所在行:").pack(side="left")
        self.split_header_spin = ttk.Spinbox(
            row_header, textvariable=self.split_header_row, width=6, from_=1, to=9999,
            command=self._on_split_header_row_changed
        )
        self.split_header_spin.pack(side="left", padx=(8, 4))
        ttk.Label(row_header, text="（第几行是列名，默认第 1 行）", foreground="gray").pack(side="left")

        # 拆分模式
        row_mode = ttk.Frame(config_frame)
        row_mode.pack(fill="x", pady=2)
        ttk.Label(row_mode, text="拆分模式:").pack(side="left")
        self.split_mode_combo = ttk.Combobox(
            row_mode, textvariable=self.split_mode,
            values=["按分组列拆分", "按行数拆分"],
            width=16, state="readonly"
        )
        self.split_mode_combo.pack(side="left", padx=(8, 4))
        self.split_mode_combo.bind("<<ComboboxSelected>>", self._on_split_mode_changed)

        # 分组列（按分组列拆分模式时显示）
        self.split_group_frame = ttk.Frame(config_frame)
        self.split_group_frame.pack(fill="x", pady=2)
        ttk.Label(self.split_group_frame, text="分组列:").pack(side="left")
        self.split_column_combo = ttk.Combobox(
            self.split_group_frame, textvariable=self.split_column, width=24, state="readonly"
        )
        self.split_column_combo.pack(side="left", padx=(8, 4))
        ttk.Label(self.split_group_frame, text="（按此列的唯一值拆分）", foreground="gray").pack(side="left")

        # 每份行数（按行数拆分模式时显示）
        self.split_rows_frame = ttk.Frame(config_frame)
        ttk.Label(self.split_rows_frame, text="每份行数:").pack(side="left")
        self.split_rows_spin = ttk.Spinbox(
            self.split_rows_frame, textvariable=self.split_rows_per_file, width=8, from_=1, to=100000
        )
        self.split_rows_spin.pack(side="left", padx=(8, 4))
        ttk.Label(self.split_rows_frame, text="（每个文件包含的数据行数）", foreground="gray").pack(side="left")
        # 默认隐藏按行数拆分的控件
        # self.split_rows_frame 不 pack，初始不显示

        # 文件名模板
        self.split_fn_frame = ttk.Frame(config_frame)
        self.split_fn_frame.pack(fill="x", pady=2)
        ttk.Label(self.split_fn_frame, text="文件名模板:").pack(side="left")
        ttk.Entry(self.split_fn_frame, textvariable=self.split_filename_template, width=40).pack(side="left", padx=(8, 4))
        self.split_fn_hint = ttk.Label(
            self.split_fn_frame, text="支持 {分组值}、{序号} 变量", foreground="gray"
        )
        self.split_fn_hint.pack(side="left")

        # ── 操作区 ──
        action_frame = ttk.Frame(main)
        action_frame.pack(fill="x", pady=(0, 8))

        self.split_btn = ttk.Button(action_frame, text="✂️ 开始拆分", command=self._on_split)
        self.split_btn.pack(side="left", padx=(0, 8))

        self.open_split_dir_btn = ttk.Button(
            action_frame, text="📂 打开输出目录", command=self._open_split_output_dir, state="disabled"
        )
        self.open_split_dir_btn.pack(side="left")

    # ─────────────────── UI 辅助 ───────────────────

    def _add_file_row(self, parent, label_text, var, command, filetypes=None):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label_text, width=24, anchor="e").pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", padx=(8, 4), fill="x", expand=True)
        ttk.Button(row, text="选择...", width=8,
                   command=lambda: self._choose_file(var, command, filetypes)).pack(side="left")

    def _add_dir_row(self, parent, label_text, var):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label_text, width=24, anchor="e").pack(side="left")
        ttk.Entry(row, textvariable=var).pack(side="left", padx=(8, 4), fill="x", expand=True)
        ttk.Button(row, text="选择...", width=8,
                   command=lambda: self._choose_dir(var)).pack(side="left")

    @staticmethod
    def _choose_file(var, callback=None, filetypes=None):
        path = filedialog.askopenfilename(filetypes=filetypes or [("所有文件", "*.*")])
        if path:
            var.set(path)
            if callback:
                callback()

    @staticmethod
    def _choose_dir(var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    # ─────────────────── Word 生成回调 ───────────────────

    def _on_template_selected(self):
        path = self.template_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            placeholders = find_placeholders(path)
            self.placeholders_label.config(
                text="、".join(f"【{p}】" for p in placeholders),
                foreground="black",
            )
            self._log(f"✅ 模板已加载，检测到 {len(placeholders)} 个占位符")
        except Exception as e:
            self._log(f"❌ 读取模板失败: {e}")
            self.placeholders_label.config(text=f"读取失败: {e}", foreground="red")

    def _on_data_selected(self):
        path = self.data_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            self.data_columns = get_columns(path)
            self.data_rows = load_data(path)
            self._refresh_preview()
            self.appendix_match_combo["values"] = self.data_columns
            self._log(f"✅ 数据已加载，共 {len(self.data_rows)} 行，列: {', '.join(self.data_columns)}")
        except Exception as e:
            self._log(f"❌ 读取数据文件失败: {e}")

    def _on_appendix_selected(self):
        path = self.appendix_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            headers, _ = load_appendix_data(path)
            self.appendix_columns = headers
            self.appendix_filter_combo["values"] = headers
            self._log(f"✅ 附录已加载，共 {len(headers)} 列: {', '.join(headers[:6])}...")
        except Exception as e:
            self._log(f"❌ 读取附录文件失败: {e}")

    def _refresh_preview(self):
        tree = self.preview_tree
        tree.delete(*tree.get_children())
        if not self.data_columns:
            return

        tree["columns"] = list(self.data_columns)
        for col in self.data_columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, minwidth=60)

        for row in self.data_rows[:5]:
            values = [str(row.get(c, ""))[:60] for c in self.data_columns]
            tree.insert("", "end", values=values)

    def _on_preview(self):
        template = self.template_path.get()
        data_file = self.data_path.get()

        if not template or not os.path.isfile(template):
            messagebox.showerror("错误", "请先选择模板文件")
            return
        if not data_file or not os.path.isfile(data_file):
            messagebox.showerror("错误", "请先选择数据文件")
            return
        if not self.data_rows:
            self.data_rows = load_data(data_file)

        row_idx = self.preview_row.get() - 1
        if row_idx < 0 or row_idx >= len(self.data_rows):
            messagebox.showwarning("提示", f"行号超出范围，数据共 {len(self.data_rows)} 行")
            return

        row_data = self.data_rows[row_idx]
        appendix = self.appendix_path.get()
        filter_col = self.appendix_filter_col.get()
        match_col = self.appendix_match_col.get()

        import tempfile
        preview_dir = os.path.join(tempfile.gettempdir(), "templateforge_preview")
        preview_path = os.path.join(preview_dir, f"预览_第{row_idx + 1}行.docx")

        try:
            generate_single(
                template_path=template,
                row_data=row_data,
                output_path=preview_path,
                appendix_path=appendix if appendix and os.path.isfile(appendix) else None,
                appendix_filter_column=filter_col or None,
                appendix_match_data_column=match_col or None,
            )

            data_preview = ", ".join(f"{k}={v}" for k, v in row_data.items() if v)
            self._log(f"👁 预览第 {row_idx + 1} 行: {data_preview}")
            self._log(f"   📄 {preview_path}")
            self.status_label.config(text=f"👁 预览已生成: 第 {row_idx + 1} 行", foreground="blue")
            self._open_file(preview_path)

        except Exception as e:
            self._log(f"❌ 预览失败: {e}")
            messagebox.showerror("预览失败", str(e))

    def _on_generate(self):
        if self.is_generating:
            return

        template = self.template_path.get()
        data_file = self.data_path.get()
        output = self.output_dir.get()

        if not template or not os.path.isfile(template):
            messagebox.showerror("错误", "请选择有效的模板文件")
            return
        if not data_file or not os.path.isfile(data_file):
            messagebox.showerror("错误", "请选择有效的数据文件")
            return
        if not output:
            output = os.path.join(os.path.dirname(data_file), "output")
            self.output_dir.set(output)

        self.is_generating = True
        self.generate_btn.config(state="disabled", text="生成中...")
        self.progress["value"] = 0

        thread = threading.Thread(target=self._do_generate, daemon=True)
        thread.start()

    def _do_generate(self):
        template = self.template_path.get()
        data_file = self.data_path.get()
        appendix = self.appendix_path.get()
        output = self.output_dir.get()
        filename_tpl = self.filename_template.get()
        filter_col = self.appendix_filter_col.get()
        match_col = self.appendix_match_col.get()

        try:
            data = self.data_rows if self.data_rows else load_data(data_file)

            def on_progress(current, total, msg):
                self.root.after(0, lambda c=current, t=total, m=msg: self._update_progress(c, t, m))

            generated = generate_documents(
                template_path=template,
                data=data,
                output_dir=output,
                appendix_path=appendix if appendix and os.path.isfile(appendix) else None,
                appendix_filter_column=filter_col or None,
                appendix_match_data_column=match_col or None,
                filename_template=filename_tpl,
                progress_callback=on_progress,
            )

            self.root.after(0, lambda: self._on_generate_done(generated))
        except Exception as e:
            self.root.after(0, lambda: self._on_generate_error(str(e)))

    def _update_progress(self, current, total, msg):
        self.progress["maximum"] = total
        self.progress["value"] = current
        self.status_label.config(text=f"{msg} ({current}/{total})")

    def _on_generate_done(self, files: list[str]):
        self.is_generating = False
        self.generate_btn.config(state="normal", text="🚀 开始生成")
        self.progress["value"] = self.progress["maximum"]
        self.status_label.config(text=f"✅ 完成！共生成 {len(files)} 个文件", foreground="green")
        self.open_dir_btn.config(state="normal")
        self._log(f"\n🎉 生成完成！共 {len(files)} 个文件:")
        for f in files:
            self._log(f"   ✓ {os.path.basename(f)}")

    def _on_generate_error(self, error: str):
        self.is_generating = False
        self.generate_btn.config(state="normal", text="🚀 开始生成")
        self.status_label.config(text="❌ 生成失败", foreground="red")
        self._log(f"\n❌ 错误: {error}")
        messagebox.showerror("生成失败", error)

    # ─────────────────── Excel 拆分回调 ───────────────────

    def _on_split_mode_changed(self, _event=None):
        mode = self.split_mode.get()
        if mode == "按分组列拆分":
            # 显示分组列，隐藏每份行数
            self.split_rows_frame.pack_forget()
            self.split_group_frame.pack(fill="x", pady=2, before=self.split_fn_frame)
            self.split_fn_hint.config(text="支持 {分组值}、{序号} 变量")
            self.split_filename_template.set("{分组值}")
        else:
            # 显示每份行数，隐藏分组列
            self.split_group_frame.pack_forget()
            self.split_rows_frame.pack(fill="x", pady=2, before=self.split_fn_frame)
            self.split_fn_hint.config(text="支持 {序号} 变量")
            self.split_filename_template.set("{序号}")

    def _on_split_excel_selected(self):
        path = self.split_excel_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            header_row_num = self.split_header_row.get()
            self.split_columns = get_columns(path, header_row_num=header_row_num)
            self.split_column_combo["values"] = self.split_columns
            if self.split_columns:
                self.split_column.set(self.split_columns[0])
            # 读取总行数，更新表头行 Spinbox 范围
            self.split_max_row = get_total_rows(path)
            self.split_header_spin.config(to=self.split_max_row)
            self._log(f"✅ Excel 已加载，共 {len(self.split_columns)} 列、{self.split_max_row} 行")
        except Exception as e:
            self._log(f"❌ 读取 Excel 文件失败: {e}")

    def _on_split_header_row_changed(self):
        """表头行号变化时，重新读取列名"""
        path = self.split_excel_path.get()
        if not path or not os.path.isfile(path):
            return
        try:
            header_row_num = self.split_header_row.get()
            self.split_columns = get_columns(path, header_row_num=header_row_num)
            self.split_column_combo["values"] = self.split_columns
            if self.split_columns:
                self.split_column.set(self.split_columns[0])
            self._log(f"✅ 已刷新列名（第 {header_row_num} 行）: {', '.join(self.split_columns)}")
        except Exception as e:
            self._log(f"❌ 读取列名失败: {e}")

    def _on_split(self):
        if self.is_splitting:
            return

        filepath = self.split_excel_path.get()
        output = self.split_output_dir.get()
        mode = self.split_mode.get()

        if not filepath or not os.path.isfile(filepath):
            messagebox.showerror("错误", "请选择有效的 Excel 文件")
            return

        if mode == "按分组列拆分":
            group_col = self.split_column.get()
            if not group_col:
                messagebox.showerror("错误", "请选择分组列")
                return
        else:
            rows_per = self.split_rows_per_file.get()
            if rows_per <= 0:
                messagebox.showerror("错误", "每份行数必须大于 0")
                return

        if not output:
            output = os.path.join(os.path.dirname(filepath), "split_output")
            self.split_output_dir.set(output)

        self.is_splitting = True
        self.split_btn.config(state="disabled", text="拆分中...")
        self.progress["value"] = 0

        thread = threading.Thread(target=self._do_split, daemon=True)
        thread.start()

    def _do_split(self):
        filepath = self.split_excel_path.get()
        output = self.split_output_dir.get()
        filename_tpl = self.split_filename_template.get()
        header_row_num = self.split_header_row.get()
        mode = self.split_mode.get()

        try:
            def on_progress(current, total, msg):
                self.root.after(0, lambda c=current, t=total, m=msg: self._update_progress(c, t, m))

            if mode == "按分组列拆分":
                group_col = self.split_column.get()
                generated = split_excel(
                    filepath=filepath,
                    group_column=group_col,
                    output_dir=output,
                    filename_template=filename_tpl,
                    progress_callback=on_progress,
                    header_row_num=header_row_num,
                )
            else:
                rows_per = self.split_rows_per_file.get()
                generated = split_excel_by_rows(
                    filepath=filepath,
                    rows_per_file=rows_per,
                    output_dir=output,
                    filename_template=filename_tpl,
                    progress_callback=on_progress,
                    header_row_num=header_row_num,
                )

            self.root.after(0, lambda: self._on_split_done(generated))
        except Exception as e:
            self.root.after(0, lambda: self._on_split_error(str(e)))

    def _on_split_done(self, files: list[str]):
        self.is_splitting = False
        self.split_btn.config(state="normal", text="✂️ 开始拆分")
        self.progress["value"] = self.progress["maximum"]
        self.status_label.config(text=f"✅ 拆分完成！共生成 {len(files)} 个文件", foreground="green")
        self.open_split_dir_btn.config(state="normal")
        self._log(f"\n🎉 拆分完成！共 {len(files)} 个文件:")
        for f in files:
            self._log(f"   ✓ {os.path.basename(f)}")

    def _on_split_error(self, error: str):
        self.is_splitting = False
        self.split_btn.config(state="normal", text="✂️ 开始拆分")
        self.status_label.config(text="❌ 拆分失败", foreground="red")
        self._log(f"\n❌ 错误: {error}")
        messagebox.showerror("拆分失败", error)

    # ─────────────────── 工具 ───────────────────

    def _log(self, message: str):
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def _open_file(self, filepath: str):
        """用系统默认程序打开文件或目录"""
        system = platform.system()
        if system == "Darwin":
            subprocess.Popen(["open", filepath])
        elif system == "Windows":
            os.startfile(filepath)
        else:
            subprocess.Popen(["xdg-open", filepath])

    def _open_output_dir(self):
        path = self.output_dir.get()
        if not path or not os.path.isdir(path):
            messagebox.showwarning("提示", "输出目录不存在")
            return
        self._open_file(path)

    def _open_split_output_dir(self):
        path = self.split_output_dir.get()
        if not path or not os.path.isdir(path):
            messagebox.showwarning("提示", "输出目录不存在")
            return
        self._open_file(path)


def main():
    root = tk.Tk()
    style = ttk.Style(root)
    available_themes = style.theme_names()
    if "clam" in available_themes:
        style.theme_use("clam")

    app = TemplateForgeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
