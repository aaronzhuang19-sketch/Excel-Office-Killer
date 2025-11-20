import os
import threading
from typing import List

import customtkinter as ctk
from tkinterdnd2 import DND_FILES

from core.excel_engine import read_headers, split_excel, merge_excels


class DropFrame(ctk.CTkFrame):
    def __init__(self, master, text_var, on_files_dropped, **kwargs):
        super().__init__(master, **kwargs)
        self.text_var = text_var
        self.on_files_dropped = on_files_dropped
        self.label = ctk.CTkLabel(self, textvariable=self.text_var)
        self.label.pack(fill="both", expand=True, padx=12, pady=12)
        self.drop_target_register(DND_FILES)
        self.dnd_bind("<<Drop>>", self._handle_drop)

    def _handle_drop(self, event):
        files = list(self.tk.splitlist(event.data))
        excel_files = [f for f in files if f.lower().endswith(".xlsx")]
        if excel_files:
            self.on_files_dropped(excel_files)


class SplitterView(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.file_path = None
        self.headers: List[str] = []
        self._headers_job_id = 0
        self.title = ctk.CTkLabel(self, text="Excel 拆分", font=ctk.CTkFont(size=18, weight="bold"))
        self.title.pack(fill="x", padx=12, pady=(12, 0))
        self.subtitle = ctk.CTkLabel(self, text="拖入文件 → 选择列 → 开始拆分")
        self.subtitle.pack(fill="x", padx=12, pady=(0, 8))
        self.drop_text = ctk.StringVar(value="将 Excel 文件拖入此处")
        self.drop = DropFrame(self, text_var=self.drop_text, on_files_dropped=self._on_dropped)
        self.drop.pack(fill="x", padx=12, pady=12)

        self.column_label = ctk.CTkLabel(self, text="选择列")
        self.column_label.pack(anchor="w", padx=12)
        self.column_var = ctk.StringVar(value="")
        self.column_menu = ctk.CTkOptionMenu(self, variable=self.column_var, values=[])
        self.column_menu.pack(fill="x", padx=12, pady=6)

        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(fill="x", padx=12, pady=6)
        self.progress.set(0)

        self.action = ctk.CTkButton(self, text="开始拆分", command=self._start_split)
        self.action.pack(padx=12, pady=6)

        self.status = ctk.CTkLabel(self, text="")
        self.status.pack(fill="x", padx=12, pady=6)
        self.tip = ctk.CTkLabel(
            self,
            text="重要提示：本工具默认第1行为表头，数据从第2行开始；拆分按表头的列名进行识别。请确保表头非空且唯一。",
            text_color=("#6b7280", "#9ca3af")
        )
        self.tip.pack(fill="x", padx=12, pady=(0, 12))

    def _on_dropped(self, files: List[str]):
        self.file_path = files[0]
        self.drop_text.set(self.file_path)
        self.column_menu.configure(values=[])
        self.column_var.set("")
        self.status.configure(text="正在读取表头...")
        self._headers_job_id += 1
        jid = self._headers_job_id
        threading.Thread(target=self._load_headers_bg, args=(jid,), daemon=True).start()

    def _load_headers_bg(self, jid: int):
        try:
            headers = read_headers(self.file_path)
        except Exception as e:
            msg = str(e)
            self.after(0, lambda: self.status.configure(text=msg))
            return
        def update():
            if jid != self._headers_job_id:
                return
            self.headers = headers
            vals = headers if headers else []
            if vals:
                self.column_menu.configure(values=vals)
                self.column_var.set(vals[0])
                self.status.configure(text="表头读取完成")
            else:
                self.status.configure(text="未读取到任何列")
        self.after(0, update)

    def _start_split(self):
        if not self.file_path:
            self.status.configure(text="请先拖入Excel文件")
            return
        col = self.column_var.get()
        if not col:
            self.status.configure(text="请选择拆分列")
            return
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.status.configure(text="正在后台拆分...")
        threading.Thread(target=self._split_bg, args=(col,), daemon=True).start()

    def _split_bg(self, col: str):
        def progress(rows: int, files: int):
            self.after(0, lambda: self.status.configure(text=f"正在后台拆分... 已处理 {rows} 行，已准备生成 {files} 个文件（写盘在完成时进行）"))
        try:
            out_dir, written = split_excel(self.file_path, os.getcwd(), col, progress_cb=progress)
        except Exception as e:
            msg = str(e)
            def fail():
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self.progress.set(0)
                self.status.configure(text=msg)
            self.after(0, fail)
            return
        def done():
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(1)
            self.status.configure(text=f"完成，共生成 {len(written)} 个文件")
            try:
                os.startfile(out_dir)
            except Exception:
                pass
        self.after(0, done)


class MergerView(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.file_paths: List[str] = []
        self.title = ctk.CTkLabel(self, text="Excel 合并", font=ctk.CTkFont(size=18, weight="bold"))
        self.title.pack(fill="x", padx=12, pady=(12, 0))
        self.subtitle = ctk.CTkLabel(self, text="拖入多个文件 → 开始合并")
        self.subtitle.pack(fill="x", padx=12, pady=(0, 8))
        self.drop_text = ctk.StringVar(value="将 Excel 文件拖入此处")
        self.drop = DropFrame(self, text_var=self.drop_text, on_files_dropped=self._on_dropped)
        self.drop.pack(fill="x", padx=12, pady=12)

        self.textbox = ctk.CTkTextbox(self, height=160)
        self.textbox.pack(fill="both", expand=False, padx=12, pady=6)
        self.textbox.configure(state="disabled")

        self.progress = ctk.CTkProgressBar(self)
        self.progress.pack(fill="x", padx=12, pady=6)
        self.progress.set(0)

        self.action = ctk.CTkButton(self, text="开始合并", command=self._start_merge)
        self.action.pack(padx=12, pady=6)

        self.status = ctk.CTkLabel(self, text="")
        self.status.pack(fill="x", padx=12, pady=6)
        self.tip = ctk.CTkLabel(
            self,
            text="重要提示：本工具默认第1行为表头，数据从第2行开始；合并会按表头对齐列。请确保各文件的表头一致。",
            text_color=("#6b7280", "#9ca3af")
        )
        self.tip.pack(fill="x", padx=12, pady=(0, 12))

    def _on_dropped(self, files: List[str]):
        self.file_paths = files
        self.drop_text.set(f"已选择 {len(files)} 个文件")
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        for p in files:
            self.textbox.insert("end", p + "\n")
        self.textbox.configure(state="disabled")

    def _start_merge(self):
        if not self.file_paths:
            self.status.configure(text="请先拖入Excel文件")
            return
        self.progress.configure(mode="indeterminate")
        self.progress.start()
        self.status.configure(text="正在后台合并...")
        threading.Thread(target=self._merge_bg, daemon=True).start()

    def _merge_bg(self):
        try:
            out_dir, out_file = merge_excels(self.file_paths, os.getcwd())
        except Exception as e:
            msg = str(e)
            def fail():
                self.progress.stop()
                self.progress.configure(mode="determinate")
                self.progress.set(0)
                self.status.configure(text=msg)
            self.after(0, fail)
            return
        def done():
            self.progress.stop()
            self.progress.configure(mode="determinate")
            self.progress.set(1)
            self.status.configure(text=f"完成，已生成文件: {out_file}")
            try:
                os.startfile(out_dir)
            except Exception:
                pass
        self.after(0, done)