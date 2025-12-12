#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import os, sys, datetime, threading
from pathlib import Path
from tkinter import *
from tkinter import ttk, filedialog
from pdf2docx import Converter


class ConvertThread(threading.Thread):
    def __init__(self, pdf_path, out_dir, progress_cb, msg_cb, done_cb):
        super().__init__(daemon=True)
        self.pdf_path = pdf_path
        self.out_dir = Path(out_dir)
        self.progress_cb = progress_cb
        self.msg_cb = msg_cb
        self.done_cb = done_cb
        self.docx_path = ""

    def run(self):
        try:
            self.out_dir.mkdir(parents=True, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
            name = Path(self.pdf_path).stem
            self.docx_path = str(self.out_dir / f"{name}_{stamp}.docx")

            cv = Converter(self.pdf_path)
            total = len(cv.pages)
            cur = 0

            def cb(prog, desc):
                nonlocal cur
                if desc.get("event") == "page_parsed":
                    cur += 1
                    self.progress_cb(int(cur / total * 100))
                    self.msg_cb(f"正在解析第 {cur}/{total} 页...")

            cv.convert(self.docx_path, progress_callback=cb)
            cv.close()
            self.progress_cb(100)  # 推满进度条
            self.done_cb(True, self.docx_path)
        except Exception as e:
            self.done_cb(False, str(e))

class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("PDF 转 Word 小工具")
        self.resizable(False, False)
        # 稍微加高，防止紧凑
        self.geometry("580x220")
        self.pdf_path = ""
        self.word_path = ""

        # ---- 变量 ----
        self.pdf_var = StringVar()
        self.out_var = StringVar(value=str(Path.home() / "Desktop/PDF转换结果"))

        # ---- 统一间距 ----
        padx_val = 12
        pady_val = 6

        # ---- 界面 ----
        frm = Frame(self)
        frm.pack(fill=X, padx=padx_val, pady=pady_val)

        # PDF 选择行
        Label(frm, text="PDF 文件：", width=10, anchor=W).grid(row=0, column=0, sticky=W)
        Entry(frm, textvariable=self.pdf_var, width=48).grid(row=0, column=1, padx=(0, 5))
        Button(frm, text="浏览...", width=8, command=self.browse_pdf).grid(row=0, column=2)

        # 输出目录行
        Label(frm, text="输出目录：", width=10, anchor=W).grid(row=1, column=0, sticky=W, pady=(pady_val, 0))
        Entry(frm, textvariable=self.out_var, width=48).grid(row=1, column=1, padx=(0, 5), pady=(pady_val, 0))
        Button(frm, text="浏览...", width=8, command=self.browse_out).grid(row=1, column=2, pady=(pady_val, 0))

        # 进度条
        self.progress = ttk.Progressbar(self, length=540, mode='determinate')
        self.progress.pack(pady=pady_val, padx=padx_val)

        # 状态标签
        self.status = Label(self, text="准备就绪", anchor=W)
        self.status.pack(fill=X, padx=padx_val)

        # 按钮行：开始转换（左）  打开文件（右）
        btn_frm = Frame(self)
        btn_frm.pack(fill=X, padx=padx_val, pady=pady_val)
        self.btn_start = Button(btn_frm, text="开始转换", width=12, command=self.start_convert)
        self.btn_start.pack(side=LEFT)
        self.btn_open = Button(btn_frm, text="打开文件", width=12, command=self.open_word)
        self.btn_open.pack_forget()          # 先隐藏

    # ---------- 逻辑 ----------
    def browse_pdf(self):
        path = filedialog.askopenfilename(title="选择 PDF 文件", filetypes=[("PDF 文件", "*.pdf")])
        if path:
            self.pdf_var.set(path)

    def browse_out(self):
        path = filedialog.askdirectory(title="选择输出文件夹")
        if path:
            self.out_var.set(path)

    def start_convert(self):
        self.pdf_path = self.pdf_var.get()
        if not self.pdf_path:
            self.status.config(text="请先选择 PDF 文件！")
            return

        out_dir = self.out_var.get()
        if not out_dir:
            self.status.config(text="请选择输出目录！")
            return

        # 重置界面
        self.btn_open.pack_forget()
        self.status.config(text="正在转换，请稍候...")
        self.progress["value"] = 0
        self.btn_start.config(state=DISABLED)

        self.worker = ConvertThread(
            self.pdf_path, out_dir,
            progress_cb=lambda v: self.progress.config(value=v),
            msg_cb=lambda s: self.status.config(text=s),
            done_cb=self.on_finished
        )
        self.worker.start()

    def on_finished(self, ok: bool, path: str):
        self.btn_start.config(state=NORMAL)
        if ok:
            self.word_path = path
            self.status.config(text="转换成功！")
            self.btn_open.pack(side=RIGHT)          # 显示到右侧
        else:
            self.status.config(text=f"转换失败：{path}")

    def open_word(self):
        if self.word_path and Path(self.word_path).is_file():
            os.startfile(self.word_path)


if __name__ == "__main__":
    App().mainloop()