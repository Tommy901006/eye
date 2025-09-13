# -*- coding: utf-8 -*-
"""
Batch Eye-Tracking Upsampler (Tkinter)
- 選擇資料夾，批次處理其中所有 CSV/Excel 檔
- 自動偵測 T/X/Y 欄位
- 設定輸出取樣率、XY 方法、時間單位
- 可選輸出格式：CSV 或 Excel
- 每個檔案另存，檔名加 "_upsampled"
"""

import os
import re
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def detect_columns(df: pd.DataFrame):
    cols = list(df.columns)
    t_candidates = [c for c in cols if re.search(r'^(t|time|timestamp)', str(c), flags=re.I)]
    t_candidates += [c for c in cols if re.search(r'時間|毫秒|秒', str(c), flags=re.I)]
    t_candidates = list(dict.fromkeys(t_candidates))
    x_candidates = [c for c in cols if re.fullmatch(r'(?i)x', str(c))] or [c for c in cols if re.search(r'(?i)x', str(c))]
    y_candidates = [c for c in cols if re.fullmatch(r'(?i)y', str(c))] or [c for c in cols if re.search(r'(?i)y', str(c))]
    t_col = t_candidates[0] if t_candidates else (cols[0] if cols else None)
    rest = [c for c in cols if c != t_col]
    x_col = x_candidates[0] if x_candidates else (rest[0] if len(rest) > 0 else None)
    rest = [c for c in rest if c != x_col]
    y_col = y_candidates[0] if y_candidates else (rest[0] if len(rest) > 0 else None)
    return t_col, x_col, y_col

def upsample_xy(df, t_col, x_col, y_col, target_rate=1000, method="nearest",
                input_time_unit="ms", output_time_unit="ms"):
    t = df[t_col].astype(float).to_numpy()
    if input_time_unit == "s":
        t_ms = t * 1000.0
    else:
        t_ms = t
    order = np.argsort(t_ms)
    t_ms = t_ms[order]
    x = df[x_col].astype(float).to_numpy()[order] if x_col else None
    y = df[y_col].astype(float).to_numpy()[order] if y_col else None
    step_ms = 1000.0 / float(target_rate)
    t_min, t_max = np.nanmin(t_ms), np.nanmax(t_ms)
    t_new_ms = np.arange(t_min, t_max + step_ms, step_ms)
    if method == "nearest":
        idx = np.searchsorted(t_ms, t_new_ms, side="right") - 1
        idx = np.clip(idx, 0, len(t_ms) - 1)
        x_new = x[idx] if x is not None else None
        y_new = y[idx] if y is not None else None
    else:
        x_new = np.interp(t_new_ms, t_ms, x) if x is not None else None
        y_new = np.interp(t_new_ms, t_ms, y) if y is not None else None
    if output_time_unit == "s":
        t_out = t_new_ms / 1000.0
        time_name = "time_sec"
    else:
        t_out = t_new_ms
        time_name = "time_ms"
    out = pd.DataFrame({time_name: t_out})
    if x_new is not None: out["X"] = x_new
    if y_new is not None: out["Y"] = y_new
    return out

class BatchApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Batch Eye-Tracking Upsampler")
        self.geometry("720x540")
        self.folder = None

        frm_top = ttk.Frame(self)
        frm_top.pack(fill="x", padx=12, pady=10)
        ttk.Button(frm_top, text="選擇資料夾", command=self.pick_folder).pack(side="left")
        self.lbl_folder = ttk.Label(frm_top, text="未選擇資料夾", foreground="#666")
        self.lbl_folder.pack(side="left", padx=10)

        frm_cfg = ttk.LabelFrame(self, text="設定")
        frm_cfg.pack(fill="x", padx=12, pady=8)
        ttk.Label(frm_cfg, text="輸出取樣率 (Hz)：").grid(row=0, column=0, sticky="e", padx=6, pady=6)
        self.ent_rate = ttk.Entry(frm_cfg, width=10)
        self.ent_rate.insert(0, "1000")
        self.ent_rate.grid(row=0, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(frm_cfg, text="XY 方法：").grid(row=0, column=2, sticky="e", padx=6, pady=6)
        self.cmb_method = ttk.Combobox(frm_cfg, state="readonly",
                                       values=["nearest (不插值，最近鄰保持)", "linear (線性插值)"],
                                       width=28)
        self.cmb_method.current(0)
        self.cmb_method.grid(row=0, column=3, sticky="w", padx=6, pady=6)
        ttk.Label(frm_cfg, text="輸入時間單位：").grid(row=1, column=0, sticky="e", padx=6, pady=6)
        self.cmb_in_unit = ttk.Combobox(frm_cfg, state="readonly", values=["ms (毫秒)", "s (秒)"], width=15)
        self.cmb_in_unit.current(0)
        self.cmb_in_unit.grid(row=1, column=1, sticky="w", padx=6, pady=6)
        ttk.Label(frm_cfg, text="輸出時間單位：").grid(row=1, column=2, sticky="e", padx=6, pady=6)
        self.cmb_out_unit = ttk.Combobox(frm_cfg, state="readonly", values=["ms (毫秒)", "s (秒)"], width=15)
        self.cmb_out_unit.current(0)
        self.cmb_out_unit.grid(row=1, column=3, sticky="w", padx=6, pady=6)
        ttk.Label(frm_cfg, text="輸出格式：").grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.cmb_out_fmt = ttk.Combobox(frm_cfg, state="readonly", values=["CSV", "Excel"], width=15)
        self.cmb_out_fmt.current(0)
        self.cmb_out_fmt.grid(row=2, column=1, sticky="w", padx=6, pady=6)

        ttk.Button(self, text="執行批次上採樣", command=self.run_batch).pack(pady=12)
        self.txt_info = tk.Text(self, height=18)
        self.txt_info.pack(fill="both", expand=True, padx=12, pady=8)

    def log(self, msg):
        self.txt_info.insert("end", str(msg) + "\n")
        self.txt_info.see("end")

    def pick_folder(self):
        folder = filedialog.askdirectory(title="選擇資料夾")
        if folder:
            self.folder = folder
            self.lbl_folder.config(text=folder)
            self.log(f"已選擇資料夾：{folder}")

    def run_batch(self):
        if not self.folder:
            messagebox.showinfo("提示", "請先選擇資料夾")
            return
        try:
            rate = float(self.ent_rate.get())
            assert rate > 0
        except:
            messagebox.showerror("錯誤", "取樣率需為正數")
            return
        method = "nearest" if self.cmb_method.get().startswith("nearest") else "linear"
        in_unit = "ms" if self.cmb_in_unit.get().startswith("ms") else "s"
        out_unit = "ms" if self.cmb_out_unit.get().startswith("ms") else "s"
        out_fmt = self.cmb_out_fmt.get()

        files = [f for f in os.listdir(self.folder) if f.lower().endswith((".csv", ".xlsx", ".xls"))]
        if not files:
            messagebox.showwarning("沒有檔案", "資料夾裡沒有 CSV/Excel 檔")
            return

        for fname in files:
            fpath = os.path.join(self.folder, fname)
            try:
                if fname.lower().endswith(".csv"):
                    df = pd.read_csv(fpath)
                else:
                    df = pd.read_excel(fpath)
                t_col, x_col, y_col = detect_columns(df)
                out_df = upsample_xy(df, t_col, x_col, y_col,
                                     target_rate=rate, method=method,
                                     input_time_unit=in_unit, output_time_unit=out_unit)
                base, _ = os.path.splitext(fname)
                if out_fmt == "CSV":
                    out_name = base + "_upsampled.csv"
                    out_df.to_csv(os.path.join(self.folder, out_name), index=False)
                else:
                    out_name = base + "_upsampled.xlsx"
                    out_df.to_excel(os.path.join(self.folder, out_name), index=False)
                self.log(f"處理完成：{fname} -> {out_name}")
            except Exception as e:
                self.log(f"錯誤處理 {fname}：{e}")

        messagebox.showinfo("完成", "批次上採樣已完成！")

if __name__ == "__main__":
    app = BatchApp()
    app.mainloop()
