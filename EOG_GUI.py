import pandas as pd
import os
import glob
import platform
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from ttkbootstrap.widgets import Meter

class SignalSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 批次訊號切割工具 - 完整美化版")
        self.folder_path = ""
        self.setup_ui()

    def setup_ui(self):
        frm = tb.Frame(self.root, padding=20)
        frm.grid(row=0, column=0, sticky="nsew")

        # ===== 資料夾選擇 =====
        folder_frame = tb.LabelFrame(frm, text="步驟一：選擇資料夾", bootstyle=INFO)
        folder_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=10)
        tb.Button(folder_frame, text="📂 選擇 CSV 資料夾", bootstyle=PRIMARY, command=self.select_folder)\
            .grid(row=0, column=0, padx=10, pady=10, sticky='w')

        # ===== 設定參數 =====
        config_frame = tb.LabelFrame(frm, text="步驟二：設定參數", bootstyle=SECONDARY)
        config_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=10)

        # 取樣率
        tb.Label(config_frame, text="取樣率（Hz）:").grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.sampling_entry = tb.Entry(config_frame, width=10)
        self.sampling_entry.insert(0, "500")
        self.sampling_entry.grid(row=0, column=1, sticky='w', padx=5)

        # 第一段時間（Pre）
        tb.Label(config_frame, text="第一段時間（Pre）分鐘（起 ~ 迄）:").grid(row=1, column=0, sticky='e', padx=5, pady=5)
        frame1 = tb.Frame(config_frame)
        frame1.grid(row=1, column=1, sticky='w')
        self.time1_start = tb.Entry(frame1, width=5)
        self.time1_start.grid(row=0, column=0)
        tb.Label(frame1, text=" ~ ").grid(row=0, column=1)
        self.time1_end = tb.Entry(frame1, width=5)
        self.time1_end.grid(row=0, column=2)

        # 第二段時間（Post）
        tb.Label(config_frame, text="第二段時間（Post）分鐘（起 ~ 迄）:").grid(row=2, column=0, sticky='e', padx=5, pady=5)
        frame2 = tb.Frame(config_frame)
        frame2.grid(row=2, column=1, sticky='w')
        self.time2_start = tb.Entry(frame2, width=5)
        self.time2_start.grid(row=0, column=0)
        tb.Label(frame2, text=" ~ ").grid(row=0, column=1)
        self.time2_end = tb.Entry(frame2, width=5)
        self.time2_end.grid(row=0, column=2)

        # ===== 執行處理 =====
        action_frame = tb.LabelFrame(frm, text="步驟三：執行處理", bootstyle=SUCCESS)
        action_frame.grid(row=2, column=0, columnspan=2, sticky="ew", pady=10)
        tb.Button(action_frame, text="✂️ 執行批次切割", bootstyle=SUCCESS, command=self.batch_process)\
            .grid(row=0, column=0, padx=10, pady=15, sticky='ew')

        # 狀態列與進度條
        self.status_label = tb.Label(frm, text="", bootstyle=INFO)
        self.status_label.grid(row=3, column=0, columnspan=2, pady=5)

        self.progress_bar = tb.Progressbar(frm, orient="horizontal", mode="determinate",
                                           length=300, bootstyle=INFO)
        self.progress_bar.grid(row=4, column=0, columnspan=2, pady=10)

    def select_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.folder_path = folder
            messagebox.showinfo("✅ 資料夾選擇完成", f"將處理資料夾中所有 CSV：\n{folder}")

    def open_folder(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                os.system(f"open '{path}'")
            else:
                os.system(f"xdg-open '{path}'")
        except Exception as e:
            print(f"無法自動開啟資料夾：{e}")

    def batch_process(self):
        try:
            if not self.folder_path:
                messagebox.showwarning("⚠️ 尚未選擇資料夾", "請先選擇資料夾")
                return

            rate = int(self.sampling_entry.get())
            min1_start = int(self.time1_start.get())
            min1_end = int(self.time1_end.get())
            min2_start = int(self.time2_start.get())
            min2_end = int(self.time2_end.get())

            idx1_start = min1_start * 60 * rate
            idx1_end = min1_end * 60 * rate
            idx2_start = min2_start * 60 * rate
            idx2_end = min2_end * 60 * rate

            pre_dir = os.path.join(self.folder_path, "Pre")
            post_dir = os.path.join(self.folder_path, "Post")
            os.makedirs(pre_dir, exist_ok=True)
            os.makedirs(post_dir, exist_ok=True)

            files = glob.glob(os.path.join(self.folder_path, "*.csv"))
            total_files = len(files)
            if total_files == 0:
                messagebox.showwarning("找不到檔案", "該資料夾中沒有任何 CSV 檔案")
                return

            self.status_label.config(text="⏳ 資料處理中...請稍候...")
            self.progress_bar["maximum"] = total_files
            self.progress_bar["value"] = 0
            self.root.update_idletasks()

            count = 0
            for i, file in enumerate(files):
                df = pd.read_csv(file)
                base = os.path.splitext(os.path.basename(file))[0]

                segment1 = df.iloc[idx1_start:idx1_end].reset_index(drop=True)
                segment2 = df.iloc[idx2_start:idx2_end].reset_index(drop=True)

                segment1.to_csv(os.path.join(pre_dir, f"{base}_Pre.csv"), index=False)
                segment2.to_csv(os.path.join(post_dir, f"{base}_Post.csv"), index=False)
                count += 1

                self.progress_bar["value"] = i + 1
                self.root.update_idletasks()

            # ✅ 處理完成，自動開啟資料夾並關閉程式
            self.open_folder(pre_dir)
            self.root.destroy()

        except Exception as e:
            messagebox.showerror("錯誤", f"處理失敗：\n{e}")
            self.status_label.config(text="❌ 發生錯誤")

if __name__ == "__main__":
    app = tb.Window(themename="flatly")  # 可改為 minty, cyborg, journal 等
    SignalSplitterApp(app)
    app.mainloop()
