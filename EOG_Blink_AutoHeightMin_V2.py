import warnings
warnings.filterwarnings("ignore")

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import neurokit2 as nk
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

plt.rcParams['font.family'] = 'Microsoft JhengHei'
plt.rcParams['axes.unicode_minus'] = False

def run_analysis():
    try:
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not file_path:
            return

        sampling_rate = int(sampling_rate_entry.get())
        eog_column = eog_column_entry.get()
        height_min_input = height_min_entry.get().strip()

        start1 = float(start1_entry.get())
        end1 = float(end1_entry.get())
        start2 = float(start2_entry.get())
        end2 = float(end2_entry.get())

        df = pd.read_csv(file_path, encoding='utf-8-sig')
        full_signal = df[eog_column].fillna(method="ffill").values

        def process_segment(start_min, end_min, label):
            start_idx = int(start_min * 60 * sampling_rate)
            end_idx = int(end_min * 60 * sampling_rate)
            if end_idx > len(full_signal):
                end_idx = len(full_signal)
            eog_segment = full_signal[start_idx:end_idx]
            eog_filtered = nk.signal_filter(eog_segment, sampling_rate=sampling_rate, lowcut=0.5, highcut=8)

            # 自動或手動設定門檻
            if height_min_input.lower() == "auto" or height_min_input == "":
                height_min = round(np.std(eog_filtered) * 1.5, 4)
            else:
                height_min = float(height_min_input)

            raw_peaks = nk.signal_findpeaks(eog_filtered, height_min=height_min)["Peaks"]
            min_distance = int(0.25 * sampling_rate)

            def remove_close_peaks(peaks, min_dist):
                if len(peaks) == 0:
                    return np.array([])
                filtered = [peaks[0]]
                for idx in peaks[1:]:
                    if idx - filtered[-1] >= min_dist:
                        filtered.append(idx)
                return np.array(filtered)

            filtered_peaks = remove_close_peaks(raw_peaks, min_distance)
            blink_times = filtered_peaks / sampling_rate
            ibi = np.diff(blink_times)
            amplitudes = eog_filtered[filtered_peaks]

            stats = {
                "區段": label,
                "眨眼總數": len(filtered_peaks),
                "平均間隔 (秒)": round(np.mean(ibi), 3) if len(ibi) > 0 else None,
                "間隔標準差": round(np.std(ibi), 3) if len(ibi) > 0 else None,
                "每分鐘眨眼率": round(len(filtered_peaks) / ((end_idx - start_idx)/sampling_rate/60), 2),
                "平均振幅": round(np.mean(amplitudes), 3),
                "振幅標準差": round(np.std(amplitudes), 3),
                "使用門檻值 height_min": height_min
            }

            fig1, ax1 = plt.subplots(figsize=(10, 3))
            ax1.plot(eog_filtered, label="Filtered EOG")
            ax1.plot(filtered_peaks, eog_filtered[filtered_peaks], "ro", label="Detected Blinks")
            ax1.set_title(f"{label} - 眨眼偵測圖")
            ax1.set_xlabel("Samples")
            ax1.set_ylabel("Amplitude")
            ax1.legend()
            ax1.grid(True)

            fig2, ax2 = plt.subplots(figsize=(8, 3))
            total_duration_sec = len(eog_segment) / sampling_rate
            total_minutes = int(np.ceil(total_duration_sec / 60))
            minute_bins = np.arange(0, total_minutes + 1) * 60
            blink_counts_per_minute, _ = np.histogram(blink_times, bins=minute_bins)
            ax2.plot(range(len(blink_counts_per_minute)), blink_counts_per_minute, marker='o', linestyle='-')
            ax2.set_title(f"{label} - 每分鐘眨眼次數圖")
            ax2.set_xlabel("分鐘")
            ax2.set_ylabel("眨眼次數")
            ax2.grid(True)

            return stats, [fig1, fig2]

        stats1, figs1 = process_segment(start1, end1, "區段1")
        stats2, figs2 = process_segment(start2, end2, "區段2")

        result_text = "\n".join([f"{k}: {v}" for k, v in stats1.items()]) + "\n\n" + \
                      "\n".join([f"{k}: {v}" for k, v in stats2.items()])
        messagebox.showinfo("分析結果（兩段）", result_text)

        for fig in figs1 + figs2:
            canvas = FigureCanvasTkAgg(fig, master=window)
            canvas.draw()
            canvas.get_tk_widget().pack()

    except Exception as e:
        messagebox.showerror("錯誤", str(e))

# GUI
window = tb.Window(themename="cosmo")
window.title("EOG 眨眼分析工具（自動門檻支援）")
window.geometry("600x600")

frm = tb.Frame(window, padding=20)
frm.pack(fill=BOTH, expand=True)

tb.Label(frm, text="取樣率 (Hz):").pack(anchor=W)
sampling_rate_entry = tb.Entry(frm)
sampling_rate_entry.insert(0, "1000")
sampling_rate_entry.pack(fill=X, pady=5)

tb.Label(frm, text="EOG 欄位名稱:").pack(anchor=W)
eog_column_entry = tb.Entry(frm)
eog_column_entry.insert(0, "CH2")
eog_column_entry.pack(fill=X, pady=5)

tb.Label(frm, text="眨眼高度門檻 (height_min):（輸入 auto 或留空可自動偵測）").pack(anchor=W)
height_min_entry = tb.Entry(frm)
height_min_entry.insert(0, "auto")
height_min_entry.pack(fill=X, pady=5)

# 區段 1
tb.Label(frm, text="區段 1 - 開始分鐘").pack(anchor=W)
start1_entry = tb.Entry(frm)
start1_entry.insert(0, "5")
start1_entry.pack(fill=X, pady=2)

tb.Label(frm, text="區段 1 - 結束分鐘").pack(anchor=W)
end1_entry = tb.Entry(frm)
end1_entry.insert(0, "10")
end1_entry.pack(fill=X, pady=2)

# 區段 2
tb.Label(frm, text="區段 2 - 開始分鐘").pack(anchor=W)
start2_entry = tb.Entry(frm)
start2_entry.insert(0, "15")
start2_entry.pack(fill=X, pady=2)

tb.Label(frm, text="區段 2 - 結束分鐘").pack(anchor=W)
end2_entry = tb.Entry(frm)
end2_entry.insert(0, "20")
end2_entry.pack(fill=X, pady=2)

tb.Button(frm, text="選擇 CSV 並分析", bootstyle=SUCCESS, command=run_analysis).pack(pady=10, fill=X)

window.mainloop()
