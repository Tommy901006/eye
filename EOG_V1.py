import neurokit2 as nk
import pandas as pd
import matplotlib.pyplot as plt
from scipy.signal import stft
import numpy as np
from matplotlib import font_manager
font_path = 'C:/Windows/Fonts/msjh.ttc'  # 替換成你的中文字體路徑
font_prop = font_manager.FontProperties(fname=font_path)
plt.rcParams['font.family'] = font_prop.get_name()
plt.rcParams['axes.unicode_minus'] = False  # Ensure minus signs are displayed

# 讀取 CSV 文件
file_path = r"E:\EOG\轉換完檔案\semicolon_VR-P4.csv"  # 替換成你的 CSV 檔案路徑
data = pd.read_csv(file_path)

# 提取 CH1 和 CH2 欄位
ch1 = data["ch1"]
ch2 = data["ch2"]

# 載入原始波形
nk.signal_plot(ch1)
plt.title("CH1 原始波形")
plt.show()

nk.signal_plot(ch2)
plt.title("CH2 原始波形")
plt.show()

# 放大特定區域
nk.signal_plot(ch1[100:1700])
plt.title("CH1 放大區域")
plt.show()

# 過濾訊號以消除一些噪音和趨勢
eog_cleaned = nk.eog_clean(ch1, sampling_rate=100, method='neurokit')

# 比較原始訊號和過濾後的訊號
nk.signal_plot([ch1, eog_cleaned], 
               labels=["原始訊號", "過濾後訊號"])
plt.title("原始訊號 vs 過濾後訊號")
plt.show()

# 找到眨眼事件
blinks = nk.eog_findpeaks(eog_cleaned, sampling_rate=100, method="mne")
print("眨眼事件:", blinks)

# 可視化眨眼事件位置
blink_peaks = blinks  # 不需要進一步索引，因為它是 numpy 陣列

# 可視化眨眼事件
nk.events_plot(blink_peaks, eog_cleaned)
plt.title("眨眼事件位置")
plt.show()

