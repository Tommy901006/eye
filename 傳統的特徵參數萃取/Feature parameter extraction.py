import os
import pandas as pd
import numpy as np
from scipy.stats import skew
from scipy.signal import welch

def compute_statistics(df, column):
    # 計算統計量
    mean = df[column].mean()
    median = df[column].median()
    std = df[column].std()
    cv = std / mean if mean != 0 else np.nan
    skewness = skew(df[column].dropna())
    total_power = np.sum(df[column] ** 2)
    
    # 計算中值頻率
    f, Pxx = welch(df[column].dropna(), nperseg=min(256, len(df[column])))
    median_frequency = f[np.argmax(np.cumsum(Pxx) >= np.sum(Pxx) / 2)]

    return mean, median, std, cv, skewness, total_power, median_frequency

# 設定資料夾路徑
folder_path = r"C:\Users\User\Desktop\post\1"

# 建立一個空的 DataFrame 來儲存結果
results = pd.DataFrame(columns=[
    'File Name', 'X Mean', 'X Median', 'X Std', 'X CV', 'X Skewness', 'X Total Power', 'X Median Frequency',
    'Y Mean', 'Y Median', 'Y Std', 'Y CV', 'Y Skewness', 'Y Total Power', 'Y Median Frequency'
])

# 建立一個空的 DataFrame 來儲存結果
results = pd.DataFrame(columns=[
    'File Name', 'X Mean', 'X Median', 'X Std', 'X CV', 'X Skewness', 'X Total Power', 'X Median Frequency',
    'Y Mean', 'Y Median', 'Y Std', 'Y CV', 'Y Skewness', 'Y Total Power', 'Y Median Frequency'
])

# 讀取資料夾中的所有 Excel 檔案
for file in os.listdir(folder_path):
    if file.endswith(".xlsx") or file.endswith(".xls"):
        file_path = os.path.join(folder_path, file)
        df = pd.read_excel(file_path)

        # 計算 X 欄位的統計量
        x_stats = compute_statistics(df, 'X')
        
        # 計算 Y 欄位的統計量
        y_stats = compute_statistics(df, 'Y')

        # 將結果添加到 DataFrame
        new_row = pd.Series([file] + list(x_stats) + list(y_stats), index=results.columns)
        results = pd.concat([results, new_row.to_frame().T], ignore_index=True)

# 輸出結果到新的 Excel 檔案
output_path = "pre-output_statistics.xlsx"
results.to_excel(output_path, index=False)

print(f"統計結果已儲存至 {output_path}")