import numpy as np
import pandas as pd

# 讀取 Excel 文件
file_path = r'D:\eyetracker sample data_原始\eye\chen\post\gaze\post-p1.xlsx'  # 請替換為你的 Excel 文件路徑
df = pd.read_excel(file_path, sheet_name=0)  # sheet_name=0 表示讀取第一個工作表

# 確保數據中包含 'Time' 和 'X' 欄位
df_selected = df[['Time', 'X']]

# 線性內插函數
def linear_interpolation(x, x1, y1, x2, y2):
    return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

# 設定上採樣頻率為1200 Hz
sampling_rate = 10  # 每秒1200次採樣
sampling_interval = 1 / sampling_rate  # 採樣間隔（秒）

# 擷取時間範圍
start_time = df_selected['Time'].min()
end_time = df_selected['Time'].max()

# 生成需要內插的時間點
time_points = np.arange(start_time, end_time, sampling_interval)

# 創建一個空的列表來存儲內插後的數據
interpolated_values = []

# 進行內插
for i in range(len(df_selected) - 1):
    x1, y1 = df_selected.iloc[i]
    x2, y2 = df_selected.iloc[i + 1]
    
    # 選擇當前段落的時間點
    segment_points = time_points[(time_points >= x1) & (time_points <= x2)]
    
    # 對每個時間點進行內插
    for point in segment_points:
        y_interpolated = linear_interpolation(point, x1, y1, x2, y2)
        interpolated_values.append((point, y_interpolated))

# 創建 DataFrame 以便查看結果
interpolated_df = pd.DataFrame(interpolated_values, columns=['Time', 'X'])
output_file_path = 'interpolated_data.csv'  # 替換為你想要的輸出文件路徑
interpolated_df.to_csv(output_file_path, index=False)

print(f'內插數據已輸出到 {output_file_path}')
