import numpy as np
import pandas as pd

# 讀取 Excel 文件
file_path = r'D:\eyetracker sample data_原始\eye\chen\post\gaze\post-p1.xlsx'  # 請替換為你的 Excel 文件路徑
df = pd.read_excel(file_path, sheet_name=0)  # 讀取第一個工作表

# 確保數據中包含 'Time', 'X', 和 'Y' 欄位
if all(col in df.columns for col in ['Time', 'X', 'Y']):
    df_selected = df[['Time', 'X', 'Y']]
    print("數據讀取成功，包含 'Time', 'X', 和 'Y' 欄位。")
else:
    raise ValueError("錯誤：Excel 文件中不包含 'Time', 'X', 和 'Y' 欄位。")

# 線性內插函數
def linear_interpolation(x, x1, y1, x2, y2):
    return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

# 設定上採樣頻率為1200 Hz
sampling_rate = 10  # 每秒1200次採樣
sampling_interval = 1 / sampling_rate  # 採樣間隔（秒）

# 擷取時間範圍
start_time = df['Time'].min()
end_time = df['Time'].max()  # 使用最大時間範圍

# 生成需要內插的時間點
time_points = np.arange(start_time, end_time, sampling_interval)

# 創建空列表來存儲內插後的數據
interpolated_values = []

# 進行內插
for i in range(len(df) - 1):
    # 使用 .iloc 確保獲取單行數據
    row1 = df.iloc[i]
    row2 = df.iloc[i + 1]
    
    # 獲取當前段落的時間點和數值
    x1, y1, y1_y = row1['Time'], row1['X'], row1['Y']
    x2, y2, y2_y = row2['Time'], row2['X'], row2['Y']
    
    # 選擇當前段落的時間點
    segment_points = time_points[(time_points >= x1) & (time_points <= x2)]
    
    # 對每個時間點進行內插
    for point in segment_points:
        y_interpolated_X = linear_interpolation(point, x1, y1, x2, y2)
        y_interpolated_Y = linear_interpolation(point, x1, y1_y, x2, y2_y)
        interpolated_values.append((point, y_interpolated_X, y_interpolated_Y))

# 創建 DataFrame 以便查看結果
interpolated_df = pd.DataFrame(interpolated_values, columns=['Time', 'X', 'Y'])

# 將結果寫入 Excel 文件
output_file_path = 'interpolated_data.xlsx'
with pd.ExcelWriter(output_file_path) as writer:
    interpolated_df.to_excel(writer, sheet_name='Interpolated_Data', index=False)

print(f"內插結果已成功保存到 '{output_file_path}' 文件中。")

# 讀取 Excel 文件中的內插數據
df_interpolated = pd.read_excel(output_file_path, sheet_name='Interpolated_Data')

# 顯示數據
print("Interpolated Data:")
print(df_interpolated.head())
