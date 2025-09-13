import os
import numpy as np
import pandas as pd
from pathlib import Path

# 設定源資料夾和目標資料夾路徑
source_folder = r'C:\Users\user\Desktop\psot'  # 替換為源資料夾的路徑
target_folder = r'C:\Users\user\Desktop\test2'  # 替換為目標資料夾的路徑

# 確保目標資料夾存在
os.makedirs(target_folder, exist_ok=True)

# 線性內插函數
def linear_interpolation(x, x1, y1, x2, y2):
    return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

# 設定上採樣頻率為1200 Hz
sampling_rate = 10  # 每秒1200次採樣
sampling_interval = 1 / sampling_rate  # 採樣間隔（秒）

# 處理每個 Excel 檔案
for file_name in os.listdir(source_folder):
    if file_name.endswith('.xlsx'):
        file_path = os.path.join(source_folder, file_name)
        df = pd.read_excel(file_path, sheet_name=0)  # 讀取第一個工作表

        # 確保數據中包含 'Time', 'X', 和 'Y' 欄位
        if all(col in df.columns for col in ['Time', 'X', 'Y']):
            df_selected = df[['Time', 'X', 'Y']]
            
            # 去除重複的行（相同時間的行只保留一筆）
            df_selected = df_selected.drop_duplicates(subset=['Time'])
            
            # 擷取時間範圍
            start_time = df_selected['Time'].min()
            end_time = df_selected['Time'].max()  # 使用最大時間範圍
            
            # 生成需要內插的時間點
            time_points = np.arange(start_time, end_time, sampling_interval)
            
            # 創建空列表來存儲內插後的數據
            interpolated_values = []
            
            # 進行內插
            for i in range(len(df_selected) - 1):
                row1 = df_selected.iloc[i]
                row2 = df_selected.iloc[i + 1]
                
                x1, y1, y1_y = row1['Time'], row1['X'], row1['Y']
                x2, y2, y2_y = row2['Time'], row2['X'], row2['Y']
                
                segment_points = time_points[(time_points >= x1) & (time_points <= x2)]
                
                for point in segment_points:
                    y_interpolated_X = linear_interpolation(point, x1, y1, x2, y2)
                    y_interpolated_Y = linear_interpolation(point, x1, y1_y, x2, y2_y)
                    interpolated_values.append((point, y_interpolated_X, y_interpolated_Y))
            
            # 創建 DataFrame 以便查看結果
            interpolated_df = pd.DataFrame(interpolated_values, columns=['Time', 'X', 'Y'])
            
            # 將結果寫入目標資料夾中的 Excel 文件
            target_file_path = os.path.join(target_folder, file_name)
            with pd.ExcelWriter(target_file_path) as writer:
                interpolated_df.to_excel(writer, sheet_name='Interpolated_Data', index=False)
            
            print(f"內插結果已成功保存到 '{target_file_path}' 文件中。")
        else:
            print(f"檔案 '{file_name}' 中不包含 'Time', 'X', 和 'Y' 欄位。")