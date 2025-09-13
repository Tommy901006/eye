import os
import numpy as np
import pandas as pd
from pathlib import Path

source_folder = r'C:\Users\User\Desktop\2'  # 替換為源資料夾的路徑
target_folder = r'C:\Users\user\Desktop'  # 替換為目標資料夾的路徑

# 確保目標資料夾存在
os.makedirs(target_folder, exist_ok=True)

# 線性內插函數
def linear_interpolation(x, x1, y1, x2, y2):
    return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

# 設定上採樣頻率為10 Hz
sampling_rate = 1  # 每秒10次採樣
sampling_interval = 1 / sampling_rate  # 採樣間隔（秒）

# 處理每個 Excel 檔案
for file_name in os.listdir(source_folder):
    if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
        file_path = os.path.join(source_folder, file_name)
        
        try:
            df = pd.read_excel(file_path, sheet_name=0)  # 讀取第一個工作表

            # 確保數據中包含 'Time', 'Pupil Diameter Left (mm)', 和 'Pupil Diameter Right (mm)' 欄位
            if all(col in df.columns for col in ['Time', 'Pupil Diameter Left (mm)', 'Pupil Diameter Right (mm)']):
                df_selected = df[['Time', 'Pupil Diameter Left (mm)', 'Pupil Diameter Right (mm)']]
                
                # 去除重複的行（相同時間的行只保留一筆）
                df_selected = df_selected.drop_duplicates(subset=['Time'])
                
                # 轉換時間為數值型，排除無效的時間點
                df_selected['Time'] = pd.to_numeric(df_selected['Time'], errors='coerce')
                df_selected = df_selected.dropna(subset=['Time'])
                
                # 確保轉換後的時間範圍是數值型
                start_time = df_selected['Time'].min()
                end_time = df_selected['Time'].max()
                
                # 生成需要內插的時間點
                time_points = np.arange(start_time, end_time, sampling_interval)
                
                # 創建空字典來存儲內插後的數據，根據時間點來合併左右眼數據
                interpolated_values = {point: [None, None] for point in time_points}
                
                # 過濾數據
                df_filtered_left = df_selected.loc[(df_selected['Pupil Diameter Left (mm)'] >= 2) & (df_selected['Pupil Diameter Left (mm)'] <= 8)]
                df_filtered_right = df_selected.loc[(df_selected['Pupil Diameter Right (mm)'] >= 2) & (df_selected['Pupil Diameter Right (mm)'] <= 8)]

                # 重新設置時間列
                df_filtered_left.loc[:, 'Time'] = df_filtered_left['Time'] - df_filtered_left['Time'].min()
                df_filtered_right.loc[:, 'Time'] = df_filtered_right['Time'] - df_filtered_right['Time'].min()

                # 進行內插 - Pupil Diameter Left (mm)
                for i in range(len(df_filtered_left) - 1):
                    row1 = df_filtered_left.iloc[i]
                    row2 = df_filtered_left.iloc[i + 1]
                    
                    x1, y1 = row1['Time'], row1['Pupil Diameter Left (mm)']
                    x2, y2 = row2['Time'], row2['Pupil Diameter Left (mm)']
                    
                    segment_points = time_points[(time_points >= x1) & (time_points <= x2)]
                    
                    for point in segment_points:
                        x_interpolated = linear_interpolation(point, x1, y1, x2, y2)
                        interpolated_values[point][0] = x_interpolated
                
                # 進行內插 - Pupil Diameter Right (mm)
                for i in range(len(df_filtered_right) - 1):
                    row1 = df_filtered_right.iloc[i]
                    row2 = df_filtered_right.iloc[i + 1]
                    
                    x1, y1 = row1['Time'], row1['Pupil Diameter Right (mm)']
                    x2, y2 = row2['Time'], row2['Pupil Diameter Right (mm)']
                    
                    segment_points = time_points[(time_points >= x1) & (time_points <= x2)]
                    
                    for point in segment_points:
                        y_interpolated = linear_interpolation(point, x1, y1, x2, y2)
                        interpolated_values[point][1] = y_interpolated
                
                # 創建 DataFrame 以便查看結果
                interpolated_df = pd.DataFrame(
                    [(time, left, right) for time, (left, right) in interpolated_values.items()],
                    columns=['Time', 'X', 'Y']
                )
                
                # 去除重複的行（相同時間的行只保留一筆）
                interpolated_df = interpolated_df.drop_duplicates(subset=['Time'])
                
                # 將結果寫入目標資料夾中的 Excel 文件
                target_file_path = os.path.join(target_folder, file_name)
                with pd.ExcelWriter(target_file_path) as writer:
                    interpolated_df.to_excel(writer, sheet_name='Interpolated_Data', index=False)
                
                print(f"內插結果已成功保存到 '{target_file_path}' 文件中。")
            else:
                print(f"檔案 '{file_name}' 中不包含 'Time', 'Pupil Diameter Left (mm)', 和 'Pupil Diameter Right (mm)' 欄位。")
        
        except Exception as e:
            print(f"處理檔案 '{file_name}' 時發生錯誤：{e}")
