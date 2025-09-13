import pandas as pd
import math
import matplotlib.pyplot as plt

# 讀取 Excel 檔案
file_path = r'F:\eyetracker sample data\R50\R50.xlsx'  # 替換為你的檔案路徑
df = pd.read_excel(file_path)

# 計算距離並新增到 DataFrame 中
def calculate_distance(row):
    return math.sqrt((row['new_X'] - row['X'])**2 + (row['new_Y'] - row['Y'])**2)

df['distance'] = df.apply(calculate_distance, axis=1)

# 繪製原始座標和新座標
plt.figure(figsize=(10, 6))
plt.scatter(df['X'], df['Y'], color='blue', label='Original Coordinates')
plt.scatter(df['new_X'], df['new_Y'], color='red', label='New Coordinates')

# 繪製每一對點之間的線條
for i in range(len(df)):
    plt.plot([df['X'][i], df['new_X'][i]], [df['Y'][i], df['new_Y'][i]], 'gray', linestyle=':')

# 添加標籤和圖例
plt.xlabel('X Coordinate')
plt.ylabel('Y Coordinate')
plt.title('Original vs New Coordinates with Distance')
plt.legend()

# 顯示圖表
plt.show()

# 儲存結果到新的 Excel 檔案
output_file = 'output_with_distances.xlsx'
df.to_excel(output_file, index=False)

print("距離已經計算並儲存在新的檔案中。")

