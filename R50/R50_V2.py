import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
from PIL import Image
def calculate_weighted_mean(t_values, values):
    # 初始化加權值的總和和權重的總和
    weighted_sum = 0
    total_weight = 0
    
    # 遍歷 t_values 和 values，計算加權值的總和和權重的總和
    for t, value in zip(t_values, values):
        weighted_sum += t * value
        total_weight += t
    
    # 計算加權平均值
    if total_weight == 0:
        return 0
    else:
        weighted_mean = weighted_sum / total_weight
        return weighted_mean
    
def calculate_distances(x_values, y_values, xcm, ycm):
    distances = []  # 用來存儲每個點的距離
    for x, y in zip(x_values, y_values):  # 同時遍歷 x_values 和 y_values
        # 計算每個點到中心點的距離
        distance = math.sqrt((x - xcm)**2 + (y - ycm)**2)
        distances.append(distance)  # 將距離加入列表
    return distances  # 返回所有距離

file_path = r'F:\eyetracker sample data\entropy\Approximate Entropy\1\post-p3-1.xlsx' 
df = pd.read_excel(file_path)

# 提取 X, Y 和 Time 列
x_values = df['X'].tolist()
y_values = df['Y'].tolist()
t_values = df['Time'].tolist()
t_values = [t / 1000 for t in t_values]
# 計算 Ycm 和 Xcm
ycm = calculate_weighted_mean(t_values, y_values)
xcm = calculate_weighted_mean(t_values, x_values)
D_CM = calculate_distances(x_values,y_values,xcm,ycm)
# 打印結果
print(f"Ycm: {ycm}")
print(f"Xcm: {xcm}")
# 計算中位數距離 CM_R50
# 將 D_CM 按升序重新排列
D_CM_sorted = np.sort(D_CM)
cm_r50 = np.median(D_CM_sorted)

# 打印結果
print(f"CM R50 : {cm_r50:.2f}")


# 加載背景圖片
background_image_path = r'F:\eyetracker sample data_原始\凝視原點.jpg'
background = Image.open(background_image_path)

# 繪製圖形
fig, ax = plt.subplots(figsize=(10, 8))
ax.imshow(background, extent=[0, 3508, 0, 2480])  # 設置背景圖片


# 繪製數據點和中心點
ax.scatter(x_values, y_values, c='blue', label='Data Points')
ax.scatter(xcm, ycm, c='red', label='Center of Mass (CM)')
ax.scatter(1750, 1250, c='green')
circle = plt.Circle((xcm, ycm), cm_r50, color='green', fill=False, linestyle='--', label='CM R50')
ax.add_artist(circle)

ax.set_xlabel('X')
ax.set_ylabel('Y')
ax.set_title('Data Points with Center of Mass and CM R50')
ax.legend()
ax.grid(True)
ax.invert_yaxis()  # 顛倒 Y 軸
plt.show()