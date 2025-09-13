# -*- coding: utf-8 -*-
"""
Created on Thu Aug  1 08:41:56 2024

@author: User
"""

import pandas as pd
import numpy as np

# 讀取 Excel 文件
df = pd.read_excel(r'F:\eyetracker sample data\entropy\Approximate Entropy\1\post-p1.xlsx' )

# 假設 X 和 Y 列的標題分別為 'X' 和 'Y'
x = df['X'].to_numpy()
y = df['Y'].to_numpy()

# 計算質心 (CM) 的座標
x_cm = np.mean(x)
y_cm = np.mean(y)

# 計算每個瞳孔軌跡與質心之間的距離 D_CM
D_CM = np.sqrt((x - x_cm)**2 + (y - y_cm)**2)

# 將 D_CM 按升序重新排列
D_CM_sorted = np.sort(D_CM)

# 計算焦點半徑 (focus radius)，即距離的中位數
focus_radius = np.median(D_CM_sorted)

print(f"x_cm: {x_cm}, y_cm: {y_cm}")
print(f"Focus Radius: {focus_radius}")
