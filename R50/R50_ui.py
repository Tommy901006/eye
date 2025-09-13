import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
from PIL import Image
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Style

# 計算加權平均值
def calculate_weighted_mean(t_values, values):
    weighted_sum = 0
    total_weight = 0
    
    for t, value in zip(t_values, values):
        weighted_sum += t * value
        total_weight += t
    
    if total_weight == 0:
        return 0
    else:
        return weighted_sum / total_weight

# 計算距離
def calculate_distances(x_values, y_values, xcm, ycm):
    distances = []
    for x, y in zip(x_values, y_values):
        distance = math.sqrt((x - xcm)**2 + (y - ycm)**2)
        distances.append(distance)
    return distances

# 處理單個文件
def process_file(file_path, output_folder):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')  # 指定引擎為 'openpyxl'
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return file_path, None, None, None

    # 檢查是否存在必要的列
    required_columns = ['X', 'Y', 'Time']
    if not all(col in df.columns for col in required_columns):
        print(f"Missing required columns in file {file_path}: {', '.join([col for col in required_columns if col not in df.columns])}")
        return file_path, None, None, None

    try:
        x_values = df['X'].tolist()
        y_values = df['Y'].tolist()
        t_values = df['Time'].tolist()
        t_values = [t / 1000 for t in t_values]
    
        ycm = calculate_weighted_mean(t_values, y_values)
        xcm = calculate_weighted_mean(t_values, x_values)
        D_CM = calculate_distances(x_values, y_values, xcm, ycm)
        
        D_CM_sorted = np.sort(D_CM)
        cm_r50 = np.median(D_CM_sorted)
    except Exception as e:
        print(f"Error processing data from file {file_path}: {e}")
        return file_path, None, None, None
    
    # 繪製圖形
    try:
        background_image_path = r'F:\eyetracker sample data_原始\凝視原點.jpg'
        background = Image.open(background_image_path)
    except Exception as e:
        print(f"Error loading background image: {e}")
        return file_path, xcm, ycm, cm_r50

    fig, ax = plt.subplots(figsize=(10, 8))
    ax.imshow(background, extent=[0, 3508, 0, 2480])
    
    ax.scatter(x_values, y_values, c='blue', label='Data Points')
    ax.scatter(xcm, ycm, c='red', label='Center of Mass (CM)')
    ax.scatter(1750, 1250, c='green')
    circle = plt.Circle((xcm, ycm), cm_r50, color='green', fill=False, linestyle='--', label='CM R50')
    ax.add_artist(circle)
    
    ax.set_xlabel('X')
    ax.set_ylabel('Y')
    ax.set_title(f'Data Points with Center of Mass and CM R50\nFile: {os.path.basename(file_path)}')
    ax.legend()
    ax.grid(True)
    ax.invert_yaxis()
    
    # Save the figure
    output_image_path = os.path.join(output_folder, f'{os.path.basename(file_path).replace(".xlsx", ".png")}')
    plt.savefig(output_image_path)
    plt.close()
    
    return os.path.basename(file_path), xcm, ycm, cm_r50

# 處理所有文件
def process_all_files(input_folder, output_folder, output_excel_path, progress_var):
    results = []
    errors = []
    total_files = len([f for f in os.listdir(input_folder) if f.endswith('.xlsx')])
    processed_files = 0

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            result = process_file(file_path, output_folder)
            if None in result:
                errors.append(result[0])
            else:
                results.append(result)
            processed_files += 1
            progress_var.set((processed_files / total_files) * 100)

    # Save results to Excel
    results_df = pd.DataFrame(results, columns=['File Name', 'Xcm', 'Ycm', 'CM R50'])
    results_df.to_excel(output_excel_path, index=False, engine='openpyxl')

    # Save errors to a separate file
    if errors:
        errors_df = pd.DataFrame(errors, columns=['File Name'])
        errors_df.to_excel(os.path.join(output_folder, 'errors.xlsx'), index=False, engine='openpyxl')

# 瀏覽輸入文件夾
def browse_input_folder():
    folder = filedialog.askdirectory()
    input_folder_var.set(folder)

# 瀏覽輸出文件夾
def browse_output_folder():
    folder = filedialog.askdirectory()
    output_folder_var.set(folder)

# 瀏覽輸出Excel文件
def browse_output_excel_path():
    file = filedialog.asksaveasfilename(defaultextension=".xlsx")
    output_excel_var.set(file)

# 開始處理
def start_processing():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()
    output_excel_path = output_excel_var.get()

    if not input_folder or not output_folder or not output_excel_path:
        messagebox.showerror("Error", "Please select all paths.")
        return

    os.makedirs(output_folder, exist_ok=True)

    process_all_files(input_folder, output_folder, output_excel_path, progress_var)
    messagebox.showinfo("Success", "Processing completed.")

# UI設置
root = tk.Tk()
root.title("Data Processing")

# 設定主色調
root.configure(bg='#2E4053')

# 定義變量
input_folder_var = tk.StringVar()
output_folder_var = tk.StringVar()
output_excel_var = tk.StringVar()
progress_var = tk.DoubleVar()

# 自定義樣式
style = Style()
style.theme_use('clam')
style.configure("TButton", background='#34495E', foreground='white', font=('Arial', 10))
style.configure("TLabel", background='#2E4053', foreground='white', font=('Arial', 10))
style.configure("TEntry", font=('Arial', 10))
style.configure("TProgressbar", troughcolor='#34495E', background='#1ABC9C')

# 顯示元件
tk.Label(root, text="Input Folder:").grid(row=0, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=input_folder_var, width=50).grid(row=0, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_input_folder).grid(row=0, column=2, padx=10, pady=5)

tk.Label(root, text="Output Folder:").grid(row=1, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=output_folder_var, width=50).grid(row=1, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_output_folder).grid(row=1, column=2, padx=10, pady=5)

tk.Label(root, text="Output Excel File:").grid(row=2, column=0, padx=10, pady=5, sticky='e')
tk.Entry(root, textvariable=output_excel_var, width=50).grid(row=2, column=1, padx=10, pady=5)
tk.Button(root, text="Browse", command=browse_output_excel_path).grid(row=2, column=2, padx=10, pady=5)

progress_bar = Progressbar(root, variable=progress_var, maximum=100, style="TProgressbar")
progress_bar.grid(row=3, column=0, columnspan=3, pady=10, padx=10, sticky='ew')

tk.Button(root, text="Start Processing", command=start_processing).grid(row=4, column=0, columnspan=3, pady=10)

root.mainloop()
