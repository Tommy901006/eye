import os
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd

# 定義列名
columns = ['GP-No', 'Name', 'Time', 'Timegap', 'X', 'Y', 'X-offset', 'Y-offset',
           'Pupil Diameter Left (mm)', 'Pupil Diameter Right (mm)', 'Mouse X-coord',
           'Mouse Y-coord', 'AoI']

# 選擇來源資料夾
def select_source_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        source_entry.delete(0, tk.END)
        source_entry.insert(0, folder_path)

# 選擇目的地資料夾
def select_destination_folder():
    folder_path = filedialog.askdirectory()
    if folder_path:
        dest_entry.delete(0, tk.END)
        dest_entry.insert(0, folder_path)

# 複製檔案
def copy_files():
    source_folder = source_entry.get()
    dest_folder = dest_entry.get()

    if not source_folder or not dest_folder:
        messagebox.showerror("錯誤", "請選擇來源和目的地資料夾")
        return

    try:
        for file_name in os.listdir(source_folder):
            full_file_name = os.path.join(source_folder, file_name)
            if os.path.isfile(full_file_name):
                shutil.copy(full_file_name, dest_folder)

        messagebox.showinfo("成功", "檔案已成功複製!")
    except Exception as e:
        messagebox.showerror("錯誤", str(e))

# 開始轉換
def start_conversion():
    folder_path = dest_entry.get()
    if not folder_path:
        messagebox.showerror("錯誤", "請選擇目的地資料夾")
        return

    try:
        files = os.listdir(folder_path)
        for file in files:
            if file.endswith('.csv'):
                data = pd.read_csv(os.path.join(folder_path, file), delimiter=';')
                data.to_excel(os.path.join(folder_path, file.replace('.csv', '.xlsx')), index=False)
        messagebox.showinfo("完成", "轉換已執行！")
    except Exception as e:
        messagebox.showerror("錯誤", str(e))


# 檢查檔案
def check():
    folder_path = dest_entry.get()
    if not folder_path:
        messagebox.showerror("錯誤", "請選擇目的地資料夾")
        return

    problem_files = []
    files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    
    for file in files:
        file_path = os.path.join(folder_path, file)
        try:
            df = pd.read_excel(file_path)

            # 檢查是否有 'Pupil Diameter Left (mm)' 這個欄位
            if 'Pupil Diameter Left (mm)' not in df.columns:
                problem_files.append(file)

                # 將數據轉換為字符串類型，然後按空格分列
                split_data = [str(x).split(',') for x in df[df.columns[0]]]
                df2 = pd.DataFrame(split_data, columns=columns)

                # 將數字列轉換為適當的數字類型
                df2 = df2.apply(pd.to_numeric, errors='ignore')

                # 將 DataFrame 保存為新的 Excel 文件
                df2.to_excel(os.path.join(folder_path, file), index=False)
                print(f"轉換完成 {file}")

        except Exception as e:
            print(f"此檔案無須轉換 {file}: {str(e)}")

    if problem_files:
        result_text.delete(1.0, tk.END)  # 清空之前的內容
        result_text.insert(tk.END, "有問題的檔案:\n" + "\n".join(problem_files))
    else:
        messagebox.showinfo("成功", "所有檔案都正確且已轉換!")
        result_text.delete(1.0, tk.END)  # 清空之前的內容

# 計算瞳孔直徑統計
def calculate_pupil_diameter_statistics():
    folder_path = dest_entry.get()
    if not folder_path:
        messagebox.showerror("錯誤", "請選擇目的地資料夾")
        return

    file_paths = [os.path.join(folder_path, f) for f in os.listdir(folder_path) if f.endswith('.xlsx')]
    results = []

    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path)

            left_pupil_diameter = df['Pupil Diameter Left (mm)'].between(2, 8)
            right_pupil_diameter = df['Pupil Diameter Right (mm)'].between(2, 8)

            left_stats = df[left_pupil_diameter]['Pupil Diameter Left (mm)'].describe()
            right_stats = df[right_pupil_diameter]['Pupil Diameter Right (mm)'].describe()

            result = pd.DataFrame({
                '檔案名字': [os.path.basename(file_path), os.path.basename(file_path)],
                '項目名稱': ['Pupil Diameter Left', 'Pupil Diameter Right'],
                'SD': [left_stats['std'], right_stats['std']],
                'MEAN': [left_stats['mean'], right_stats['mean']]
            })

            results.append(result)
        except Exception as e:
            error_message = f"無法計算 {file_path}: {str(e)}\n"
            result_text.insert(tk.END, error_message)

    if results:
        all_results = pd.concat(results)
        all_results.to_excel(os.path.join(folder_path, '瞳孔直徑統計結果.xlsx'), index=False)
        messagebox.showinfo("完成", "瞳孔直徑統計已成功計算並保存！")
    else:
        messagebox.showinfo("結果", "未找到可計算的檔案！")

# 設定主視窗
root = tk.Tk()
root.title("檔案複製器")
root.geometry("600x500")  # 設定視窗大小
root.configure(bg="#f0f0f0")

# 來源資料夾選擇
source_label = tk.Label(root, text="來源資料夾:", bg="#f0f0f0")
source_label.pack(pady=5)
source_entry = tk.Entry(root, width=70)
source_entry.pack(pady=5)
source_button = tk.Button(root, text="選擇來源資料夾", command=select_source_folder, bg="#4CAF50", fg="white")
source_button.pack(pady=5)

# 目的地資料夾選擇
dest_label = tk.Label(root, text="目的地資料夾:", bg="#f0f0f0")
dest_label.pack(pady=5)
dest_entry = tk.Entry(root, width=70)
dest_entry.pack(pady=5)
dest_button = tk.Button(root, text="選擇目的地資料夾", command=select_destination_folder, bg="#4CAF50", fg="white")
dest_button.pack(pady=5)

# 複製檔案按鈕
copy_button = tk.Button(root, text="複製檔案  Step 1", command=copy_files, bg="#2196F3", fg="white")
copy_button.pack(pady=10)

# 開始轉換按鈕
convert_button = tk.Button(root, text="開始轉換 Excel  Step 2", command=start_conversion, bg="#2196F3", fg="white")
convert_button.pack(pady=10)

# 資料確認按鈕
convert_button1 = tk.Button(root, text="資料確認  Step 3", command=check, bg="#2196F3", fg="white")
convert_button1.pack(pady=10)

# 計算瞳孔直徑按鈕
calculate_button = tk.Button(root, text="計算瞳孔大小統計  Step 4", command=calculate_pupil_diameter_statistics, bg="#2196F3", fg="white")
calculate_button.pack(pady=10)

# 結果顯示區域
result_text = scrolledtext.ScrolledText(root, height=15, width=70, bg="#ffffff")
result_text.pack(pady=10)

# 開始主循環
root.mainloop()
