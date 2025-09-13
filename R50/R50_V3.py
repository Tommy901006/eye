import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
from PIL import Image
import os

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

def calculate_distances(x_values, y_values, xcm, ycm):
    distances = []
    for x, y in zip(x_values, y_values):
        distance = math.sqrt((x - xcm)**2 + (y - ycm)**2)
        distances.append(distance)
    return distances

def process_file(file_path, background_image_path, output_folder):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')  # 指定引擎為 'openpyxl'
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return file_path, None, None, None

    x_values = df['X'].tolist()
    y_values = df['Y'].tolist()
    t_values = df['Time'].tolist()
    t_values = [t / 1000 for t in t_values]
    
    try:
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

def process_all_files(input_folder, background_image_path, output_folder, output_excel_path):
    results = []
    errors = []

    for file_name in os.listdir(input_folder):
        if file_name.endswith('.xlsx'):
            file_path = os.path.join(input_folder, file_name)
            result = process_file(file_path, background_image_path, output_folder)
            if None in result:
                errors.append(result[0])
            else:
                results.append(result)
    
    # Save results to Excel
    results_df = pd.DataFrame(results, columns=['File Name', 'Xcm', 'Ycm', 'CM R50'])
    results_df.to_excel(output_excel_path, index=False, engine='openpyxl')

    # Save errors to a separate file
    if errors:
        errors_df = pd.DataFrame(errors, columns=['File Name'])
        errors_df.to_excel(os.path.join(output_folder, 'errors.xlsx'), index=False, engine='openpyxl')

input_folder = r'E:\eyetracker sample data\取樣後 date\post'
background_image_path = r'F:\eyetracker sample data_原始\凝視原點.jpg'
output_folder = r'E:\eyetracker sample data\取樣後 date\post-eyetracker_output_images'
output_excel_path = r'E:\eyetracker sample data\取樣後 date\post-R50.xlsx'

os.makedirs(output_folder, exist_ok=True)

process_all_files(input_folder, background_image_path, output_folder, output_excel_path)
