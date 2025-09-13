import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
from PIL import Image
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Function to calculate weighted mean
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

# Function to calculate distances between points and center of mass
def calculate_distances(x_values, y_values, xcm, ycm):
    distances = []
    for x, y in zip(x_values, y_values):
        distance = math.sqrt((x - xcm)**2 + (y - ycm)**2)
        distances.append(distance)
    return distances

# Function to calculate the distance between a fixed point (1750, 1250) and the center of mass (Xcm, Ycm)
def calculate_distance_to_fixed_point(xcm, ycm, fixed_x=1750, fixed_y=1250):
    return math.sqrt((xcm - fixed_x)**2 + (ycm - fixed_y)**2)

# Function to process a single file and generate visualization
def process_file(file_path, background_image_path, output_folder):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')  # Specify engine as 'openpyxl'
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return file_path, None, None, None, None

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
        
        # Calculate distance from (1750, 1250) to (Xcm, Ycm)
        distance_to_fixed_point = calculate_distance_to_fixed_point(xcm, ycm)
        
    except Exception as e:
        print(f"Error processing data from file {file_path}: {e}")
        return file_path, None, None, None, None
    
    # Plotting
    try:
        background = Image.open(background_image_path)
    except Exception as e:
        print(f"Error loading background image: {e}")
        return file_path, xcm, ycm, cm_r50, distance_to_fixed_point

    fig, ax = plt.subplots(figsize=(10, 8))
    ax.imshow(background, extent=[0, 3508, 0, 2480])
    
    ax.scatter(x_values, y_values, c='blue', label='Data Points', alpha=0.1)
    ax.scatter(xcm, ycm, c='red', label='Center of Mass (CM)')
    ax.scatter(1750, 1250, c='green', label='Fixed Point (1750, 1250)')
    
    # Draw line between CM and fixed point
    ax.plot([xcm, 1750], [ycm, 1250], color='purple', linestyle='-', label='Distance to Fixed Point')

    # Annotate distance
    mid_x = (xcm + 1750) / 2
    mid_y = (ycm + 1250) / 2
    ax.text(mid_x + 150, mid_y, f'd={distance_to_fixed_point:.2f}', color='purple', fontsize=14, fontweight='bold', 
            bbox=dict(facecolor='white', alpha=0.6, edgecolor='purple', boxstyle='round,pad=0.5'))

    # Plot CM R50 circle
    circle = plt.Circle((xcm, ycm), cm_r50, color='yellow', fill=False, linestyle='--', linewidth=2.5, label='CM R50')
    ax.add_artist(circle)

    # Annotate CM R50 value
    ax.text(xcm - 150, ycm + 300, f'cm_r50={cm_r50:.2f}', color='purple', fontsize=14, fontweight='bold',
            bbox=dict(facecolor='white', alpha=0.6, edgecolor='purple', boxstyle='round,pad=0.5'))

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
    
    return os.path.basename(file_path), xcm, ycm, cm_r50, distance_to_fixed_point

# Function to process all files in a folder and save results with distances
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
    
    # Save results to Excel with additional 'distance' column
    results_df = pd.DataFrame(results, columns=['File Name', 'Xcm', 'Ycm', 'CM R50', 'Distance'])
    results_df.to_excel(output_excel_path, index=False, engine='openpyxl')

    # Save errors to a separate file
    if errors:
        errors_df = pd.DataFrame(errors, columns=['File Name'])
        errors_df.to_excel(os.path.join(output_folder, 'errors.xlsx'), index=False, engine='openpyxl')

# GUI Functions
def select_input_folder():
    folder_selected = filedialog.askdirectory()
    input_folder_entry.delete(0, tk.END)
    input_folder_entry.insert(0, folder_selected)

def select_background_image():
    file_selected = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png")])
    background_image_entry.delete(0, tk.END)
    background_image_entry.insert(0, file_selected)

def select_output_folder():
    folder_selected = filedialog.askdirectory()
    output_folder_entry.delete(0, tk.END)
    output_folder_entry.insert(0, folder_selected)

def run_processing():
    input_folder = input_folder_entry.get()
    background_image_path = background_image_entry.get()
    output_folder = output_folder_entry.get()
    output_excel_path = os.path.join(output_folder, 'post-R50.xlsx')
    
    if not input_folder or not background_image_path or not output_folder:
        messagebox.showerror("Input Error", "Please select all required folders and files.")
        return
    
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Run the processing
    process_all_files(input_folder, background_image_path, output_folder, output_excel_path)
    messagebox.showinfo("Success", "Processing completed successfully!")

# GUI Setup
root = tk.Tk()
root.title("Data Processing Tool")

# Input folder
input_folder_label = tk.Label(root, text="Input Folder:")
input_folder_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
input_folder_entry = tk.Entry(root, width=50)
input_folder_entry.grid(row=0, column=1, padx=10, pady=10)
input_folder_button = tk.Button(root, text="Browse", command=select_input_folder)
input_folder_button.grid(row=0, column=2, padx=10, pady=10)

# Background image
background_image_label = tk.Label(root, text="Background Image:")
background_image_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")
background_image_entry = tk.Entry(root, width=50)
background_image_entry.grid(row=1, column=1, padx=10, pady=10)
background_image_button = tk.Button(root, text="Browse", command=select_background_image)
background_image_button.grid(row=1, column=2, padx=10, pady=10)

# Output folder
output_folder_label = tk.Label(root, text="Output Folder:")
output_folder_label.grid(row=2, column=0, padx=10, pady=10, sticky="e")
output_folder_entry = tk.Entry(root, width=50)
output_folder_entry.grid(row=2, column=1, padx=10, pady=10)
output_folder_button = tk.Button(root, text="Browse", command=select_output_folder)
output_folder_button.grid(row=2, column=2, padx=10, pady=10)

# Run button
run_button = tk.Button(root, text="Run Processing", command=run_processing, width=20)
run_button.grid(row=3, column=0, columnspan=3, pady=20)

root.mainloop()
