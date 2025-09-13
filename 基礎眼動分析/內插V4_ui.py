import os
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


# Linear interpolation function
def linear_interpolation(x, x1, y1, x2, y2):
    return y1 + (x - x1) * (y2 - y1) / (x2 - x1)

# Function to process files and perform interpolation
def process_files(source_folder, target_folder):
    # Ensure target folder exists
    os.makedirs(target_folder, exist_ok=True)

    # Set up sampling rate and interval
    sampling_rate = 1  # 10 Hz
    sampling_interval = 1 / sampling_rate  # Sampling interval in seconds

    # Process each Excel file in the source folder
    for file_name in os.listdir(source_folder):
        if file_name.endswith('.xlsx') and not file_name.startswith('~$'):
            file_path = os.path.join(source_folder, file_name)
            
            try:
                df = pd.read_excel(file_path, sheet_name=0)  # Read the first sheet

                # Check if required columns are present
                if all(col in df.columns for col in ['Time', 'X', 'Y']):
                    df_selected = df[['Time', 'X', 'Y']]
                    df_selected = df_selected.drop_duplicates(subset=['Time'])
                    
                    # Convert time to numeric and drop invalid entries
                    df_selected['Time'] = pd.to_numeric(df_selected['Time'], errors='coerce')
                    df_selected = df_selected.dropna(subset=['Time'])

                    # Get time range
                    start_time = df_selected['Time'].min()
                    end_time = df_selected['Time'].max()
                    time_points = np.arange(start_time, end_time, sampling_interval)

                    interpolated_values = []

                    # Interpolation process
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
                    
                    # Create DataFrame for interpolated data
                    interpolated_df = pd.DataFrame(interpolated_values, columns=['Time', 'X', 'Y'])
                    interpolated_df = interpolated_df.drop_duplicates(subset=['Time'])

                    # Save results to target folder
                    target_file_path = os.path.join(target_folder, file_name)
                    with pd.ExcelWriter(target_file_path) as writer:
                        interpolated_df.to_excel(writer, sheet_name='Interpolated_Data', index=False)
                    
                    print(f"Interpolated results saved to '{target_file_path}'.")
                else:
                    print(f"File '{file_name}' does not contain required columns.")
            
            except Exception as e:
                print(f"Error processing file '{file_name}': {e}")

    messagebox.showinfo("Done", "Interpolation process completed!")


# Function to select source folder
def select_source_folder():
    folder = filedialog.askdirectory()
    if folder:
        source_folder_entry.delete(0, tk.END)
        source_folder_entry.insert(0, folder)

# Function to select target folder
def select_target_folder():
    folder = filedialog.askdirectory()
    if folder:
        target_folder_entry.delete(0, tk.END)
        target_folder_entry.insert(0, folder)

# Function to run the interpolation process
def run_interpolation():
    source_folder = source_folder_entry.get()
    target_folder = target_folder_entry.get()
    
    if not source_folder or not target_folder:
        messagebox.showwarning("Input Error", "Please select both source and target folders.")
        return
    
    process_files(source_folder, target_folder)


# Set up the GUI
root = tk.Tk()
root.title("Linear Interpolation Tool")

# Source folder selection
tk.Label(root, text="Source Folder:").grid(row=0, column=0, padx=10, pady=10)
source_folder_entry = tk.Entry(root, width=50)
source_folder_entry.grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_source_folder).grid(row=0, column=2, padx=10, pady=10)

# Target folder selection
tk.Label(root, text="Target Folder:").grid(row=1, column=0, padx=10, pady=10)
target_folder_entry = tk.Entry(root, width=50)
target_folder_entry.grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Browse", command=select_target_folder).grid(row=1, column=2, padx=10, pady=10)

# Run button
tk.Button(root, text="Run Interpolation", command=run_interpolation).grid(row=2, column=1, padx=10, pady=20)

# Run the GUI loop
root.mainloop()
