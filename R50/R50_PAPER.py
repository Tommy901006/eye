import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math
from PIL import Image
import os

# Function to calculate weighted mean
def calculate_weighted_mean(t_values, values):
    weighted_sum = 0
    total_weight = 0
    for t, value in zip(t_values, values):
        weighted_sum += t * value
        total_weight += t
    
    return weighted_sum / total_weight if total_weight != 0 else 0

# Function to calculate distances between points
def calculate_distances(x_values, y_values, xcm, ycm):
    distances = [math.sqrt((x - xcm)**2 + (y - ycm)**2) for x, y in zip(x_values, y_values)]
    return distances

# Function to process each file and plot results
def process_file(file_path, background_image_path, output_folder):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
    except Exception as e:
        print(f"Error reading file {file_path}: {e}")
        return file_path, None, None, None
    
    # Print column names for debugging
    print(f"Columns in {file_path}: {df.columns}")
    
    try:
        x_values = df['X'].tolist()
        y_values = df['Y'].tolist()
        t_values = [t / 1000 for t in df['Time'].tolist()]  # Convert milliseconds to seconds
        
        # Calculate center of mass and CM R50
        ycm = calculate_weighted_mean(t_values, y_values)
        xcm = calculate_weighted_mean(t_values, x_values)
        D_CM = calculate_distances(x_values, y_values, xcm, ycm)
        cm_r50 = np.median(np.sort(D_CM))
    except Exception as e:
        print(f"Error processing data from file {file_path}: {e}")
        return file_path, None, None, None
    
    # Load background image
    try:
        background = Image.open(background_image_path)
    except Exception as e:
        print(f"Error loading background image: {e}")
        return file_path, xcm, ycm, cm_r50

    # Create plot with background and data points
    fig, ax = plt.subplots(figsize=(10, 8))
    ax.imshow(background, extent=[0, 3508, 0, 2480])
    
    # Scatter plot for data points, CM, and specific point
    ax.scatter(x_values, y_values, c='blue', alpha=0.1,label='Data Points')
    ax.scatter(xcm, ycm, c='red', label='Center of Mass (CM)')
    ax.scatter(1750, 1250, c='green')
    
    # Draw a line between (xcm, ycm) and (1750, 1250) and label it as 'd'
    ax.plot([xcm, 1750], [ycm, 1250], color='purple', linestyle='-', label='Line d')

    # Annotate the distance 'd' on the line
    mid_x = (xcm + 1750) / 2
    mid_y = (ycm + 1250) / 2
    distance_d = math.sqrt((xcm - 1750) ** 2 + (ycm - 1250) ** 2)
    # Make the CM R50 circle more prominent by adjusting linewidth and color
    circle = plt.Circle((xcm, ycm), cm_r50, 
                        color='yellow', 
                        fill=False, 
                        linestyle='--', 
                        linewidth=2.5,   # Increase the line width to make the circle clearer
                        label='CM R50')  # Add label
    ax.add_artist(circle)
    # Annotate the cm_r50 value on the plot
    ax.text(xcm - 150, ycm + 250,  # Adjust position for better visibility
            f'cm_r50={cm_r50:.2f}', # Format cm_r50 to two decimal places
            color='purple', 
            fontsize=14, 
            fontweight='bold', 
            bbox=dict(facecolor='white', alpha=0.6, edgecolor='purple', boxstyle='round,pad=0.5'))  # Background box for clarity

    # Annotate the distance 'd' on the line, with more prominent styling
    ax.text(mid_x + 150, mid_y, 
            f'd={distance_d:.2f}', 
            color='purple', 
            fontsize=14,               # Increase font size
            fontweight='bold',         # Make the text bold
            bbox=dict(facecolor='white', alpha=0.6, edgecolor='purple', boxstyle='round,pad=0.5'))  # Add a white background box with padding


    # Customize plot
    ax.set_xlabel('X')
    ax.set_ylabel('Y')
    ax.set_title(f'Data Points with Center of Mass, CM R50, and Distance d\nFile: {os.path.basename(file_path)}')
    ax.legend()
    ax.grid(True)
    ax.invert_yaxis()
    
    # Save the plot
    output_image_path = os.path.join(output_folder, f'{os.path.basename(file_path).replace(".xlsx", ".png")}')
    plt.savefig(output_image_path)
    plt.close()
    
    return os.path.basename(file_path), xcm, ycm, cm_r50

# Function to process all files in the folder
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

    # Save errors if any
    if errors:
        errors_df = pd.DataFrame(errors, columns=['File Name'])
        errors_df.to_excel(os.path.join(output_folder, 'errors.xlsx'), index=False, engine='openpyxl')

# Function to calculate distances for R50 and plot results
def calculate_distance(row):
    return math.sqrt((row['new_X'] - row['X'])**2 + (row['new_Y'] - row['Y'])**2)

def process_r50_file(file_path):
    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        print(f"Error reading R50 file: {e}")
        return

    # Calculate distances between original and new coordinates
    df['distance'] = df.apply(calculate_distance, axis=1)

    # Plot original vs new coordinates
    plt.figure(figsize=(10, 6))
    plt.scatter(df['X'], df['Y'], color='blue', label='Original Coordinates')
    plt.scatter(df['new_X'], df['new_Y'], color='red', label='New Coordinates')

    # Draw lines between original and new points
    for i in range(len(df)):
        plt.plot([df['X'][i], df['new_X'][i]], [df['Y'][i], df['new_Y'][i]], 'gray', linestyle=':')
    
    # Customize plot
    plt.xlabel('X Coordinate')
    plt.ylabel('Y Coordinate')
    plt.title('Original vs New Coordinates with Distance')
    plt.legend()
    plt.show()

    # Save the modified data with distances
    output_file = 'output_with_distances.xlsx'
    df.to_excel(output_file, index=False)
    print("Distances calculated and saved in:", output_file)

# Example usage
input_folder = r'C:\Users\User\Desktop\T'
background_image_path = r'F:\eyetracker sample data_原始\凝視原點.jpg'
output_folder = r'C:\Users\User\Desktop\T\post-eyetracker_output_images'
output_excel_path = r'C:\Users\User\Desktop\T\R50\post-R50.xlsx'
r50_file_path = r'C:\Users\User\Desktop\T\R50\R50.xlsx'

# Ensure output folder exists
os.makedirs(output_folder, exist_ok=True)

# Process eyetracker data and generate images
process_all_files(input_folder, background_image_path, output_folder, output_excel_path)

# Process R50 file and calculate distances
process_r50_file(r50_file_path)
