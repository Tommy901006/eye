import tkinter as tk
import subprocess
import os
from tkinter import messagebox

# Function to run the specified script
def run_script(script_name):
    try:
        # Get the current working directory (where the script is located)
        script_path = os.path.join(os.getcwd(), script_name)
        subprocess.run(["python", script_path], check=True)
        messagebox.showinfo("Success", f"{script_name} has been executed successfully.")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("Error", f"An error occurred while executing {script_name}: {e}")
    except FileNotFoundError:
        messagebox.showerror("Error", f"{script_name} not found.")

# Set up the GUI
root = tk.Tk()
root.title("Script Runner")
root.geometry("400x300")  # Set the window size
root.configure(bg='#2c3e50')  # Set the background color

# Create a label for the title
title_label = tk.Label(root, text="Script Runner", font=("Arial", 20), bg='#2c3e50', fg='white')
title_label.pack(pady=20)

# Button styles
button_style = {
    'font': ('Arial', 14),
    'bg': '#3498db',
    'fg': 'white',
    'width': 25,
    'height': 2,
}

# Button to run run_V2.py
button1 = tk.Button(root, text="1、Run run_V2.py", command=lambda: run_script("run_V2.py"), **button_style)
button1.pack(pady=10)

# Button to run 內插V4_ui.py
button2 = tk.Button(root, text="2、Run 內插V4_ui.py", command=lambda: run_script("內插V4_ui.py"), **button_style)
button2.pack(pady=10)

# Button to run R50_ui_V2.py2
button3 = tk.Button(root, text="3、Run R50_ui_V2.py", command=lambda: run_script("R50_ui_V2.py"), **button_style)
button3.pack(pady=10)

# Run the GUI loop
root.mainloop()
