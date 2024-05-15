import os
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
from main import run_main_script

# Function to count HTML files in the selected source folder
def count_html_files(source_folder):
    try:
        return len([f for f in os.listdir(source_folder) if f.endswith(".htm") or f.endswith(".html")])
    except Exception as e:
        return 0

# Function to handle Run button click
def run_script():
    global source_folder, destination_file, stop_event, selected_time_frame

    if not source_folder or not os.path.exists(source_folder):
        messagebox.showerror("Error", "Please choose a valid source folder.")
        return

    if not destination_file or not os.path.exists(destination_file):
        messagebox.showerror("Error", "Please choose a valid destination file.")
        return

    if selected_time_frame is None:
        messagebox.showerror("Error", "Please select a time frame.")
        return

    # Clear the stop event before running the script
    stop_event.clear()

    # Run the main script in a separate thread to avoid freezing the UI
    threading.Thread(target=run_main_script, args=(source_folder, destination_file, stop_event, selected_time_frame)).start()

# Function to handle Stop button click
def stop_script():
    global stop_event
    stop_event.set()
    messagebox.showinfo("Info", "Script execution stopped.")

# Function to select source folder
def select_source_folder():
    global source_folder, html_file_count_label
    source_folder = filedialog.askdirectory()
    html_file_count = count_html_files(source_folder)
    html_file_count_label.config(text=f"Number of HTML files: {html_file_count}")

# Function to select destination file
def select_destination_file():
    global destination_file
    destination_file = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")])

# Function to handle time frame selection
def select_time_frame(time_frame):
    global selected_time_frame
    selected_time_frame = time_frame
    time_frame_label.config(text=f"Selected Time Frame: {time_frame}")

# Initialize the main window
root = tk.Tk()
root.title("Backtest Reports Processor")

source_folder = ""
destination_file = ""
stop_event = threading.Event()
selected_time_frame = None

# Source folder selector
source_folder_button = tk.Button(root, text="Select Source Folder", command=select_source_folder)
source_folder_button.pack(pady=10)

# Label to display number of HTML files
html_file_count_label = tk.Label(root, text="Number of HTML files: 0")
html_file_count_label.pack(pady=5)

# Destination file selector
destination_file_button = tk.Button(root, text="Select Destination File", command=select_destination_file)
destination_file_button.pack(pady=10)

# Time frame selection buttons
time_frame_label = tk.Label(root, text="Selected Time Frame: None")
time_frame_label.pack(pady=10)

time_frames = {
    "10 mins": 10,
    "30 mins": 30,
    "1 hour": 60,
    "4 hours": 240,
    "1 day": 1440
}

for label, time_frame in time_frames.items():
    time_frame_button = tk.Button(root, text=label, command=lambda tf=time_frame: select_time_frame(tf))
    time_frame_button.pack(pady=2)

# Run and Stop buttons
run_button = tk.Button(root, text="Run", command=run_script)
run_button.pack(pady=10)

stop_button = tk.Button(root, text="Stop", command=stop_script)
stop_button.pack(pady=10)

# Run the Tkinter event loop
root.mainloop()
