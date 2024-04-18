import tkinter as tk
from tkinter import messagebox, filedialog
import os, glob
import threading
from openpyxl import Workbook, load_workbook
import schedule
import time
from pystray import Icon as icon, MenuItem as item, Menu as menu
from PIL import Image, ImageTk

def check_and_prepare_files():
    global destination_path, source_path, keywords
    # Check destination file
    dest_file_path = os.path.join(destination_path.get(), "result.xlsx")
    if not os.path.isfile(dest_file_path):
        wb = Workbook()
        wb.save(filename=dest_file_path)

    # Check source path
    if not os.path.exists(source_path.get()):
        messagebox.showerror("Error", "Source folder path does not exist")
        return False

    return True

def process_files():
    if not check_and_prepare_files():
        return
    
    # Open destination Excel
    dest_file_path = os.path.join(destination_path.get(), "result.xlsx")
    wb_dest = load_workbook(dest_file_path)
    ws_dest = wb_dest.active

    # Ensure keywords are columns in the destination file
    keyword_list = keywords.get().split(',')
    for keyword in keyword_list:
        if keyword not in ws_dest[1]:
            ws_dest.append([keyword])

    # Process each source file
    for file in glob.glob(f"{source_path.get()}/*.xlsx"):
        if 'result.xlsx' in file:
            continue

        wb_src = load_workbook(file, data_only=True)
        ws_src = wb_src.active
        
        for row in ws_src.iter_rows():
            for cell in row:
                if cell.value in keyword_list:
                    # Assuming the keyword is found and the next cell is what we need
                    next_cell = row[cell.column + 1]
                    ws_dest[keyword_list.index(cell.value) + 1].append(next_cell.value)

    wb_dest.save(dest_file_path)
    messagebox.showinfo("Info", "Processing completed successfully!")

def on_run():
    # Run processing immediately then schedule it periodically
    process_files()
    time_frame_val = int(time_frame.get())
    schedule.every(time_frame_val).minutes.do(process_files)
    
    while True:
        schedule.run_pending()
        time.sleep(1)

def on_stop():
    global running
    running = False
    schedule.clear()
    icon.stop()

def create_tray_icon():
    icon_image = Image.open("icon.png")  # Ensure you have an icon.png in the script directory
    icon('Test', icon_image, menu=menu(item('Exit', on_stop))).run()

# Setup the GUI
root = tk.Tk()
root.title("Excel Processor")

destination_path = tk.StringVar()
source_path = tk.StringVar()
keywords = tk.StringVar()
time_frame = tk.StringVar()

tk.Label(root, text="Destination Folder Path:").pack()
tk.Entry(root, textvariable=destination_path).pack()
tk.Label(root, text="Source Folder Path:").pack()
tk.Entry(root, textvariable=source_path).pack()
tk.Label(root, text="Keywords (comma separated):").pack()
tk.Entry(root, textvariable=keywords).pack()
tk.Label(root, text="Time Frame (in minutes):").pack()
tk.Entry(root, textvariable=time_frame).pack()

tk.Button(root, text="Run", command=lambda: threading.Thread(target=on_run).start()).pack()
tk.Button(root, text="Stop", command=on_stop).pack()

root.protocol("WM_DELETE_WINDOW", create_tray_icon)
root.mainloop()
