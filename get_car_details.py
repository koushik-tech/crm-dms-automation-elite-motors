import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd

WIN_W, WIN_H = 800, 500
SELECTED_COLUMNS = ["Customer ID", "Order Date", "Amount"]

def load_background(path):
    """Load and set background image."""
    try:
        img = Image.open(path)
        img = img.resize((WIN_W, WIN_H), Image.ANTIALIAS)
        bg_photo = ImageTk.PhotoImage(img)
        canvas.bg_photo = bg_photo  # keep reference
        canvas.create_image(0, 0, anchor="nw", image=bg_photo)
    except Exception as e:
        messagebox.showwarning("Background Error", f"Unable to load background image:\n{e}")

def process_file():
    """Select Excel, filter columns, and save."""
    input_path = filedialog.askopenfilename(
        title="Select Input Excel File",
        filetypes=[("Excel Files", "*.xlsx *.xls")]
    )
    if not input_path:
        return

    try:
        df = pd.read_excel(input_path)
        df.columns = df.columns.str.strip()
    except Exception as e:
        messagebox.showerror("Read Error", f"Unable to read Excel file:\n{e}")
        return

    missing_cols = [c for c in SELECTED_COLUMNS if c not in df.columns]
    if missing_cols:
        messagebox.showerror("Missing Columns", f"Columns not found:\n{missing_cols}")
        return

    df_selected = df[SELECTED_COLUMNS]

    output_path = filedialog.asksaveasfilename(
        title="Save Output Excel File",
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")]
    )
    if not output_path:
        return

    try:
        df_selected.to_excel(output_path, index=False)
    except Exception as e:
        messagebox.showerror("Save Error", f"Unable to save file:\n{e}")
        return

    messagebox.showinfo("Success", f"Saved filtered file:\n{output_path}")

# --- GUI ---
root = tk.Tk()
root.title("Excel Column Selector")
root.resizable(False, False)
root.geometry(f"{WIN_W}x{WIN_H}")

# Canvas for background
canvas = tk.Canvas(root, width=WIN_W, height=WIN_H, highlightthickness=0)
canvas.pack(fill="both", expand=True)

# Load default background if present
script_dir = os.path.dirname(os.path.abspath(__file__))
bg_path = os.path.join(script_dir, "background.jpg")
if os.path.exists(bg_path):
    load_background(bg_path)

# Controls Frame (on top of background)
controls = tk.Frame(root, bg="#ffffff", padx=12, pady=12)
canvas.create_window(WIN_W//2, 30, anchor="n", window=controls)

title = tk.Label(controls, text="Select an Excel file to process",
                 font=("Segoe UI", 14, "bold"), bg="white")
title.pack(pady=(0, 8))

select_file_btn = tk.Button(controls, text="Select Excel & Process",
                            command=process_file, width=20,
                            bg="#2e86de", fg="white")
select_file_btn.pack(pady=6)

root.mainloop()
