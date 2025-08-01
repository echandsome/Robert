import pandas as pd
import tkinter as tk
from tkinter import filedialog
import os
import string

def col_letter_to_index(letter):
    letter = letter.upper()
    total = 0
    for i, char in enumerate(reversed(letter)):
        total += (string.ascii_uppercase.index(char) + 1) * (26 ** i)
    return total - 1

def split_file_by_column(file_path, column_letter):
    ext = os.path.splitext(file_path)[1].lower()
    if ext == '.csv':
        df = pd.read_csv(file_path, header=None)
    elif ext in ['.xls', '.xlsx']:
        df = pd.read_excel(file_path, header=None)
    else:
        raise ValueError("Unsupported file type.")

    col_index = col_letter_to_index(column_letter)
    if col_index >= df.shape[1]:
        raise ValueError("Column letter exceeds available columns in the file.")

    # Create outputs folder in same directory
    base_dir = os.path.dirname(file_path)
    output_dir = os.path.join(base_dir, "outputs")
    os.makedirs(output_dir, exist_ok=True)

    grouped = df.groupby(df[col_index])

    for group_name, group_df in grouped:
        safe_name = str(group_name).replace("/", "_").replace("\\", "_")
        output_filename = f"split_{column_letter.upper()}_{safe_name}{ext}"
        output_path = os.path.join(output_dir, output_filename)
        if ext == '.csv':
            group_df.to_csv(output_path, index=False, header=False)
        else:
            group_df.to_excel(output_path, index=False, header=False)

    result_label.config(text=f"Split complete!\nFiles saved in 'outputs' folder.", fg="green")

# --- GUI ---
def browse_file():
    path = filedialog.askopenfilename(title="Select File", filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
    if path:
        file_path_var.set(path)

def run_split():
    file_path = file_path_var.get()
    column_letter = column_var.get().strip()
    if not file_path or not column_letter:
        result_label.config(text="Please select a file and enter a column letter.", fg="red")
        return
    result_label.config(text="Processing...", fg="blue")
    root.update()
    try:
        split_file_by_column(file_path, column_letter)
    except Exception as e:
        result_label.config(text=f"Error: {str(e)}", fg="red")

root = tk.Tk()
root.title("Universal File Splitter")
root.geometry("500x250")

file_path_var = tk.StringVar()
column_var = tk.StringVar()

tk.Button(root, text="Select CSV/XLSX File", command=browse_file).pack(pady=(10, 0))
tk.Label(root, textvariable=file_path_var, wraplength=480).pack()

tk.Label(root, text="Enter Column Letter to Split By (e.g., A or P):").pack(pady=(10, 0))
tk.Entry(root, textvariable=column_var).pack()

tk.Button(root, text="Split File", command=run_split, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
