import os
import string
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from openpyxl.utils import get_column_letter, column_index_from_string

col_num = {
    'U': 20, 'W': 22, 'Y': 24, 'AA': 26, 'AC': 28, 'AE': 30,
    'AG': 32, 'AI': 34, 'AK': 36, 'AM': 38, 'AO': 40, 'AQ': 42,
    'AS': 44, 'AU': 46, 'AW': 48, 'AY': 50, 'BA': 52, 'BC': 54,
    'BE': 56, 'BG': 58, 'BI': 60, 'BK': 62
}

# Define ranges
def map_to_range(val, group_list):
    try:
        val = int(val)
        for group in group_list:
            start, end = group.split('-')
            if int(start) <= val <= int(end):
                return group
    except:
        return val  # return original if not an int

# Process the Excel file
def process_file():
    file_path = file_entry.get().strip()
    txt_path = txt_entry.get().strip()
    if not os.path.exists(file_path) or not txt_path:
        messagebox.showerror("Error", "File not found!")
        return

    try:
        df = pd.read_excel(file_path, header=None)
        with open(txt_path, "r", encoding="utf-8") as wf:
            group_list = wf.readlines()
            group_list = [group.strip() for group in group_list]
        
        print(group_list)

        selected_cols = [col for col, var in checkbox_vars.items() if var.get()]

        for col in selected_cols:
            df[col_num[col]] = df[col_num[col]].apply(lambda x: map_to_range(x, group_list))

        # col_map = {'U': var_u, 'W': var_w, 'Y': var_y, 'AA': var_aa}
        # selected_cols = []
        # for col, var in col_map.items():
        #     if var.get():
        #         selected_cols.append(col)
        #         df[col_num[col]] = df[col_num[col]].apply(lambda x: map_to_range(x, group_list))

        output_path = os.path.splitext(file_path)[0] + f"_{''.join(selected_cols)}_Grouped.xlsx"
        df.to_excel(output_path, index=False, header=None)
        messagebox.showinfo("Success", f"File saved to:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read Excel file:\n{e}")
        return

# GUI setup
root = tk.Tk()
root.title("Degree Grouper")

file_frame = tk.Frame(root)
file_frame.pack(pady=10)

file_label = tk.Label(file_frame, text="Input Excel File")
file_label.pack(side=tk.LEFT, padx=5)
file_entry = tk.Entry(file_frame, width=50)
file_entry.pack(side=tk.LEFT)
browse_btn = tk.Button(file_frame, text="Browse", command=lambda: file_entry.insert(0, filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])))
browse_btn.pack(side=tk.LEFT, padx=5)

txt_frame = tk.Frame(root)
txt_frame.pack(pady=10)

txt_label = tk.Label(txt_frame, text="Groups TXT File")
txt_label.pack(side=tk.LEFT, padx=5)
txt_entry = tk.Entry(txt_frame, width=50)
txt_entry.pack(side=tk.LEFT)
txt_btn = tk.Button(txt_frame, text="Browse", command=lambda: txt_entry.insert(0, filedialog.askopenfilename(filetypes=[("TXT files", "*.txt")])))
txt_btn.pack(side=tk.LEFT, padx=5)

checkbox_frame = tk.Frame(root)
checkbox_frame.pack(pady=10)

checkbox_vars = {}

start_col = column_index_from_string('U')
end_col = column_index_from_string('BK')

col_letters = [get_column_letter(i) for i in range(start_col, end_col + 1, 2)]

for idx, col in enumerate(col_letters):
    var = tk.BooleanVar()
    checkbox_vars[col] = var
    tk.Checkbutton(checkbox_frame, text=f"Column {col}", variable=var).grid(row=idx // 6, column=idx % 6, padx=10, pady=2)

# var_u = tk.BooleanVar()
# tk.Checkbutton(checkbox_frame, text="Column U", variable=var_u).grid(row=0, column=0, padx=10)

# var_w = tk.BooleanVar()
# tk.Checkbutton(checkbox_frame, text="Column W", variable=var_w).grid(row=0, column=1, padx=10)

# var_y = tk.BooleanVar()
# tk.Checkbutton(checkbox_frame, text="Column Y", variable=var_y).grid(row=0, column=2, padx=10)

# var_aa = tk.BooleanVar()
# tk.Checkbutton(checkbox_frame, text="Column AA", variable=var_aa).grid(row=0, column=3, padx=10)

process_btn = tk.Button(root, text="Process", command=process_file)
process_btn.pack(pady=10)

root.mainloop()
