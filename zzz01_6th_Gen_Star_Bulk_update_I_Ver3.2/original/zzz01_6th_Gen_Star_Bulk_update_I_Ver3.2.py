import os
import itertools
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from multiprocessing import Pool, cpu_count

# Available columns for selection
COLUMN_MAPPING = {
    'AP': 41, 'AR': 43, 'AT': 45, 'AV': 47, 'AX': 49, 'AZ': 51,
    'BB': 53, 'BD': 55, 'BF': 57, 'BH': 59, 'BJ': 61
}

def get_next_column(col):
    first_letter = col[0]
    second_letter = col[1]

    new_second_letter = chr(ord(second_letter) + 1)

    if new_second_letter > 'Z':
        new_second_letter = 'A'
        new_first_letter = chr(ord(first_letter) + 1)
    else:
        new_first_letter = first_letter

    return new_first_letter + new_second_letter

# Function to process a single file
def process_file(input_df, is_degree, combination):
    selected_columns = ["Player"]
    col_indexes = [0]
    for item in combination:
        selected_columns.append(item[0])
        col_indexes.append(item[1])

        if is_degree:
            selected_columns.append(get_next_column(item[0]))
            col_indexes.append(item[1] + 1)
    
    # Extract group columns + result column (assume result in column 7)
    group_cols = input_df.iloc[:, col_indexes]
    group_cols.columns = selected_columns
    results = input_df.iloc[:, 7].astype(str).str.lower()
    df_grouped = pd.concat([group_cols, results.rename("Result")], axis=1)

    # Group by selected columns and count results
    grouped = df_grouped.groupby(selected_columns)
    output_rows = []

    for group_values, group_df in grouped:
        over = (group_df["Result"] == "over").sum()
        win = (group_df["Result"] == "win").sum()
        under = (group_df["Result"] == "under").sum()
        lose = (group_df["Result"] == "lose").sum()

        over += win
        under += lose
        total = over + under
        
        if total == 0:
            continue
        # row = {col: val for col, val in zip(selected_columns, group_values)}
        row = {}
        col_id = 0
        for col, val in zip(selected_columns, group_values):
            if col != "Player":
                row[f"Col_{col_id}"] = col
            row[f"Col_{col_id}_val"] = val
            col_id += 1
        # row["Over Count"] = over
        # row["Under Count"] = under
        row["Total"] = total
        row["WIN% OVER"] = round(over / total, 2)
        # row["WIN% UNDER"] = round(under / total, 2)
        output_rows.append(row)

    return output_rows

def browse_input_folder():
    folder_path = filedialog.askdirectory()
    input_entry.set(folder_path)

def process_file_wrapper(args):
    input_path, combinations, is_degree, set_size, output_dir = args
    results = []
    try:
        filename = os.path.basename(input_path)
        print(f"→ {filename} started")

        input_df = pd.read_excel(input_path, header=None)
        for comb_id, combination in enumerate(combinations, start=1):
            print(f"  Combo {comb_id}/{len(combinations)}: {combination}")
            results.extend(process_file(input_df, is_degree, combination))

        if results:
            output_name = f"{os.path.splitext(filename)[0]}_Size_{set_size}_Degree_{'YES' if is_degree else 'NO'}.csv"
            output_path = os.path.join(output_dir, output_name)
            pd.DataFrame(results).to_csv(output_path, header=None, index=None)
            print(f"✓ Saved to {output_path}")
        return f"{filename} completed"
    except Exception as e:
        return f"{input_path} failed: {e}"

def run():
    input_dir = input_entry.get()
    set_size = set_size_var.get()

    if not input_dir:
        messagebox.showerror("Error", "Please select an input folder.")
        return
    
    is_degree = include_degrees_var.get()

    try:
        output_dir = input_dir + "_output"
        os.makedirs(output_dir, exist_ok=True)
        combinations = list(itertools.combinations(COLUMN_MAPPING.items(), set_size))

        tasks = [
            (
                os.path.join(input_dir, filename),
                combinations,
                is_degree,
                set_size,
                output_dir
            )
            for filename in os.listdir(input_dir)
            if filename.endswith(".xlsx") and not filename.startswith("~$")
        ]

        with Pool(processes=cpu_count()) as pool:
            for result in pool.imap_unordered(process_file_wrapper, tasks):
                print(result)

        messagebox.showinfo("Success", f"Processed all files.\nSaved in: {output_dir}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    # GUI setup
    # Create the main window
    root = tk.Tk()
    root.title("Bulk File Processor")

    # First row: Input field and browse button for the folder
    input_entry = tk.StringVar()
    folder_frame = tk.Frame(root)
    folder_frame.pack(pady=10)

    folder_label = tk.Label(folder_frame, text="Select Folder with Excel Files:")
    folder_label.grid(row=0, column=0, padx=5)

    folder_entry = tk.Entry(folder_frame, textvariable=input_entry, width=40)
    folder_entry.grid(row=0, column=1, padx=5)

    browse_button = tk.Button(folder_frame, text="Browse", command=browse_input_folder)
    browse_button.grid(row=0, column=2, padx=5)

    # Second row: Label "Including Degrees" and checkbox
    include_degrees_var = tk.BooleanVar()
    include_degrees_checkbox = tk.Checkbutton(root, text="Including Degrees", variable=include_degrees_var)
    include_degrees_checkbox.pack(pady=10)

    # Third row: Set Size Radio Buttons (default is 3)
    set_size_var = tk.IntVar(value=3)  # Default set size is 3

    set_size_frame = tk.Frame(root)
    set_size_frame.pack(pady=10)

    set_size_label = tk.Label(set_size_frame, text="Select Set Size:")
    set_size_label.grid(row=0, column=0, padx=5)

    set_sizes = [3, 4, 5, 6, 7, 8]
    for idx, size in enumerate(set_sizes):
        rb = tk.Radiobutton(set_size_frame, text=str(size), variable=set_size_var, value=size)
        rb.grid(row=0, column=idx + 1, padx=5)

    # Fourth row: Process button
    process_button = tk.Button(root, text="Process", command=run)
    process_button.pack(pady=20)

    # Start the GUI loop
    root.mainloop()