import os
import math
import multiprocessing
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

daily_cols = ['AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ', 'BA', 'BB', 'BC', 'BD', 'BE', 'BF', 'BG', 'BH', 'BI', 'BJ', 'BK']
degree_cols = ['AQ','AS','AU', 'AW', 'AY', 'BA', 'BC', 'BE', 'BG', 'BI', 'BK']

def browse_file(entry_field):
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel and CSV files", "*.xlsx *.xls *.csv")]
    )
    if file_path:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, file_path)

def browse_folder(entry_field):
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_path)

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

def parse_row_to_dict(row):
    row = list(row)
    # row is a list like [Player, AT, Virgo, AU, 010-014, ..., Total, WinPercent]
    data = {}
    data['Player'] = row[0]
    for i in range(1, len(row)-4, 2):  # skip last 3 (Sum, Total, WinPercent)
        key = row[i]
        value = row[i+1]
        data[key] = value
    data["Count"] = row[-4]
    data['Total'] = row[-3]
    data['WinTotal'] = row[-2]
    data['WinPercent'] = row[-1]
    return data

def degree_match(daily_val, hist_range):
    try:
        low, high = map(int, hist_range.split('-'))
        return low <= int(daily_val) <= high
    except:
        return False
    
def multiprocess_rows(args):
    chunk, daily_df, raw_daily_df = args
    matches = []
    for idx, row in chunk:
        try:
            print(f"Processing {idx} row...")
            row_dict = parse_row_to_dict(row)
            hist_row = pd.Series(row_dict)

            for i, daily_row in daily_df.iterrows():
                is_match = True
                # row_sum = 0
                for col in hist_row.index:
                    if col in ['Count', 'Total', 'WinTotal', 'WinPercent']:
                        continue
                    daily_val = daily_row[col]
                    hist_val = hist_row[col]

                    print(col, daily_val, hist_val)

                    if not daily_val or not hist_val:
                        continue

                    if col == 'Player':
                        if daily_val != hist_val:
                            is_match = False
                            break
                    
                    else:
                        if int(daily_val) != int(hist_val):
                            is_match = False
                            break

                if is_match:
                    print(f"***** Found the matching result of {idx} row *****")
                    matched_row = raw_daily_df.iloc[i].tolist()
                    matched_row = matched_row + list(row)
                    matches.append(matched_row)
        except Exception as e:
            print(f"Error occured in the loop as {e}")
            continue
    return matches

def process_files():
    try:
        daily_file = daily_entry.get()
        historical_folder = hist_entry.get()

        if not os.path.exists(daily_file) or not os.path.exists(historical_folder):
            messagebox.showerror("Error", "One or both files not found.")
            return

        raw_daily_df = pd.read_csv(daily_file, header=None) if daily_file.endswith('csv') else pd.read_excel(daily_file, header=None)

        # filter Daily Data
        columns_to_extract = [0] + list(range(41, 63))
        daily_df = raw_daily_df[columns_to_extract]

        column_labels = ['Player'] + daily_cols
        daily_df.columns = column_labels

        all_rows = []
        for file_name in os.listdir(historical_folder):
            input_path = os.path.join(historical_folder, file_name)
            raw_hist_df = pd.read_csv(input_path, header=None) if input_path.endswith('csv') else pd.read_excel(input_path, header=None)

            all_rows.extend(list(raw_hist_df.iterrows()))

        chunk_size = math.ceil(len(all_rows) / os.cpu_count())
        chunks = [all_rows[i:i + chunk_size] for i in range(0, len(all_rows), chunk_size)]

        args = [(chunk, daily_df, raw_daily_df) for chunk in chunks]
        with multiprocessing.get_context("spawn").Pool() as pool:
            results = pool.map(multiprocess_rows, args)

        all_matches = [row for group in results for row in group if row]

        if all_matches:
            matched_df = pd.DataFrame(all_matches)
            matched_df = matched_df.dropna(how='all')
            output_path = os.path.splitext(daily_file)[0] + '_Matches.csv'
            matched_df.to_csv(output_path, index=False, header=None)
            # Extend your processing logic here
            messagebox.showinfo("Success", f"Processing finished...")
        else:
            messagebox.showerror("Error", f"NO Matches found...")
    except Exception as e:
        messagebox.showerror("Error", f"Error occured as {e}")

if __name__ == "__main__":
    # GUI setup
    root = tk.Tk()
    root.title("Matcher Processor with non-coloring")

    # Row 1 - Daily Input
    tk.Label(root, text="Daily File:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
    daily_entry = tk.Entry(root, width=50)
    daily_entry.grid(row=0, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_file(daily_entry)).grid(row=0, column=2, padx=5, pady=5)

    # Row 2 - Historical Input
    tk.Label(root, text="Historical % Input Folder:").grid(row=1, column=0, padx=5, pady=5, sticky="e")
    hist_entry = tk.Entry(root, width=50)
    hist_entry.grid(row=1, column=1, padx=5, pady=5)
    tk.Button(root, text="Browse", command=lambda: browse_folder(hist_entry)).grid(row=1, column=2, padx=5, pady=5)

    # Row 3 - Process Button
    tk.Button(root, text="Process", command=process_files, width=20).grid(row=2, column=1, pady=15)

    root.mainloop()