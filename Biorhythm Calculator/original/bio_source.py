import pandas as pd
import os
import time
import threading
from tkinter import Tk, Label, Button, StringVar, OptionMenu, filedialog, ttk, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# --- Global variables ---
df = None
file_path = ""
folder_path = ""
letter_to_col = {}
name_column = "" # Always column A
dob_column = ""
date_column = ""
biorhythm_columns = ["Emotional","Physical","Intellectual","Spiritual","Awareness","Intuitive","Aesthetic"]

# --- Processing function ---
def process_file(progress_var, progress_bar):
    global df, folder_path, name_column, dob_column, date_column
    driver, wait = start_browser()
    total = len(df)
    for idx, row in df.iterrows():
        name = row[name_column]
        dob = row[dob_column]
        set_date(driver, wait, dob)
        check_boxes(driver) # Click the 4 checkboxes solo si no est�n marcadas
        time.sleep(2)
        percentages = get_percentages(driver, wait)
        for col, value in percentages.items():
            if col in df.columns:
                df.at[idx, col] = value
                save_graph(driver, folder_path, name)
                progress_var.set(f"Processing {idx+1}/{total}: {name}")
                progress_bar['value'] = (idx+1)/total*100
                root.update_idletasks()
                driver.quit()
                # Save back to the same file
                if file_path.endswith(".csv"):
                    df.to_csv(file_path, index=False)
            else:
                df.to_excel(file_path, index=False)
                messagebox.showinfo("Done", f"Process completed. Images saved in folder: {folder_path}")
                progress_var.set("Finished")
                progress_bar['value'] = 100

# --- UI functions ---
def select_file():
    global df, file_path, folder_path, letter_to_col, name_column
    file_path = filedialog.askopenfilename(title="Select an Excel or CSV file", filetypes=[("Excel files","*.xlsx *.xls"),("CSV files","*.csv")])
    if not file_path:
        return
    file_name = os.path.splitext(os.path.basename(file_path))[0]
    folder_path = os.path.join(os.path.dirname(file_path), file_name)
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
        if file_path.endswith(".csv"):
            df = pd.read_csv(file_path)
    else:
        df = pd.read_excel(file_path)

    name_column = df.columns[0]

    num_cols = len(df.columns)
    letters = [chr(65 + i) for i in range(num_cols)]
    letter_to_col = {letters[i]: df.columns[i] for i in range(num_cols)}

    dob_dropdown['menu'].delete(0, 'end')
    date_dropdown['menu'].delete(0, 'end')
    for letter in letters:
        dob_dropdown['menu'].add_command(label=letter, command=lambda value=letter: dob_var.set(value))
        date_dropdown['menu'].add_command(label=letter, command=lambda value=letter: date_var.set(value))

        dob_var.set('B' if num_cols > 1 else 'A')
        date_var.set('F' if num_cols > 5 else 'A')
        progress_var.set(f"File loaded: {file_path}")

def start_processing():
    global dob_column, date_column
    if df is None:
        messagebox.showerror("Error", "Please select a file first.")
        return
    dob_column = letter_to_col[dob_var.get()]
    date_column = letter_to_col[date_var.get()]
    threading.Thread(target=process_file, args=(progress_var, progress_bar), daemon=True).start()

# --- Tkinter UI ---
root = Tk()
root.title("Biorhythm Processor")

Label(root, text="Select File:").grid(row=0, column=0, padx=10, pady=5)
Button(root, text="Browse", command=select_file).grid(row=0, column=1, padx=10, pady=5)

dob_var = StringVar()
Label(root, text="Date of Birth Column:").grid(row=1, column=0, padx=10, pady=5)
dob_dropdown = OptionMenu(root, dob_var, "")
dob_dropdown.grid(row=1, column=1, padx=10, pady=5)

date_var = StringVar()
Label(root, text="Date Column:").grid(row=2, column=0, padx=10, pady=5)
date_dropdown = OptionMenu(root, date_var, "")
date_dropdown.grid(row=2, column=1, padx=10, pady=5)

Button(root, text="Start Processing", command=start_processing).grid(row=3, column=0, columnspan=2, padx=10, pady=10)

progress_var = StringVar()
progress_var.set("Waiting...")
progress_label = Label(root, textvariable=progress_var)
progress_label.grid(row=4, column=0, columnspan=2)

progress_bar = ttk.Progressbar(root, length=300)
progress_bar.grid(row=5, column=0, columnspan=2, pady=10)

root.mainloop()