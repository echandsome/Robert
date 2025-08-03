import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup
import time
import random

def convert_to_yyyymm(date_str):
    try:
        date_str = date_str.strip("()").strip()
        try:
            dt = datetime.datetime.strptime(date_str, "%m-%d-%Y")
        except ValueError:
            dt = datetime.datetime.strptime(date_str, "%m/%d/%Y")
        return dt.strftime("%Y-%m")
    except Exception as e:
        print(f"Date parse error: {date_str} -> {e}")
        return None

def scrape_moon_data(yyyymm_list):
    result = []
    
    # Create a session for better request handling
    session = requests.Session()
    session.headers.update({
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive",
        "Upgrade-Insecure-Requests": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "none",
        "Cache-Control": "max-age=0"
    })

    for yyyymm in yyyymm_list:
        url = f"https://www.almanac.com/astronomy/moon/calendar/AZ/Phoenix/{yyyymm}"
        try:
            print(f"🔍 Scraping: {url}")
            
            # Add random delay between requests (1-3 seconds)
            time.sleep(random.uniform(1, 3))
            
            response = session.get(url, timeout=30)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")
            tds = soup.find_all("td", class_="calday")
            year, month = map(int, yyyymm.split("-"))

            for td in tds:
                try:
                    day_tag = td.find("p", class_="daynumber")
                    if not day_tag or not day_tag.text.strip().isdigit():
                        continue

                    day = int(day_tag.text.strip())
                    phase = None
                    percent = None

                    minor = td.find("p", class_="phasename_minor")
                    if minor:
                        parts = list(minor.stripped_strings)
                        if len(parts) == 2:
                            phase = parts[0].lower()
                            percent = parts[1]
                    else:
                        major = td.find("p", class_="phasename")
                        if major:
                            parts = list(major.stripped_strings)
                            if len(parts) >= 1:
                                phase = parts[0].lower()
                                if "moon" in phase and "full" in phase:
                                    phase = 'full moon'
                                    percent = "100%"
                                elif "moon" in phase and "new" in phase:
                                    phase = 'new moon'
                                    percent = "0%"
                                elif "quarter" in phase:
                                    percent = "50%"
                                else:
                                    percent = "50%"

                    result.append({
                        "date": f"{year:04d}-{month:02d}-{day:02d}",
                        "phase": phase,
                        "percent": percent
                    })

                except Exception as td_error:
                    print(f"Error parsing td: {td_error}")

        except requests.exceptions.HTTPError as e:
            if e.response.status_code == 403:
                print(f"❌ Access forbidden (403) for {url}")
                print("   This usually means the website is blocking automated requests.")
                print("   Try running the script again later or use a different approach.")
            else:
                print(f"Failed to fetch {url} → {e}")
        except Exception as e:
            print(f"Failed to fetch {url} → {e}")

    return result

def normalize_date(date_str):
    try:
        date_str = date_str.strip("()")
        try:
            dt = datetime.datetime.strptime(date_str, "%m-%d-%Y")
        except ValueError:
            dt = datetime.datetime.strptime(date_str, "%m/%d/%Y")
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return None
    
def process_file(filepath, mode):
    df = pd.read_excel(filepath, header=None)

    # Ensure enough columns exist
    max_required_col = 17
    while df.shape[1] <= max_required_col:
        df[df.shape[1]] = ""

    yyyymm_set = set()
    date_columns = []

    if mode == "B":
        date_columns = [(1, 9, 10)]  # (input_col, phase_col, percent_col)
    elif mode == "P":
        date_columns = [(15, 16, 17)]
    elif mode == "Both":
        date_columns = [(1, 9, 10), (15, 16, 17)]

    # Collect all YYYY-MM for scraping
    for input_col, _, _ in date_columns:
        for i in range(len(df)):
            cell = df.iat[i, input_col]
            if isinstance(cell, str):
                yyyymm = convert_to_yyyymm(cell)
                if yyyymm:
                    yyyymm_set.add(yyyymm)
            elif isinstance(cell, datetime.datetime):
                yyyymm = cell.strftime("%Y-%m")
                yyyymm_set.add(yyyymm)

    moon_map = scrape_moon_data(yyyymm_set)

    # Fill in moon data
    for input_col, phase_col, percent_col in date_columns:
        for i in range(len(df)):
            cell = df.iat[i, input_col]
            date = None
            if isinstance(cell, str):
                date = normalize_date(cell)
            elif isinstance(cell, datetime.datetime):
                date = cell.strftime("%Y-%m-%d")
            
            if date:
                for moon in moon_map:
                    if date == moon["date"]:
                        df.iat[i, phase_col] = moon["phase"]
                        df.iat[i, percent_col] = moon["percent"]
                        break

    dir_path = os.path.dirname(filepath)
    base_name = os.path.splitext(os.path.basename(filepath))[0]
    output_path = os.path.join(dir_path, f"{base_name}_moon.xlsx")
    df.to_excel(output_path, index=False, header=False)
    return output_path

# GUI functions
def browse_file():
    path = filedialog.askopenfilename(title="Select Excel file", filetypes=[("Excel files", "*.xlsx")])
    if path:
        file_path_var.set(path)

def run_process():
    file_path = file_path_var.get()
    mode = radio_var.get()
    if not file_path:
        messagebox.showerror("Error", "Please select a file.")
        return

    result_label.config(text="Processing...", fg="blue")
    root.update()

    try:
        output_path = process_file(file_path, mode)
        result_label.config(text=f"Complete!\nResult saved:\n{output_path}", fg="green")
    except Exception as e:
        print(str(e))
        result_label.config(text=f"Error occurred: {str(e)}", fg="red")

# GUI Setup
root = tk.Tk()
root.title("Moon Scrapping")
root.geometry("500x300")

file_path_var = tk.StringVar()
radio_var = tk.StringVar(value="Both")

tk.Button(root, text="Select Excel File (.xlsx)", command=browse_file).pack(pady=(10, 0))
tk.Label(root, textvariable=file_path_var, wraplength=480).pack()

tk.Label(root, text="Select mode:").pack(pady=(10, 0))
tk.Radiobutton(root, text="B", variable=radio_var, value="B").pack()
tk.Radiobutton(root, text="P", variable=radio_var, value="P").pack()
tk.Radiobutton(root, text="Both", variable=radio_var, value="Both").pack()

tk.Button(root, text="Execute", command=run_process, bg="green", fg="white").pack(pady=20)

result_label = tk.Label(root, text="", font=("Arial", 10))
result_label.pack()

root.mainloop()
