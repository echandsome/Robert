import pandas as pd
import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import requests
from bs4 import BeautifulSoup
import time
import random
import re

# City mapping dictionary
CITY_MAPPING = {
    "The Bronx (US), NY": "Bronx, NY",
    "Anaheim (US), CA": "Anaheim, CA",
    "Seattle (US), WA": "Seattle, WA",
    "Kansas City (US), MO": "Kansas City, MO",
    "Baltimore (US), MD": "Baltimore, MD",
    "St. Petersburg (US), FL": "Saint Petersburg, FL",
    "Pittsburgh (US), PA": "Pittsburgh, PA",
    "Boston (US), MA": "Boston, MA",
    "Chicago (US), IL": "Chicago, IL",
    "Cincinnati (US), OH": "Cincinnati, OH",
    "Toronto (CA), ON": "Toronto, ON",
    "West Sacramento (US), CA": "West Sacramento, CA",
    "Houston (US), TX": "Houston, TX",
    "Minneapolis (US), MN": "Minneapolis, MN",
    "Cleveland (US), OH": "Cleveland, OH",
    "Flushing (US), NY": "Flushing, NY",
    "Detroit (US), MI": "Detroit, MI",
    "Arlington (US), TX": "Arlington, TX",
    "Washington, D.C. (US), DC": "Washington, DC",
    "Philadelphia (US), PA": "Philadelphia, PA",
    "Los Angeles (US), CA": "Los Angeles, CA",
    "Atlanta (US), GA": "Atlanta, GA",
    "Miami (US), FL": "Miami, FL",
    "St. Louis (US), MO": "Saint Louis, MO",
    "Milwaukee (US), WI": "Milwaukee, WI",
    "Phoenix (US), AZ": "Phoenix, AZ",
    "San Diego (US), CA": "San Diego, CA",
    "San Francisco (US), CA": "San Francisco, CA",
    "Denver (US), CO": "Denver, CO"
}

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

def normalize_city_name(city_str):
    """Convert city name from format 'Denver (US), CO' to 'Denver, CO'"""
    if city_str in CITY_MAPPING:
        return CITY_MAPPING[city_str]
    
    # For cities not in mapping, remove "The" and "(US)" and clean up
    cleaned = city_str.strip()  # Remove leading/trailing whitespace
    if cleaned.startswith("The "):
        cleaned = cleaned[4:]  # Remove "The "
    if " (US)" in cleaned:
        cleaned = cleaned.replace(" (US)", "").strip()
    if "(US)" in cleaned:
        cleaned = cleaned.replace("(US)", "").strip()
    
    # Remove any remaining extra spaces
    cleaned = " ".join(cleaned.split())
    return cleaned

def extract_city_state(city_str):
    """Extract city and state from 'Denver, CO' format"""
    normalized = normalize_city_name(city_str)
    if ", " in normalized:
        city, state = normalized.split(", ", 1)
        return city, state
    return None, None

def scrape_moon_data_b(yyyymm_list):
    """Scrape moon data for B mode (Phoenix, AZ fixed)"""
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
            print(f"🔍 Scraping B mode: {url}")
            
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

def scrape_moon_data_p(city_date_list):
    """Scrape moon data for P mode (dynamic city-based URLs)"""
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

    # Group by city and YYYY-MM to minimize API calls (like B mode)
    city_month_groups = {}
    for city, date in city_date_list:
        city, state = extract_city_state(city)
        if city and state:
            yyyymm = date[:7]  # Extract YYYY-MM from date
            key = f"{city}_{state}_{yyyymm}"
            if key not in city_month_groups:
                city_month_groups[key] = set()
            city_month_groups[key].add(date)

    print(f"📊 Found {len(city_month_groups)} unique city-month combinations to scrape")

    for city_month_key, dates in city_month_groups.items():
        city, state, yyyymm = city_month_key.split("_", 2)
        url = f"https://www.almanac.com/astronomy/moon/calendar/{state}/{city}/{yyyymm}"
        
        try:
            print(f"🔍 Scraping P mode: {url}")
            
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
                        "city": f"{city}_{state}",
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

    if mode == "B":
        # Process B mode (columns 1, 9, 10)
        yyyymm_set = set()
        for i in range(len(df)):
            cell = df.iat[i, 1]  # Column 1
            if isinstance(cell, str):
                yyyymm = convert_to_yyyymm(cell)
                if yyyymm:
                    yyyymm_set.add(yyyymm)
            elif isinstance(cell, datetime.datetime):
                yyyymm = cell.strftime("%Y-%m")
                yyyymm_set.add(yyyymm)

        moon_map = scrape_moon_data_b(yyyymm_set)

        # Fill in moon data for B mode
        for i in range(len(df)):
            cell = df.iat[i, 1]  # Column 1
            date = None
            if isinstance(cell, str):
                date = normalize_date(cell)
            elif isinstance(cell, datetime.datetime):
                date = cell.strftime("%Y-%m-%d")
            
            if date:
                for moon in moon_map:
                    if date == moon["date"]:
                        df.iat[i, 9] = moon["phase"]   # Column 9
                        df.iat[i, 10] = moon["percent"] # Column 10
                        break

    elif mode == "P":
        # Process P mode (columns 15, 16, 17) with city-based URLs
        city_date_list = []
        for i in range(len(df)):
            date_cell = df.iat[i, 15]  # Column 15 (date)
            city_cell = df.iat[i, 6]   # Column G (city) - G column is index 6
            
            date = None
            if isinstance(date_cell, str):
                date = normalize_date(date_cell)
            elif isinstance(date_cell, datetime.datetime):
                date = date_cell.strftime("%Y-%m-%d")
            
            if date and city_cell:
                city_date_list.append((str(city_cell), date))

        moon_map = scrape_moon_data_p(city_date_list)

        # Fill in moon data for P mode
        for i in range(len(df)):
            date_cell = df.iat[i, 15]  # Column 15
            city_cell = df.iat[i, 6]   # Column G
            
            date = None
            if isinstance(date_cell, str):
                date = normalize_date(date_cell)
            elif isinstance(date_cell, datetime.datetime):
                date = date_cell.strftime("%Y-%m-%d")
            
            if date and city_cell:
                city, state = extract_city_state(str(city_cell))
                if city and state:
                    city_state = f"{city}_{state}"
                    for moon in moon_map:
                        if date == moon["date"] and moon.get("city") == city_state:
                            df.iat[i, 16] = moon["phase"]   # Column 16
                            df.iat[i, 17] = moon["percent"] # Column 17
                            break

    elif mode == "Both":
        # Process both B and P modes
        # B mode processing
        yyyymm_set = set()
        for i in range(len(df)):
            cell = df.iat[i, 1]  # Column 1
            if isinstance(cell, str):
                yyyymm = convert_to_yyyymm(cell)
                if yyyymm:
                    yyyymm_set.add(yyyymm)
            elif isinstance(cell, datetime.datetime):
                yyyymm = cell.strftime("%Y-%m")
                yyyymm_set.add(yyyymm)

        moon_map_b = scrape_moon_data_b(yyyymm_set)

        # Fill in moon data for B mode
        for i in range(len(df)):
            cell = df.iat[i, 1]  # Column 1
            date = None
            if isinstance(cell, str):
                date = normalize_date(cell)
            elif isinstance(cell, datetime.datetime):
                date = cell.strftime("%Y-%m-%d")
            
            if date:
                for moon in moon_map_b:
                    if date == moon["date"]:
                        df.iat[i, 9] = moon["phase"]   # Column 9
                        df.iat[i, 10] = moon["percent"] # Column 10
                        break

        # P mode processing
        city_date_list = []
        for i in range(len(df)):
            date_cell = df.iat[i, 15]  # Column 15 (date)
            city_cell = df.iat[i, 6]   # Column G (city)
            
            date = None
            if isinstance(date_cell, str):
                date = normalize_date(date_cell)
            elif isinstance(date_cell, datetime.datetime):
                date = date_cell.strftime("%Y-%m-%d")
            
            if date and city_cell:
                city_date_list.append((str(city_cell), date))

        moon_map_p = scrape_moon_data_p(city_date_list)

        # Fill in moon data for P mode
        for i in range(len(df)):
            date_cell = df.iat[i, 15]  # Column 15
            city_cell = df.iat[i, 6]   # Column G
            
            date = None
            if isinstance(date_cell, str):
                date = normalize_date(date_cell)
            elif isinstance(date_cell, datetime.datetime):
                date = date_cell.strftime("%Y-%m-%d")
            
            if date and city_cell:
                city, state = extract_city_state(str(city_cell))
                if city and state:
                    city_state = f"{city}_{state}"
                    for moon in moon_map_p:
                        if date == moon["date"] and moon.get("city") == city_state:
                            df.iat[i, 16] = moon["phase"]   # Column 16
                            df.iat[i, 17] = moon["percent"] # Column 17
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
