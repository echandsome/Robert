import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys
import time

class CSVProcessor:
    def __init__(self):
        self.file_path = None
        self.driver = None
        self.setup_gui()
    
    def setup_gui(self):
        """Setup GUI"""
        self.root = tk.Tk()
        self.root.title("CSV Scraping Processor")
        self.root.geometry("400x200")
        
        # Main frame
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File selection button
        self.file_button = tk.Button(main_frame, text="Select CSV File", command=self.select_file, width=20, height=2)
        self.file_button.pack(pady=10)
        
        # Selected file display
        self.file_label = tk.Label(main_frame, text="No file selected", wraplength=350)
        self.file_label.pack(pady=5)
        
        # Start button (initially disabled)
        self.start_button = tk.Button(main_frame, text="Start Processing", command=self.process_csv, width=20, height=2, state=tk.DISABLED)
        self.start_button.pack(pady=10)
    
    def select_file(self):
        """Select CSV file"""
        file_path = filedialog.askopenfilename(
            title="Select CSV file",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        
        if file_path:
            self.file_path = file_path
            self.file_label.config(text=f"Selected: {os.path.basename(file_path)}")
            self.start_button.config(state=tk.NORMAL)
    
    def setup_browser(self):
        """Setup Chrome browser with options"""
        chrome_options = Options()
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--disable-web-security")
        chrome_options.add_argument("--allow-running-insecure-content")
        chrome_options.add_argument("--disable-features=VizDisplayCompositor")
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option('useAutomationExtension', False)
        
        self.driver = webdriver.Chrome(options=chrome_options)
        self.driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        # Navigate to FanDuel site
        self.driver.get("https://sportsbook.fanduel.com/search")
        
        # Wait for page to fully load
        wait = WebDriverWait(self.driver, 15)
        wait.until(lambda driver: driver.execute_script("return document.readyState") == "complete")
        time.sleep(2)  # Additional wait for dynamic content
        
        # Find and store the input element for later use
        self.find_search_input()
    
    def find_search_input(self):
        """Find the search input element and store it"""
        try:
            wait = WebDriverWait(self.driver, 10)
            
            # Try multiple selectors to find the input element
            selectors = [
                "input[placeholder='Search']",
                "input[type='text']",
                "input.search",
                "input[autocorrect='off']",
                "div[role='search'] input",
                "input"
            ]
            
            # Also try XPath selectors
            xpath_selectors = [
                "//input[@placeholder='Search']",
                "//input[@type='text']",
                "//input[contains(@class, 'search')]",
                "//input[contains(@placeholder, 'Search')]",
                "//input"
            ]
            
            input_element = None
            used_selector = None
            
            # Try CSS selectors first
            for selector in selectors:
                try:
                    input_element = wait.until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                    )
                    # Verify it's the search input by checking placeholder or other attributes
                    if input_element.get_attribute('placeholder') == 'Search' or 'search' in input_element.get_attribute('class', '').lower():
                        used_selector = selector
                        break
                except:
                    continue
            
            # If CSS selectors failed, try XPath selectors
            if not input_element:
                for xpath in xpath_selectors:
                    try:
                        input_element = wait.until(
                            EC.presence_of_element_located((By.XPATH, xpath))
                        )
                        # Verify it's the search input
                        if input_element.get_attribute('placeholder') == 'Search' or 'search' in input_element.get_attribute('class', '').lower():
                            used_selector = xpath
                            break
                    except:
                        continue
            
            if input_element:
                # Store the input element and selector for later use
                self.search_input = input_element
                self.search_selector = used_selector
                print(f"Successfully found search input using selector: {used_selector}")
            else:
                print("Could not find the search input element")
                # Print all input elements for debugging
                try:
                    all_inputs = self.driver.find_elements(By.TAG_NAME, "input")
                    print(f"Found {len(all_inputs)} input elements on the page:")
                    for i, inp in enumerate(all_inputs):
                        print(f"  Input {i+1}: placeholder='{inp.get_attribute('placeholder')}', type='{inp.get_attribute('type')}', class='{inp.get_attribute('class')}'")
                except Exception as debug_e:
                    print(f"Error getting debug info: {debug_e}")
                
        except Exception as e:
            print(f"Error finding search input element: {e}")
            # Try to print page source for debugging
            try:
                print("Page title:", self.driver.title)
                print("Current URL:", self.driver.current_url)
            except:
                pass
    
    def search_player(self, player_name):
        """Search for a player name in the search field"""
        try:
            if hasattr(self, 'search_input') and self.search_input:
                # Wait for element to be clickable
                wait = WebDriverWait(self.driver, 10)
                if hasattr(self, 'search_selector') and self.search_selector:
                    if self.search_selector in ["input[placeholder='Search']", "input[type='text']", "input.search", "input[autocorrect='off']", "div[role='search'] input", "input"]:
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, self.search_selector)))
                    else:
                        wait.until(EC.element_to_be_clickable((By.XPATH, self.search_selector)))
                
                # Format the search query: replace - with space and add "home run"
                formatted_query = player_name.replace('-', ' ') + ' home run'
                
                # Clear the input field completely and enter new search query
                self.search_input.click()  # Click to focus
                self.search_input.clear()  # Clear existing content
                time.sleep(0.5)  # Small delay to ensure clear is complete
                
                # Use Ctrl+A to select all text (in case clear didn't work completely)
                self.search_input.send_keys(Keys.CONTROL + 'a')
                time.sleep(0.5)
                
                # Enter the new search query
                self.search_input.send_keys(formatted_query)
                print(f"Successfully searched for: {formatted_query}")
                
                # Wait a bit for search results to load
                time.sleep(2)
                
                return True
            else:
                print("Search input element not found, trying to find it again...")
                self.find_search_input()
                if hasattr(self, 'search_input') and self.search_input:
                    return self.search_player(player_name)  # Retry
                else:
                    print(f"Failed to search for player: {player_name}")
                    return False
                    
        except Exception as e:
            print(f"Error searching for player {player_name}: {e}")
            return False
    
    def process_csv(self):
        """Read CSV file and iterate through each row's Player Name to add Home run line column and save."""
        if not self.file_path:
            messagebox.showinfo("Info", "Please select a file first")
            return
        
        try:
            # Setup browser and navigate to site (outside the for loop as requested)
            self.setup_browser()

            time.sleep(5)  # Reduced wait time since we're already waiting in setup_browser
            
            # Read CSV file
            df = pd.read_csv(self.file_path)
            
            # Iterate through each row to process Player Name
            for index, row in df.iterrows():
                player_name = row['Player Name']  # Get data from column A (Player Name)
                print(f"Processing player: {player_name}")
                
                # Search for the player
                search_success = self.search_player(player_name)
                
                if search_success:
                    # TODO: Add code here to extract home run line data from search results
                    # For now, setting a placeholder value
                    df.at[index, 'Home run line'] = 'SEARCHED'  # Changed from 'NM' to indicate search was performed
                    print(f"Successfully searched for {player_name}")
                else:
                    df.at[index, 'Home run line'] = 'SEARCH_FAILED'
                    print(f"Failed to search for {player_name}")
                
                # Add a small delay between searches to avoid overwhelming the site
                time.sleep(5)
            
            # Save with processed_ prefix
            file_dir = os.path.dirname(self.file_path)
            file_name = os.path.basename(self.file_path)
            processed_file_path = os.path.join(file_dir, f"Processed_{file_name}")
            
            df.to_csv(processed_file_path, index=False)
            
            messagebox.showinfo("Success", f"Processing completed! Saved as: Processed_{file_name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing file: {str(e)}")
        finally:
            # Close browser at the end of for loop (as requested)
            if self.driver:
                self.driver.quit()
                print("Browser closed")
    
    def run(self):
        """Run GUI"""
        self.root.mainloop()

if __name__ == "__main__":
    app = CSVProcessor()
    app.run()
