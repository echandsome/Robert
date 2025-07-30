import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
import traceback
from bs4 import BeautifulSoup
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from datetime import datetime
import csv


STAT_OPTIONS = {
    "1": {"name": "Home Runs", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Home Runs')]"},  
    "2": {"name": "Hits", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Hits')]"},
    "3": {"name": "Runs", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Runs')]"},
    "4": {"name": "RBI", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'RBI')]"},
    "5": {"name": "Strikeouts", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Strikeouts')]"},
    "6": {"name": "Doubles", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Doubles')]"},
    "7": {"name": "Total Bases", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Total Bases')]"},
    "8": {"name": "Singles", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Singles')]"},
    "9": {"name": "Steals", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Steals')]"},
    "10": {"name": "Earned Runs", "xpath": "//label[contains(@class, 'checkbox__label') and contains(., 'Earned Runs')]"},
}


def scrape_player_stats(driver, player_url, num_games, home_or_away):
    """Scrape the last X games' stats for the player and corresponding HOME/AWAY games."""
    # Open player profile in a new tab
    driver.execute_script(f"window.open('{player_url}', '_blank');")
    driver.switch_to.window(driver.window_handles[-1])  # Switch to the new tab

    try:
        # Wait until the table is present on the page
        player_stats_table = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/div/div/div[1]/div/main/div/div/div[2]/section[4]/div/div[2]/table/tbody"))
        )
        
        # Wait until rows are present in the table
        rows = WebDriverWait(player_stats_table, 15).until(
            EC.presence_of_all_elements_located((By.TAG_NAME, 'tr'))
        )

        all_stats = []  # For the last 'num_games' games
        filtered_stats = []  # For the last 'num_games' of the specified type (home/away)

        # Iterate through rows to collect stats
        for row in rows:
            columns = row.find_elements(By.TAG_NAME, 'td')
            if len(columns) > 1:
                # Extract matchup column (2nd column)
                matchup = columns[1].text.strip()
                is_away = "@" in matchup

                # Determine if the game type matches the required home/away
                matches_filter = (home_or_away == "AWAY" and is_away) or (home_or_away == "HOME" and not is_away)

                # Extract stat value from the 7th column (index 6)
                stat_value = columns[5].text.strip()  # Column 7 (index 6)
                if "O" in stat_value or "U" in stat_value:
                    # Split the string by space and take the second part (the number)
                    stat_value = stat_value.split()[1]
                
                # Convert the stat value to an integer (if possible)
                try:
                    stat_value = int(stat_value)
                except ValueError:
                    stat_value = 0  # Default value for invalid or missing data
                
                # Add to the general stats list
                if len(all_stats) < num_games:
                    all_stats.append(stat_value)
                
                # Add to the filtered stats list if it matches the home/away type
                if matches_filter and len(filtered_stats) < num_games:
                    filtered_stats.append(stat_value)

            # Break early if we've gathered enough stats
            if len(all_stats) >= num_games and len(filtered_stats) >= num_games:
                break

        # Calculate averages
        avg_all = round(float(sum(all_stats) / len(all_stats)), 1) if all_stats else 0.0
        avg_filtered = round(float(sum(filtered_stats) / len(filtered_stats)), 1) if filtered_stats else 0.0

        return avg_all, avg_filtered  # Return both averages
    except Exception as e:
        print(f"Error while scraping player stats from player page: {e}")
        return None, None
    finally:
        driver.close()  # Close the current tab
        driver.switch_to.window(driver.window_handles[0])  # Switch back to the main tab



def scrape_page_data(driver, num_games, stat_category):
    """Scrape the current page's data and save it to CSV."""
    try:
        # Wait until the player prop cards container is present
        print("Waiting for player prop cards to load...")
        target_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[1]/div/main/div/div[2]/section/div[2]'))

        )


        current_date = datetime.now().strftime("%m-%d")

        # Extract all player prop cards
        player_prop_cards = WebDriverWait(target_element, 10).until(
            EC.presence_of_all_elements_located((By.CLASS_NAME, "player-prop-cards-container__card"))
        )
        print(f"Found {len(player_prop_cards)} player cards on this page.")

        stat_filename = f"{stat_category.lower().replace(' ', '_')}_player_props.csv"

        with open(stat_filename, 'a', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Player Name', 'Number', 'Odds', 'Projection', 'Avg', 'Home/Away Avg', 'Home/Away', 'Date', 'Stat Category','Team']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)

            # Write headers only once
            if csvfile.tell() == 0:
                writer.writeheader()

            # Process each player card
            for i, card in enumerate(player_prop_cards, start=1):
                try:
                    print(f"Processing card {i}...")

                    # Extract player URL
                    href = card.get_attribute("href")
                    if not href:
                        print(f"Card {i}: Missing href attribute!")
                        continue
                    player_name = href.split("/")[5] if href else "N/A"
                    print(f"Card {i}: Player Name - {player_name}")

                    # Extract number
                    try:
                        number_element = WebDriverWait(card, 30).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "div.flex.player-prop-card__prop-container span.typography:not(.player-prop-card__market)[style*='--9f36f340: var(--neutral-900, #16191D)']"))
                        )
                        number = number_element.text.strip() if number_element else "N/A"
                    except Exception as e:
                        number = "N/A"
                        print(f"Card {i}: Error extracting number element: {e}")
                        print(traceback.format_exc())

                    # Extract odds
                    try:
                        odds_element = WebDriverWait(card, 30).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "span.typography:not(.player-prop-card__team-pos)[style*='--9f36f340: var(--neutral-800, #525A67)']"))
                        )
                        odds = odds_element.text.strip() if odds_element else "N/A"
                    except Exception as e:
                        odds = "N/A"
                        print(f"Card {i}: Error extracting odds element: {e}")
                        print(traceback.format_exc())

                    # Extract projection
                    try:
                        projection_element = WebDriverWait(card, 30).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "span[style*='--9f36f340: var(--green-400, #1F845A)'], span[style*='--9f36f340: var(--red-400, #C9372C)']"))
                        )
                        projection = projection_element.text.strip() if projection_element else "N/A"
                    except Exception as e:
                        projection = "N/A"
                        print(f"Card {i}: Error extracting projection element: {e}")
                        print(traceback.format_exc())

                    # Extract team info
                    try:
                        team_info_element = WebDriverWait(card, 30).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, "span.typography.player-prop-card__team-pos"))
                        )
                        team_info = team_info_element.text.strip() if team_info_element else "N/A"
                    except Exception as e:
                        team_info = "N/A"
                        print(f"Card {i}: Error extracting team info element: {e}")
                        print(traceback.format_exc())

                    clean_team_info = team_info.split("-", 1)[-1].strip()


                    if "vs" in clean_team_info:
                        home_or_away = "HOME"
                        first_team = clean_team_info.split("vs")[0].strip()  # Extract first team before "vs"
                    elif "@" in clean_team_info:
                        home_or_away = "AWAY"
                        first_team = clean_team_info.split("@")[0].strip()  # Extract first team before "@"
                    else:
                        home_or_away = "N/A"
                        first_team = "N/A"

                    # Scrape the player's stats
                    print(f"Scraping stats for {player_name}...")
                    try:
                        avg_all, avg_filtered = scrape_player_stats(driver, href, num_games, home_or_away)
                    except Exception as e:
                        avg_all, avg_filtered = "N/A", "N/A"
                        print(f"Card {i}: Error scraping stats for {player_name}: {e}")
                        print(traceback.format_exc())

                    writer.writerow({
                        'Player Name': player_name,
                        'Number': number,
                        'Odds': odds,
                        'Projection': projection,
                        'Avg': avg_all,
                        'Home/Away Avg': avg_filtered,
                        'Home/Away': home_or_away,
                        'Date': f'"{current_date}"',  
                        'Stat Category': stat_category, 
                        'Team': first_team  
                    })

                    print(f"Successfully processed card {i} for player: {player_name}")
                    print("-" * 40)

                except Exception as e:
                    print(f"Error processing card {i}: {e}. Skipping to the next card.")
                    print(traceback.format_exc())  # This will give you the full stack trace for debugging
                    print("-" * 40)

    except Exception as e:
        print(f"Error in scraping page data: {e}")
        print(traceback.format_exc())  # This will give you the stack trace for the outer try block


def scrape_selected_stats():
    """Prompt user to select stats and start the scraping process."""
    # Prompt user to select one or more stats
    print("Select stats to scrape (enter numbers separated by commas):")
    for key, value in STAT_OPTIONS.items():
        print(f"{key}: {value['name']}")
    
    user_input = input("Your choice: ").split(",")
    selected_options = [option.strip() for option in user_input if option.strip() in STAT_OPTIONS]

    if not selected_options:
        print("No valid options selected. Exiting.")
        return

    print(f"Selected options: {[STAT_OPTIONS[opt]['name'] for opt in selected_options]}")

    # Prompt user for the number of games to scrape
    try:
        num_games = int(input("Enter the number of recent games to scrape (e.g., 5): "))
        if num_games <= 0:
            print("Please enter a positive number of games.")
            return
    except ValueError:
        print("Invalid input. Please enter a valid integer for the number of games.")
        return

    # Set up the WebDriver
    options = webdriver.ChromeOptions()
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--start-maximized")
   
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    
    try:
        # Navigate to the website
        url = "https://www.bettingpros.com/mlb/props/"
        driver.get(url)

        # Wait for the page to load completely
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

        # Handle cookie consent banner first
        try:
            # Try to find and close cookie consent banner
            cookie_banner = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((By.ID, "onetrust-policy-text"))
            )
            # Look for accept button in the banner
            accept_button = driver.find_element(By.XPATH, "//button[contains(text(), 'Accept') or contains(text(), 'Accept All') or contains(text(), 'OK')]")
            accept_button.click()
            print("Closed cookie consent banner")
            time.sleep(2)  # Wait for banner to disappear
        except Exception as e:
            print("No cookie banner found or already handled")

        for option in selected_options:
            stat_name = STAT_OPTIONS[option]["name"]
            stat_xpath = STAT_OPTIONS[option]["xpath"]

            print(f"Starting scrape for {stat_name}...")

            # Navigate to the website
            url = "https://www.bettingpros.com/mlb/props/"
            driver.get(url)

            # Wait for the page to load completely
            WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.TAG_NAME, 'body')))

            # Click the button corresponding to the current stat
            try:
                # Wait for the element to be present and clickable
                button = WebDriverWait(driver, 10).until(
                    EC.element_to_be_clickable((By.XPATH, stat_xpath))
                )
                
                button.click()
                
                print(f"Clicked on {stat_name}")
            except Exception as e:
                print(f"Error clicking on {stat_name}: {e}")
                continue  # Skip to the next stat category if there's an issue

            time.sleep(3)  # Allow some time for the page to update

            # Scrape data for the current stat
            scrape_page_data(driver, num_games=num_games, stat_category=stat_name)
            scraped_stats = set()

        

            # Attempt to click "Next Page" button repeatedly
            next_button_xpath = '/html/body/div[1]/div/div/div[1]/div/main/div/div[2]/section/div[3]/button[2]'
            while True:
                try:
                    # Wait until the next button is clickable
                    next_button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, next_button_xpath)))
                    # Scroll to the button to ensure it's visible
                    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", next_button)
                    time.sleep(1)
                    
                    if "disabled" not in next_button.get_attribute("outerHTML"):
                        next_button.click()
                        print("Clicked 'Next Page' button.")
                        time.sleep(3)  # Wait for the page to load
                        scrape_page_data(driver, num_games=num_games, stat_category=stat_name)
                    else:
                        print("Next page button is disabled, no more pages to scrape.")
                        break
                except Exception as e:
                    print(f"Error while clicking 'Next Page' button: {e}")
                    break

    except Exception as e:
            print(f"An error occurred: {e}")
    finally:
        # Quit the driver
        driver.quit()


if __name__ == "__main__":
    scrape_selected_stats()

