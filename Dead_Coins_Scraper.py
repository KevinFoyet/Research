from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
import pandas as pd
import time

BASE_URL = "https://www.coinopsy.com/dead-coins/"

def get_detail_link(driver, link_element):
    # Click on the coin's link to navigate to the detail page.
    link_element.click()

    # Try to find the 'Links' section and retrieve the link.
    try:
        link = driver.find_element(By.XPATH, "//b[text()='Links']/following-sibling::a").get_attribute('href')
    except:
        link = None  # Set to None if not found.

    # Navigate back to the main list.
    driver.back()

    return link

def get_data_from_section(driver):
    data = []

    # Wait for the table to load.
    wait = WebDriverWait(driver, 10)
    wait.until(EC.presence_of_element_located((By.XPATH, "//table")))

    rows = driver.find_elements(By.XPATH, "//table//tbody/tr")

    for row in rows:
        columns = row.find_elements(By.TAG_NAME, "td")
        name = columns[0].text
        summary = columns[1].text
        start_date = columns[2].text
        end_date = columns[3].text
        detail_link = None  # Initialize with None

        try:
            # Extract the 'href' attribute to get the link
            detail_link = columns[0].find_element(By.TAG_NAME, "a").get_attribute('href')
        except NoSuchElementException:
            # No link element found for this coin. Continue without the link.
            pass

        coin_data = {
            'Name': name,
            'Reason of Suspicion': summary,
            'Project Start Date': start_date,
            'Project End Date': end_date,
            'Detail Link': detail_link  # Store the link in your dictionary
        }

        data.append(coin_data)

    return data


# Rest of your code remains unchanged.



def main():
    all_data = []
    options = webdriver.ChromeOptions()
    options.add_argument('--headless')
    driver = webdriver.Chrome(options=options)
    driver.get(BASE_URL)

    # Assuming 248 sections as you mentioned
    for _ in range(248):
        print(f"Fetching data from current section...")
        data = get_data_from_section(driver)
        all_data.extend(data)

        # Try to click the 'Next' pagination button. If not found, break
        try:
            next_button = driver.find_element(By.XPATH, "//a[contains(text(), '>')]")
            next_button.click()
            time.sleep(2)  # Wait a bit for the next page to load
        except:
            break

    driver.quit()

    path = r"C:\Users\Kevin\OneDrive - Cal Poly Pomona\Documents\Dead_Coins_Scraper.xlsx"
    df = pd.read_excel(path, engine='openpyxl')
    new_df = pd.DataFrame(all_data)
    df = pd.concat([df, new_df], ignore_index=True)
    df.to_excel(path, index=False, engine='openpyxl')

    print("Data updated successfully!")

if __name__ == "__main__":
    main()
