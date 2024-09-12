import requests
from bs4 import BeautifulSoup
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# URL of the webpage to scrape
url = "https://eldonjames.com/chemical-resistance/"

# Setup Chrome WebDriver using Selenium
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
driver.get(url)

# Wait for the table to load (adjust based on website structure)
try:
    # Wait for the table element to be present on the page (adjust selector accordingly)
    table = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "table"))
    )

    # Parse the page source with BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')

    # Find the table in the parsed HTML (adjust as per the website structure)
    table = soup.find('table')

    # Extract table rows
    rows = table.find_all('tr')

    # Prepare list to hold all the table data
    table_data = []

    # Loop through rows and extract data
    for row in rows:
        cols = row.find_all('td')  # Use 'th' if headers are needed
        cols = [col.text.strip() for col in cols]  # Get text content and clean spaces
        table_data.append(cols)

    # Convert the list to a Pandas DataFrame
    df = pd.DataFrame(table_data)

    # Save DataFrame to Excel
    df.to_excel('scraped_table_data.xlsx', index=False)

    print("Table scraped and saved to 'scraped_table_data.xlsx' successfully.")

except Exception as e:
    print(f"Error occurred: {e}")

finally:
    driver.quit()