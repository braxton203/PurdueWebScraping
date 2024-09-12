from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Setup Chrome WebDriver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

# Open the webpage
url = "https://www.coleparmer.com/chemical-resistance"
driver.get(url)

# Wait for the dropdown to load (adjust the selector based on the website structure)
try:
    # Wait until the dropdown is visible on the page
    dropdown = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "ChemicalListChemical"))  # Replace with the correct ID or class of the dropdown
    )

    # Find all the options inside the dropdown
    options = dropdown.find_elements(By.TAG_NAME, 'option')

    # Extract and print the text of each option
    print("Chemical List:")
    for option in options:
        print(option.text)

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    # Close the browser after scraping and success
    driver.quit()