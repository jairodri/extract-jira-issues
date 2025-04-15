from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import time
import os
from dotenv import load_dotenv


def create_chrome_driver(headless=False):
    """
    Creates and returns a configured Chrome WebDriver instance

    Args:
        headless (bool): Whether to run Chrome in headless mode

    Returns:
        WebDriver: Configured Chrome webdriver instance
    """
    print("Setting up Chrome options...")
    options = Options()
    options.add_argument("--start-maximized")

    if headless:
        options.add_argument("--headless")
        options.add_argument("--disable-gpu")
        options.add_argument("--window-size=1920,1080")

    print("Initializing Chrome browser...")
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=options
    )

    return driver


def navigate_to_url(driver, url, wait_element=None, timeout=10):
    """
    Navigates to a specified URL using the provided WebDriver and waits for an element to load

    Args:
        driver (WebDriver): Chrome WebDriver instance
        url (str): URL to navigate to
        wait_element (str, optional): CSS selector or ID of element to wait for
        timeout (int): Maximum time to wait for the element in seconds

    Returns:
        str: The title of the loaded page
    """
    print(f"Navigating to: {url}")
    driver.get(url)

    if wait_element:
        # Determine if the selector is an ID or a class
        if wait_element.startswith("#"):
            print(f"Waiting for element with ID: {wait_element[1:]}")
            by_method = By.ID
            selector = wait_element[1:]
        elif wait_element.startswith("."):
            print(f"Waiting for element with class: {wait_element[1:]}")
            by_method = By.CLASS_NAME
            selector = wait_element[1:]
        else:
            print(f"Waiting for element with selector: {wait_element}")
            by_method = By.CSS_SELECTOR
            selector = wait_element

        try:
            print(f"Waiting for element to load (timeout: {timeout}s)...")
            element = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((by_method, selector))
            )
            print(f"Element found: {element.tag_name}")
        except Exception as e:
            print(f"Timeout waiting for element: {wait_element}")
            print(f"Error: {e}")
    else:
        # Default wait for page to load if no specific element is specified
        print("No specific element to wait for. Waiting for page to load...")
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

    page_title = driver.title
    print(f"Current page title: {page_title}")

    return page_title


def extract_jira_issues(driver):
    """
    Extracts JIRA issue keys from the issuetable and stores them in a pandas DataFrame

    Args:
        driver (WebDriver): Chrome WebDriver instance with JIRA page loaded

    Returns:
        DataFrame: Pandas DataFrame containing the issue keys
    """
    print("Extrayendo issues de JIRA desde la tabla...")

    # Esperar a que la tabla de issues esté presente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "issuetable"))
    )

    # Encontrar todas las filas en el tbody de la tabla
    rows = driver.find_elements(By.CSS_SELECTOR, "#issuetable tbody tr")

    issue_keys = []

    # Extraer el atributo data-issuekey de cada fila
    for row in rows:
        issue_key = row.get_attribute("data-issuekey")
        if issue_key:
            issue_keys.append(issue_key)

    # Crear un DataFrame con los issue keys recolectados
    df = pd.DataFrame({"Issue Key": issue_keys})

    print(f"Se encontraron {len(issue_keys)} issues de JIRA")

    return df


def main():
    # Load environment variables from .env file
    load_dotenv()

    headless = os.getenv("HEADLESS_MODE", "False").lower() == "true"
    timeout = int(os.getenv("WAIT_TIME", "10"))

    # Get URLs from environment variables
    first_url = os.getenv("JIRA_SIN_ASIGNAR")

    # Element to wait for - can be defined in .env
    wait_element = os.getenv("WAIT_ELEMENT", ".issue-table-wrapper")

    # Create the driver
    driver = create_chrome_driver(headless=headless)

    try:
        # Navigate to first page with explicit wait for the issue table
        navigate_to_url(driver, first_url, wait_element=wait_element, timeout=timeout)

        # Extraer issues de JIRA en un DataFrame
        jira_issues_df = extract_jira_issues(driver)

        # Imprimir el DataFrame
        print("\nDataFrame de Issues de JIRA:")
        print(jira_issues_df)

        # Wait for user input before closing
        input("Press Enter to close the browser...")

    finally:
        # Always close the driver to prevent resource leaks
        print("Closing browser...")
        driver.quit()


if __name__ == "__main__":
    main()
