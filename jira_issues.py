from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
import os
from dotenv import load_dotenv, dotenv_values
from datetime import datetime, timezone
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill
from dateutil import parser
from urllib.parse import quote


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


def load_jira_filters():
    """
    Carga todas las variables definidas en el fichero .env que comienzan por "JIRA_FILTER_"
    y las devuelve en un diccionario.

    Returns:
        dict: Diccionario con las variables y sus valores.
    """
    # Cargar todas las variables del archivo .env
    env_vars = dotenv_values()
    # Filtrar las variables que empiezan con "JIRA_FILTER_"
    filters = {
        key: value for key, value in env_vars.items() if key.startswith("JIRA_FILTER_")
    }
    return filters


def extract_jira_issues(driver):
    """
    Extracts JIRA issue keys, links, summaries, statuses, priorities, customer object IDs
    and assignees from the issuetable and stores them in a pandas DataFrame

    Args:
        driver (WebDriver): Chrome WebDriver instance with JIRA page loaded

    Returns:
        DataFrame: Pandas DataFrame containing the issue keys, links, summaries, statuses,
        priorities, customer object IDs and assignees
    """
    print("Extrayendo issues de JIRA desde la tabla...")

    # Esperar a que la tabla de issues esté presente
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "issuetable"))
    )

    # Encontrar todas las filas en el tbody de la tabla
    rows = driver.find_elements(By.CSS_SELECTOR, "#issuetable tbody tr")

    issue_keys = []
    issue_types = []
    issue_links = []
    summaries = []
    statuses = []
    priorities = []
    customer_object_ids = []
    assignees = []
    created_dates = []
    classifications = []

    # Extraer el atributo data-issuekey de cada fila y el link de la issue
    for row in rows:
        # Get the issue key
        issue_key = row.get_attribute("data-issuekey")

        # Get the issue link
        issue_link = ""
        try:
            # Find the td with class "issuekey"
            issuekey_td = row.find_element(By.CSS_SELECTOR, "td.issuekey")

            # Find the a with class "issue-link" inside the td
            issue_link_element = issuekey_td.find_element(
                By.CSS_SELECTOR, "a.issue-link"
            )

            # Get the href attribute
            issue_link = issue_link_element.get_attribute("href")
        except Exception as e:
            print(f"Error getting link for issue {issue_key}: {e}")

        # Get the type of issue
        issue_type = ""
        try:
            # Find the td with class "issuetype"
            issue_type_td = row.find_element(By.CSS_SELECTOR, "td.issuetype")

            # Find the img element inside the issuetype td
            issue_type_img = issue_type_td.find_element(By.CSS_SELECTOR, "img")

            # Get the alt attribute
            issue_type = issue_type_img.get_attribute("alt").strip()
        except Exception as e:
            print(f"Error getting issuetype for issue {issue_key}: {e}")

        # Get the summary
        summary = ""
        try:
            # Find the td with class "summary"
            summary_td = row.find_element(By.CSS_SELECTOR, "td.summary")

            # Find the p element inside the summary td
            summary_element = summary_td.find_element(By.CSS_SELECTOR, "p")

            # Get the text content
            summary = summary_element.text.strip()
        except Exception as e:
            print(f"Error getting summary for issue {issue_key}: {e}")

        # Get the status
        status = ""
        try:
            # Find the td with class "status"
            status_td = row.find_element(By.CSS_SELECTOR, "td.status")

            # Find the span element inside the status td
            status_element = status_td.find_element(By.CSS_SELECTOR, "span")

            # Get the text content
            status = status_element.text.strip()
        except Exception as e:
            print(f"Error getting status for issue {issue_key}: {e}")

        # Get the priority
        priority = ""
        try:
            # Find the td with class "priority"
            priority_td = row.find_element(By.CSS_SELECTOR, "td.priority")

            # Find the img element inside the priority td
            priority_img = priority_td.find_element(By.CSS_SELECTOR, "img")

            # Get the alt attribute
            priority = priority_img.get_attribute("alt").strip()
        except Exception as e:
            print(f"Error getting priority for issue {issue_key}: {e}")

        # Get the Customer Object ID
        customer_object_id = ""
        try:
            # Find the td with class "customfield_14400"
            customfield_td = row.find_element(By.CSS_SELECTOR, "td.customfield_14400")

            # Get the text content directly from the td
            customer_object_id = customfield_td.text.strip()
        except Exception as e:
            print(f"Error getting Customer Object ID for issue {issue_key}: {e}")

        # Get the Assignee
        assignee = ""
        try:
            # Find the td with class "assignee"
            assignee_td = row.find_element(By.CSS_SELECTOR, "td.assignee")

            # Check if there's an <em> element (sin asignar)
            try:
                em_element = assignee_td.find_element(By.CSS_SELECTOR, "em")
                assignee = em_element.text.strip()
            except:
                # If no <em>, try to find <a> element (user assigned)
                try:
                    a_element = assignee_td.find_element(
                        By.CSS_SELECTOR, "a.user-hover"
                    )
                    assignee = a_element.text.strip()
                except:
                    # If neither is found, get direct text from td
                    assignee = assignee_td.text.strip()
        except Exception as e:
            print(f"Error getting assignee for issue {issue_key}: {e}")

        # Get the Created date
        created_date = ""
        try:
            # Find the td with class "created"
            created_td = row.find_element(By.CSS_SELECTOR, "td.created")

            # Find the time element inside the created td
            time_element = created_td.find_element(By.CSS_SELECTOR, "time")

            # Get the datetime attribute
            iso_date = time_element.get_attribute("datetime")

            # Convertir string ISO a objeto datetime de Python sin zona horaria
            if iso_date:
                try:
                    # Parsear la fecha ISO
                    dt_with_tz = parser.isoparse(iso_date)
                    # Convertir a UTC y eliminar la información de zona horaria
                    created_date = dt_with_tz.astimezone(timezone.utc).replace(
                        tzinfo=None
                    )
                except Exception as e:
                    print(f"Error converting date format for issue {issue_key}: {e}")
                    created_date = iso_date  # Fallback al formato original si hay error

        except Exception as e:
            print(f"Error getting Classification for issue {issue_key}: {e}")

        # Get the Classification
        classification = ""
        try:
            # Find the td with class "customfield_15400"
            customfield_td = row.find_element(By.CSS_SELECTOR, "td.customfield_15400")

            # Get the text content directly from the td
            classification = customfield_td.text.strip()
        except Exception as e:
            print(f"Error getting Classification for issue {issue_key}: {e}")

        if issue_key:
            issue_keys.append(issue_key)
            issue_types.append(issue_type)
            issue_links.append(issue_link)
            summaries.append(summary)
            statuses.append(status)
            priorities.append(priority)
            customer_object_ids.append(customer_object_id)
            assignees.append(assignee)
            created_dates.append(created_date)
            classifications.append(classification)

    # Crear un DataFrame con todos los datos recolectados
    df = pd.DataFrame(
        {
            "Issue Key": issue_keys,
            "Issue Type": issue_types,
            "Issue link": issue_links,
            "Summary": summaries,
            "Status": statuses,
            "Priority": priorities,
            "Customer Object ID": customer_object_ids,
            "Assignee": assignees,
            "Created": created_dates,
            "Classification": classifications,
        }
    )

    print(f"Se encontraron {len(issue_keys)} issues de JIRA")

    return df


def adjust_column_widths(sheet, max_width=80):
    """
    Adjusts the width of each column in the Excel sheet based on the maximum width of the data and header values.

    Parameters:
    -----------
    sheet : openpyxl.worksheet.worksheet.Worksheet
        The worksheet where column widths need to be adjusted.

    max_width : int, optional (default=80)
        The maximum allowed width for any column. If the calculated width exceeds this value,
        the column width will be set to this maximum value.

    Returns:
    --------
    None
    """
    for col in sheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name

        # Calculate the width required by the header (considering formatting)
        header_length = len(str(col[0].value))
        adjusted_header_length = (
            header_length * 1.5
        )  # Factor to account for bold and larger font size

        # Compare the header length with the lengths of the data values
        for cell in col:
            try:
                cell_length = len(str(cell.value))
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass

        # Use the greater of the header length or data length for column width
        max_length = max(max_length, adjusted_header_length)

        # Adjust the column width and apply the max_width limit
        adjusted_width = min(max_length + 2, max_width)
        sheet.column_dimensions[column].width = adjusted_width


def main():
    # Load environment variables from .env file
    load_dotenv()

    headless = os.getenv("HEADLESS_MODE", "False").lower() == "true"
    timeout = int(os.getenv("WAIT_TIME", "10"))

    # Element to wait for - can be defined in .env
    wait_element = os.getenv("WAIT_ELEMENT", ".issue-table-wrapper")

    # Get base URL from environment variables
    url_base = os.getenv("JIRA_URL_BASE")

    # Cargar todos los filtros JIRA disponibles
    jira_filters = load_jira_filters()
    print(f"Se han encontrado {len(jira_filters)} filtros JIRA para procesar")

    # Diccionario para almacenar los dataframes resultantes
    dataframes_dict = {}

    # Create the driver
    driver = create_chrome_driver(headless=headless)

    try:
        # Iterar sobre cada filtro JIRA
        for filter_name, filter_value in jira_filters.items():
            print(f"\nProcesando filtro: {filter_name}")

            # Extraer el nombre de la pestaña (eliminando el prefijo JIRA_FILTER_)
            sheet_name = filter_name.replace("JIRA_FILTER_", "")

            # Codificar el filtro JQL para uso seguro en URL
            filter_encoded = quote(filter_value) if filter_value else ""

            # Construir la URL completa con el filtro codificado
            complete_url = f"{url_base}?jql={filter_encoded}"

            # Navigate to page with explicit wait for the issue table
            navigate_to_url(
                driver, complete_url, wait_element=wait_element, timeout=timeout
            )

            # Extraer issues de JIRA en un DataFrame
            jira_issues_df = extract_jira_issues(driver)

            # Almacenar el DataFrame en el diccionario usando el nombre de la pestaña como clave
            dataframes_dict[sheet_name] = jira_issues_df

        # Generar un nombre de archivo con la fecha actual
        current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"jira_issues_{current_time}.xlsx"

        print(
            f"\nCreando archivo Excel con {len(dataframes_dict)} pestañas: {excel_filename}"
        )

        # Usar ExcelWriter para tener más control sobre el formato
        with pd.ExcelWriter(excel_filename, engine="openpyxl") as writer:
            # Crear una pestaña por cada DataFrame
            for sheet_name, df in dataframes_dict.items():
                print(f"Creando pestaña: {sheet_name} con {len(df)} issues")

                # Guardar el DataFrame en su pestaña correspondiente
                df.to_excel(writer, sheet_name=sheet_name, index=False)

                # Obtener el objeto worksheet actual
                worksheet = writer.sheets[sheet_name]

                # Aplicar formatos a la cabecera
                header_font = Font(bold=True)
                header_fill = PatternFill(
                    start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"
                )

                # Dar formato a la fila de encabezado
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill

                # Convertir enlaces de la columna "Issue link" a hipervínculos y dar formato a la columna de fechas
                link_col_index = None
                date_col_index = None

                # Encontrar las columnas relevantes
                for i, header in enumerate(worksheet[1], 1):
                    if header.value == "Issue link":
                        link_col_index = i
                    elif header.value == "Created":
                        date_col_index = i

                # Si encontramos la columna de enlaces, convertirlos a hipervínculos
                if link_col_index:
                    link_col_letter = get_column_letter(link_col_index)

                    # Para cada fila de datos (desde la 2 porque la 1 es el encabezado)
                    for row in range(2, worksheet.max_row + 1):
                        cell = worksheet[f"{link_col_letter}{row}"]
                        if cell.value and isinstance(cell.value, str):
                            # Establecer el hipervínculo
                            cell.hyperlink = cell.value
                            # Aplicar estilo de hipervínculo (azul subrayado)
                            cell.style = "Hyperlink"

                # Si encontramos la columna de fechas, asegurarse de que se muestran correctamente
                if date_col_index:
                    date_col_letter = get_column_letter(date_col_index)

                    # Para cada fila de datos (desde la 2 porque la 1 es el encabezado)
                    for row in range(2, worksheet.max_row + 1):
                        cell = worksheet[f"{date_col_letter}{row}"]
                        if cell.value:
                            # Aplicar formato de fecha y hora
                            cell.number_format = "yyyy-mm-dd hh:mm:ss"

                # Ajustar el ancho de las columnas basado en el contenido
                adjust_column_widths(worksheet)

                # Apply a filter to all columns
                worksheet.auto_filter.ref = worksheet.dimensions

            print(f"Archivo Excel creado y formateado exitosamente: {excel_filename}")

    finally:
        # Always close the driver to prevent resource leaks
        print("Closing browser...")
        driver.quit()


# def main():
#     # Load environment variables from .env file
#     load_dotenv()

#     headless = os.getenv("HEADLESS_MODE", "False").lower() == "true"
#     timeout = int(os.getenv("WAIT_TIME", "10"))

#     # Element to wait for - can be defined in .env
#     wait_element = os.getenv("WAIT_ELEMENT", ".issue-table-wrapper")

#     # Get URLs from environment variables
#     url_base = os.getenv("JIRA_URL_BASE")
#     # jira_filter = os.getenv("JIRA_SIN_ASIGNAR")
#     jira_filter = os.getenv("JIRA_FILTER_SIN_CERRAR")

#     # Codificar el filtro JQL para uso seguro en URL
#     jira_filter_encoded = quote(jira_filter) if jira_filter else ""

#     # Construir la URL completa con el filtro codificado
#     complete_url = f"{url_base}?jql={jira_filter_encoded}"

#     # Create the driver
#     driver = create_chrome_driver(headless=headless)

#     try:
#         # Navigate to first page with explicit wait for the issue table
#         navigate_to_url(
#             driver, complete_url, wait_element=wait_element, timeout=timeout
#         )

#         # Extraer issues de JIRA en un DataFrame
#         jira_issues_df = extract_jira_issues(driver)

#         # Generar un nombre de archivo con la fecha actual
#         current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
#         excel_filename = f"jira_issues_{current_time}.xlsx"

#         print(f"\nCreando archivo Excel: {excel_filename}")

#         # Usar ExcelWriter para tener más control sobre el formato
#         with pd.ExcelWriter(excel_filename, engine="openpyxl") as writer:
#             # Guardar el DataFrame en la hoja "JIRA Issues"
#             jira_issues_df.to_excel(writer, sheet_name="JIRA Issues", index=False)

#             # Obtener el objeto workbook y la hoja activa
#             workbook = writer.book
#             worksheet = writer.sheets["JIRA Issues"]

#             # Aplicar formatos a la cabecera
#             header_font = Font(bold=True)
#             header_fill = PatternFill(
#                 start_color="DDEBF7", end_color="DDEBF7", fill_type="solid"
#             )

#             # Dar formato a la fila de encabezado
#             for cell in worksheet[1]:
#                 cell.font = header_font
#                 cell.fill = header_fill

#             # Convertir enlaces de la columna "Issue link" a hipervínculos y dar formato a la columna de fechas
#             link_col_index = None
#             date_col_index = None

#             # Encontrar las columnas relevantes
#             for i, header in enumerate(worksheet[1], 1):
#                 if header.value == "Issue link":
#                     link_col_index = i
#                 elif header.value == "Created":
#                     date_col_index = i

#             # Si encontramos la columna de enlaces, convertirlos a hipervínculos
#             if link_col_index:
#                 link_col_letter = get_column_letter(link_col_index)

#                 # Para cada fila de datos (desde la 2 porque la 1 es el encabezado)
#                 for row in range(2, worksheet.max_row + 1):
#                     cell = worksheet[f"{link_col_letter}{row}"]
#                     if cell.value and isinstance(cell.value, str):
#                         # Establecer el hipervínculo
#                         cell.hyperlink = cell.value
#                         # Aplicar estilo de hipervínculo (azul subrayado)
#                         cell.style = "Hyperlink"

#             # Si encontramos la columna de fechas, asegurarse de que se muestran correctamente
#             if date_col_index:
#                 date_col_letter = get_column_letter(date_col_index)

#                 # Para cada fila de datos (desde la 2 porque la 1 es el encabezado)
#                 for row in range(2, worksheet.max_row + 1):
#                     cell = worksheet[f"{date_col_letter}{row}"]
#                     if cell.value:
#                         # Aplicar formato de fecha y hora
#                         cell.number_format = "yyyy-mm-dd hh:mm:ss"

#             # Ajustar el ancho de las columnas basado en el contenido
#             adjust_column_widths(worksheet)

#             # Apply a filter to all columns
#             worksheet.auto_filter.ref = worksheet.dimensions

#         print(f"Archivo Excel creado y formateado exitosamente: {excel_filename}")

#         # input("Presione Enter para salir...")

#     finally:
#         # Always close the driver to prevent resource leaks
#         print("Closing browser...")
#         driver.quit()


if __name__ == "__main__":
    main()
