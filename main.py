# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.
import json

from selenium.webdriver.support.wait import WebDriverWait
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
import pandas as pd
import time
from wakepy import keep

def load_configs(filepath="process_config.json"):
    """Loads CRM login credentials from JSON file."""
    with open(filepath, "r") as f:
        data = json.load(f)
    return (data["url"],data["username"], data["password"],data["input_file"],
            data["output_file"] , data["column_offset"] , data["delay_1"] , data["delay_2"] ,
            data["contact_headers"],data["srv_hist_headers"],data["drop_columns_name"])

def click_logout_button(driver):
    """
    Wait for and click the Logout button in Siebel Open UI.
    """
    try:
        wait = WebDriverWait(driver, 20)
        # Locate the button by its class and the 'title' attribute 'Logout'
        logout_button = wait.until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button.siebui-btn-logout[title='Logout']"))
        )
        # Scroll into view just in case
        driver.execute_script("arguments[0].scrollIntoView(true);", logout_button)
        logout_button.click()
        print("Logout button clicked successfully.")
    except Exception as e:
        print(f"Error clicking Logout button: {e}")


# Press the green button in the gutter to run the script.
if __name__ == '__main__':

    with (keep.running()):
        try:
            start_time = time.time()
            # Setup browser (Chrome here)
            driver = webdriver.Chrome()

            # Login if required
            url, username, password ,input_file , output_file,column_offset, delay_1, delay_2,\
            contact_headers,srv_hist_headers, cols_to_drop= load_configs()
            # print(url,username, password)
            print(url,input_file, output_file)
            driver.get(url)
            driver.find_element(By.ID, "s_swepi_1").send_keys(username)
            driver.find_element(By.ID, "s_swepi_2").send_keys(password)
            wait = WebDriverWait(driver, 10)
            login_button = wait.until(EC.element_to_be_clickable((By.ID, "s_swepi_22")))
            login_button.click()

            print("Login Successful")

            time.sleep(delay_2)  # wait for page to load

            # === STEP 2: Wait for Home Page ===
            wait = WebDriverWait(driver, 10)
            # wait.until(EC.presence_of_element_located((By.ID, "s_1_1_0_0_Ctrl")))  # Adjust based on what loads

            # === STEP 3: Navigate to Service History Tab ===
            # Option A: If it's a top menu link
            vehicle_history_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Vehicles")))
            vehicle_history_tab.click()

            # Read Excel file
            df = pd.read_excel(input_file)  # Reads entire sheet by default

            # Select only two columns (replace with your actual column names)
            # df = df[['Chassis No', 'Registration No']]
            df = df[['Registration_No']]

            contact_data_list = []
            srv_hist_data_list = []
            contact_df = pd.DataFrame()
            serv_hist_df = pd.DataFrame()
            # Loop over the DataFrame and print one column's value
            for index, row in df.iterrows():
                # chasis_no = row["Chassis No"]
                # print(chasis_no)
                registration_no = row["Registration_No"]
                print(registration_no)
                print("Before search click")
                time.sleep(delay_1)
                # Wait for the scrollable container to be present
                scroll_container = WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.siebui-applet-container"))
                )
                # Scroll to top
                driver.execute_script("arguments[0].scrollTop = 0;", scroll_container)
                search_button = driver.find_element(By.ID, "s_1_1_346_0_Ctrl")
                # Optionally add a small wait to ensure scrolling completed
                time.sleep(0.5)
                search_button.click()
                print("After search click")
                time.sleep(delay_1)
                # driver.find_element(By.NAME, "s_1_1_298_0").send_keys(chasis_no) # search by Chassis No.
                driver.find_element(By.NAME, "s_1_1_301_0").send_keys(registration_no) # search by Registration No.
                # 2. Wait for 'Go' button and click
                go_button = wait.until(EC.element_to_be_clickable((By.ID, "s_1_1_343_0_Ctrl")))
                go_button.click()

                # SERVICE HISTORY

                print("Before Service History click")

                time.sleep(delay_1)
                serv_hist_tab = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable(
                        (By.XPATH, "//a[@data-tabindex='tabScreen5' and contains(., 'Service History')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", serv_hist_tab)
                time.sleep(delay_1)
                try:
                    serv_hist_tab.click()
                except:
                    driver.execute_script("arguments[0].click();", serv_hist_tab)

                time.sleep(delay_1)

                print("After Service History click")
                # time.sleep(5)

                table_body = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#s_2_l tbody"))
                )

                # Find all data rows with class jqgrow inside the tbody
                srv_hist_data_rows = table_body.find_elements(By.CSS_SELECTOR, "tr.jqgrow")

                # time.sleep(5)

                # traverse through all service history rows
                srv_hist_rows_data = []
                # srv_hist_headers = ['SH #', 'Chassis No.', 'Registration No.', 'Account', 'SR #', 'Service Date/Time',
                #                     'Serviced At Dealer','Odometer Reading', 'Hours', 'SR Type', 'Summary', 'Survey Customer',
                #                     'Revisit','Service Request','Job Card Open Date', 'Customer Segment', 'Contact Full Name']

                # Locate the table by ID (adjust selector if needed)
                srv_hist_table = driver.find_element(By.ID, "s_2_l")

                # Locate all data rows inside tbody with class jqgrow and role row, or your actual selector
                srv_hist_data_rows = srv_hist_table.find_elements(By.CSS_SELECTOR, "tbody tr.jqgrow")
                srv_hist_row_dict = {}  # Initialize dict for 1st row

                if srv_hist_data_rows:
                    srv_hist_first_row = srv_hist_data_rows[0]

                    cells = srv_hist_first_row.find_elements(By.TAG_NAME, "td")
                    cell_texts = [cell.text.strip() for cell in cells]

                    # Detect if first cell is a checkbox or non-data cell
                    # first_cell_html = cells[0].get_attribute('innerHTML').lower()
                    # has_checkbox = 'type="checkbox"' in first_cell_html or 'checkbox' in first_cell_html
                    # offset = 2 if has_checkbox else 0
                    offset = 2

                    for i, header in enumerate(srv_hist_headers):
                        cell_idx = i + offset
                        if cell_idx < len(cell_texts):
                            print(f"  {header}: {cell_texts[cell_idx]}")
                            srv_hist_row_dict[header] = cell_texts[cell_idx]
                        else:
                            print(f"  {header}: <no data>")
                            srv_hist_row_dict[header] = None

                    # srv_hist_row_dict now contains the data for the first row with headers as keys
                else:
                    print("No data rows found in the table.")
                print("Adding srv_hist_row_dict to srv_hist_data_list")
                srv_hist_data_list.append(srv_hist_row_dict)


                print("Before contact click")

                # time.sleep(5)
                contacts_tab = WebDriverWait(driver, 20).until(
                    EC.element_to_be_clickable((By.XPATH, "//a[@data-tabindex='tabScreen0' and contains(., 'Contacts')]"))
                )
                driver.execute_script("arguments[0].scrollIntoView(true);", contacts_tab)
                # time.sleep(5)
                try:
                    contacts_tab.click()
                except:
                    driver.execute_script("arguments[0].click();", contacts_tab)

                print("After contact click")
                time.sleep(delay_2)

                # Wait for the grid table body to load and be present
                table_body = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "#s_2_l tbody"))
                )

                # Find all data rows with class jqgrow inside the tbody
                data_rows = table_body.find_elements(By.CSS_SELECTOR, "tr.jqgrow")

                # contact_headers = ['Customer Rel. No.', 'M/M', 'Last Name', 'First Name', 'Asset Relationship',
                #                    'Driving License',
                #                    'Expiry Date', 'RC Book Check', 'Company/Account', 'Site', 'Contact Address', 'City',
                #                    'State',
                #                    'Phone (R)', 'Phone(O)', 'Cell Phone No.', 'Email Address', 'Final Validation', 'Created By',
                #                    'Created Date', 'Customer Segment', 'Primary', 'Verification Status', 'Verified Date',
                #                    'Invoice Created Date']

                org_sale_dt_element = driver.find_element(By.XPATH,
                                                      '//input[contains(@class, "siebui-ctrl-date") and @aria-label="Original Sale Date"]')
                # Read the value attribute (this gives input’s content)
                org_sale_dt_value = org_sale_dt_element.get_attribute("value")
                print(f"Original Sale Date value: '{org_sale_dt_value}'")

                chassis_no_element = driver.find_element(By.XPATH,
                                                         '//input[contains(@class, "siebui-ctrl-input") and @aria-label="Chassis No"]')
                # Read the value attribute (this gives input’s content)
                chassis_no_value = chassis_no_element.get_attribute("value")
                print(f"Chassis No value: '{chassis_no_value}'")

                last_serv_dt_element = driver.find_element(By.XPATH,
                                                           '//input[contains(@class, "siebui-ctrl-date") and @aria-label="Last Service Date"]')
                # Read the value attribute (this gives input’s content)
                last_serv_dt_value = last_serv_dt_element.get_attribute("value")
                print(f"Last Service Date value: '{last_serv_dt_value}'")

                next_serv_dt_element = driver.find_element(By.XPATH, '//input[contains(@class, "siebui-ctrl-date") and @aria-label="Next Service Date"]')
                # Read the value attribute (this gives input’s content)
                next_serv_dt_value = next_serv_dt_element.get_attribute("value")
                print(f"Next Service Date value: '{next_serv_dt_value}'")

                war_exp_dt_element = driver.find_element(By.XPATH, '//input[contains(@class, "siebui-ctrl-date") and @aria-label="Warranty Expiry Date"]')
                # Read the value attribute (this gives input’s content)
                war_exp_dt_value = war_exp_dt_element.get_attribute("value")
                print(f"Warranty Expiry Date value: '{war_exp_dt_value}'")

                next_serv_type_element = driver.find_element(By.XPATH, '//input[contains(@class, "siebui-ctrl-input") and @aria-label="Next Service Type"]')
                # Read the value attribute (this gives input’s content)
                next_serv_type_value = next_serv_type_element.get_attribute("value")
                print(f"Next Service Type value: '{next_serv_type_value}'")

                prod_line_element = driver.find_element(By.XPATH, '//input[contains(@class, "siebui-ctrl-input") and @aria-label="Product Line"]')
                # Read the value attribute (this gives input’s content)
                prod_line_value = prod_line_element.get_attribute("value")
                print(f"Product Line value: '{prod_line_value}'")

                # traverse through all rows
                rows_data = []
                for row_index, row in enumerate(data_rows, start=1):
                    time.sleep(delay_1)
                    cells = row.find_elements(By.TAG_NAME, "td")
                    cell_texts = [cell.text.strip() for cell in cells]

                    # Detect if first cell is checkbox — check for input element of type checkbox
                    # first_cell_html = cells[0].get_attribute('innerHTML').lower()
                    # has_checkbox = 'type="checkbox"' in first_cell_html or 'checkbox' in first_cell_html
                    offset = 2
                    print(f"Row {row_index}:")
                    row_dict = {}
                    for i, header in enumerate(contact_headers):
                        cell_idx = i + offset
                        if cell_idx < len(cell_texts):
                            print(f"  {header}: {cell_texts[cell_idx]}")
                            row_dict[header] = cell_texts[cell_idx]
                        else:
                            print(f"  {header}: <no data>")
                            row_dict[header] = None  # or '', depending on preference
                    print("-" * 40)
                    rows_data.append(row_dict)

                # Now find the single row where 'Primary' == 'Y'
                primary_row = None
                for row in rows_data:
                    if row.get('Primary', '').upper() == 'Y':
                        primary_row = row
                        break  # remove break if multiple matches needed

                if primary_row:
                    print("Row where 'Primary' is 'Y':")
                    for k, v in primary_row.items():
                        print(f"{k}: {v}")
                else:
                    print("No row found where 'Primary' is 'Y'.")

                primary_dict = dict(primary_row) if primary_row is not None else {}
                primary_dict['Chassis No'] = chassis_no_value or 'Not Found'
                primary_dict['Original Sale Date'] = org_sale_dt_value  or 'Not Found'
                primary_dict['Warranty Expiry Date:'] = war_exp_dt_value  or 'Not Found'
                primary_dict['Last Service Date'] = last_serv_dt_value  or 'Not Found'
                primary_dict['Next Service Date'] = next_serv_dt_value  or 'Not Found'
                primary_dict['Next Service Type'] = next_serv_type_value  or 'Not Found'
                primary_dict['Product_Line'] = prod_line_value  or 'Not Found'

                print("Adding primary_dict to contact_data_list")
                contact_data_list.append(primary_dict)
                # time.sleep(5)
                # # Convert the single row dict to a DataFrame with one row
                # contact_df = pd.DataFrame([primary_row])

            # time.sleep(5)

            # Convert the single-row dictionary into a DataFrame
            contact_df = pd.DataFrame(contact_data_list)
            serv_hist_df = pd.DataFrame(srv_hist_data_list)
            final_df= pd.concat([contact_df, serv_hist_df], axis=1)
            # Columns you want to drop
            # cols_to_drop = ['Customer Rel. No.','M/M','Asset Relationship','Driving License','RC Book Check','Company/Account',
            #                 'Site','Final Validation','Created By','Created Date','Customer Segment','Primary','Verification Status',
            #                 'Verified Date','Invoice Created Date','SH #','Chassis No.','Account','SR #','Hours','Survey Customer',
            #                 'Revisit','Service Request','Job Card Open Date','Customer Segment']

            final_df.drop(columns=cols_to_drop, inplace=True)
            # Save to Excel file - specify your file name and path
            output_excel_file = output_file
            final_df.to_excel(output_excel_file, index=False)

            print(f"DataFrame saved to '{output_excel_file}'")
            print("dataframe created . wait is over. please log off")

            elapsed_seconds = time.time() - start_time
            print(f"Total time elapsed: {elapsed_seconds:.2f} seconds")
            wait = WebDriverWait(driver, 20)
            # Locate the button using CSS selector by class and aria-label attribute
            button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "li.siebui-toolbar-enable[aria-label='Settings']"))
            )
            # Scroll it into view then click
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            button.click()
            print("Settings toolbar button clicked successfully.")
            time.sleep(delay_2)
            wait = WebDriverWait(driver, 20)
            # Locate the button by its class and the 'title' attribute 'Logout'
            logout_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.siebui-btn-logout[title='Logout']"))
            )
            # Scroll into view just in case
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button)
            logout_button.click()
            print("Logout button clicked successfully.")
            time.sleep(delay_1)
            driver.quit()


        except Exception as e:
            print("Error Ocurred . log off:", e)
            time.sleep(delay_2)
            wait = WebDriverWait(driver, 20)
            # Locate the button using CSS selector by class and aria-label attribute
            button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "li.siebui-toolbar-enable[aria-label='Settings']"))
            )
            # Scroll it into view then click
            driver.execute_script("arguments[0].scrollIntoView(true);", button)
            button.click()
            print("Settings toolbar button clicked successfully.")
            time.sleep(delay_2)
            wait = WebDriverWait(driver, 20)
            # Locate the button by its class and the 'title' attribute 'Logout'
            logout_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.siebui-btn-logout[title='Logout']"))
            )
            # Scroll into view just in case
            driver.execute_script("arguments[0].scrollIntoView(true);", logout_button)
            logout_button.click()
            print("Logout button clicked successfully.")
            # Step 9: Close browser session
            driver.quit()
