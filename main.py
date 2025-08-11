# This is a sample Python script.

# Press ⌃R to execute it or replace it with your code.
# Press Double ⇧ to search everywhere for classes, files, tool windows, actions, and settings.

from selenium.webdriver.support.wait import WebDriverWait


def print_hi(name):
   # Use a breakpoint in the code line below to debug your script.
   print(f'Hi, {name}')  # Press ⌘F8 to toggle the breakpoint.

# Press the green button in the gutter to run the script.
if __name__ == '__main__':
   print_hi('PyCharm')

   from selenium import webdriver
   from selenium.webdriver.common.by import By
   from selenium.webdriver.support.ui import WebDriverWait
   from selenium.webdriver.support import expected_conditions as EC
   from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException
   import pandas as pd
   import time

   # Load Excel
   # df = pd.read_excel("car_input.xlsx")  # Assume columns: CarNo, ChassisNo

   try:
       # Setup browser (Chrome here)
       driver = webdriver.Chrome()

       # Login if required
       driver.get("https://carsdms.inservices.tatamotors.com/siebel/app/workshop/enu?SWECmd=Start")
       driver.find_element(By.ID, "s_swepi_1").send_keys("IR_3084271")
       driver.find_element(By.ID, "s_swepi_2").send_keys("ITI@0225elite")
       wait = WebDriverWait(driver, 10)
       login_button = wait.until(EC.element_to_be_clickable((By.ID, "s_swepi_22")))
       login_button.click()
       # driver.find_element(By.ID, "s_swepi_22").click()

       output = []

       print("Login Successful")



       # for idx, row in df.iterrows():
       #     driver.get("https://your-crm-url.com/search")
       #     driver.find_element(By.ID, "car_no").send_keys(row["CarNo"])
       #     driver.find_element(By.ID, "chassis_no").send_keys(row["ChassisNo"])
       #     driver.find_element(By.ID, "submit_btn").click()
       #
       time.sleep(20)  # wait for page to load
       #
       #     # Extract results (adjust selector as per page)
       #     service_data = driver.find_element(By.ID, "service_details").text
       #     output.append({
       #         "CarNo": row["CarNo"],
       #         "ChassisNo": row["ChassisNo"],
       #         "ServiceData": service_data
       #     })

       # === STEP 2: Wait for Home Page ===
       wait = WebDriverWait(driver, 10)
       # wait.until(EC.presence_of_element_located((By.ID, "s_1_1_0_0_Ctrl")))  # Adjust based on what loads

       # === STEP 3: Navigate to Service History Tab ===
       # Option A: If it's a top menu link
       vehicle_history_tab = wait.until(EC.element_to_be_clickable((By.LINK_TEXT, "Vehicles")))
       vehicle_history_tab.click()
       print("Before vehicle click")
       button = WebDriverWait(driver, 25).until(
           EC.element_to_be_clickable((By.ID, "s_1_1_346_0_Ctrl"))
       )
       button.click()

       print("After vehicle click")

       time.sleep(20)
       driver.find_element(By.NAME, "s_1_1_301_0").send_keys("WB24AW9959")

       # 2. Wait for 'Go' button and click
       go_button = wait.until(EC.element_to_be_clickable((By.ID, "s_1_1_343_0_Ctrl")))
       go_button.click()

       print("Before contact click")

       time.sleep(30)
       contacts_tab = WebDriverWait(driver, 20).until(
           EC.element_to_be_clickable((By.XPATH, "//a[@data-tabindex='tabScreen0' and contains(., 'Contacts')]"))
       )
       driver.execute_script("arguments[0].scrollIntoView(true);", contacts_tab)
       time.sleep(10)
       try:
           contacts_tab.click()
       except:
           driver.execute_script("arguments[0].click();", contacts_tab)

       print("After contact click")
       time.sleep(20)

       # Wait for the grid table body to load and be present
       table_body = WebDriverWait(driver, 20).until(
           EC.presence_of_element_located((By.CSS_SELECTOR, "#s_2_l tbody"))
       )

       # Find all data rows with class jqgrow inside the tbody
       data_rows = table_body.find_elements(By.CSS_SELECTOR, "tr.jqgrow")


       # If you want the last row:
       if data_rows:
           last_row = data_rows[-1]
           # Optionally scroll into view
           driver.execute_script("arguments[0].scrollIntoView(true);", last_row)

           # For example, to click the last row somewhere (e.g., on first cell or the row itself)
           last_row.click()  # or last_row.find_element(...).click()

           # Get headers
           table = driver.find_element(By.ID, "s_2_l")
           time.sleep(20)
           # header_row = tablee.find_element(By.TAG_NAME, "thead").find_element(By.TAG_NAME, "tr")

           headers = ['Customer Rel. No.','M/M','Last Name','First Name','Asset Relationship','Driving License',
                      'Expiry Date','RC Book Check','Company/Account','Site','Contact Address','City','State',
                      'Phone (R)','Phone(O)','Cell Phone No.','Email Address','Final Validation','Created By',
                      'Created Date','Customer Segment','Primary','Verification Status','Verified Date','Invoice Created Date']
           # for th in header_row.find_elements(By.TAG_NAME, "th"):
           #     if th.is_displayed():
           #         header_text = th.text.strip()
           #         if header_text:
           #             headers.append(header_text)

           # Find all cell elements (<td>) inside the last row
           cells = last_row.find_elements(By.TAG_NAME, "td")
           # Extract the text from each cell
           cell_texts = [cell.text.strip() for cell in cells]

           print("Last row clicked or selected.")
           # for i in range(min(len(headers), len(cell_texts))):
           #     print(f"{headers[i]}: {cell_texts[i]}")
           # for idx, text in enumerate(cell_texts, start=0):
           #     print(f"Cell {idx}: {text}")

           for idx, cell in enumerate(cells):
               text1 = cell.text.strip()
               text2 = cell.get_attribute('textContent').strip()
               print(f"Cell {idx} - .text: '{text1}', textContent: '{text2}'")

           # first_name_cell = last_row.find_element(By.ID, "2_s_2_l_First_Name")
           # print("First Name:", first_name_cell.text.strip())
       else:
           print("No rows found in the grid.")

       # traverse through all rows
       for row_index, row in enumerate(data_rows, start=1):
           cells = row.find_elements(By.TAG_NAME, "td")
           cell_texts = [cell.text.strip() for cell in cells]

           # Detect if first cell is checkbox — check for input element of type checkbox
           first_cell_html = cells[0].get_attribute('innerHTML').lower()
           has_checkbox = 'type="checkbox"' in first_cell_html or 'checkbox' in first_cell_html
           offset = 1 if has_checkbox else 0
           offset = 2
           print(f"Row {row_index}:")
           for i, header in enumerate(headers):
               cell_idx = i + offset
               if cell_idx < len(cell_texts):
                   print(f"  {header}: {cell_texts[cell_idx]}")
               else:
                   print(f"  {header}: <no data>")
           print("-" * 40)

       # # Get all cells from the first row
       # first_row_cells = driver.find_elements(By.XPATH, "//td[starts-with(@id, '1_s_2_l_')]")
       #
       # # Extract text from each cell
       # first_row_data = [cell.text for cell in first_row_cells]
       #
       # print(first_row_data)

       # Wait for the element to be present and visible
       # input_element = WebDriverWait(driver, 10).until(
       #     EC.visibility_of_element_located((By.NAME, "s_1_1_301_0"))
       # )
       #
       # # Clear the field and enter the value
       # input_element.clear()
       # input_element.send_keys("MH12AB1234")
       # Option B: If it's a button or span
       # service_history_tab = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[text()='Service History']")))
       # service_history_tab.click()

       # === STEP 4: Wait for the tab view to load ===
       # wait.until(EC.presence_of_element_located((By.XPATH, "//input[@name='s_1_1_1_0']")))
       #
       # # === STEP 5: Set LOB and Value ===
       # driver.find_element(By.NAME, "s_1_1_1_0").send_keys("LOB")  # Find field
       # driver.find_element(By.NAME, "s_1_1_2_0").send_keys("Cars")  # Value field
       # driver.find_element(By.XPATH, "//button[@title='Go']").click()

       # Step 4: Click the "Go" button
       # go_button = driver.find_element(By.ID, "s_2_1_0_0_Ctrl")
       # go_button.click()

       # Wait for the button to be clickable
       # go_button = WebDriverWait(driver, 30).until(
       #     EC.element_to_be_clickable((By.ID, "s_1_1_346_0_Ctrl"))
       # )
       #
       # # Scroll into view (optional)
       # driver.execute_script("arguments[0].scrollIntoView(true);", go_button)
       #
       # # Click the button
       # go_button.click()

       # Optional: wait for results
       time.sleep(20)

       print("wait is over. please log off now ")
       # Get headers
       # headers = [header.text for header in driver.find_elements(By.XPATH, "//table//th")]
       # headers=[ 'SH #', 'Chassis No.', 'Registration No.', 'Account', 'SR #']
       # # Get data
       # rows = driver.find_elements(By.XPATH, "//table//tr")
       # data = []
       # for row in rows:
       #     cells = row.find_elements(By.TAG_NAME, "td")[:5]  # first 5 columns only
       #     data.append([cell.text.strip() for cell in cells])

       # all_data = []
       # while True:
       #     try:
       #         # Re-fetch rows inside loop to avoid stale references
       #         rows = driver.find_elements(By.XPATH, "//table[@class='siebui-list']//tr")
       #
       #         for row in rows:
       #             cells = row.find_elements(By.TAG_NAME, "td")
       #             if len(cells) < 5:
       #                 continue  # Skip rows that are too short or headers
       #             data = [cell.text.strip() for cell in cells[:5]]
       #             all_data.append(data)
       #
       #         # Try clicking the "Next" button
       #         if len(all_data)>1000:
       #             print("1000 records retrieved")
       #             break
       #         print("1......")
       #         # try:
       #         #     print("2......")
       #         #     next_button = driver.find_element(By.XPATH, "//span[@title='Next record set']")
       #         #     # Sometimes class may contain "disabled"
       #         #     if "ui-state-disabled" in next_button.get_attribute("class"):
       #         #         print("Reached last page of results.")
       #         #         break
       #         #     print("3......")
       #         #     driver.execute_script("arguments[0].click();", next_button)  # Use JS click
       #         #     print("4......")
       #         #     time.sleep(3)  # Wait for new page data
       #         # except NoSuchElementException:
       #         #     print("No 'Next' button found. Assuming end of pages.")
       #         #     break
       #
       #     except StaleElementReferenceException:
       #         print("Stale element detected. Retrying...")
       #         time.sleep(2)
       #         continue

       # all_data = []

       # print(f"Extracted {len(all_data)} rows from the service history grid.")

       # # Create DataFrame and save
       # df = pd.DataFrame(all_data, columns=headers)
       # df.to_excel("service_history.xlsx", index=False, engine='openpyxl')

       print("dataframe created . wait is over. please log off")

       time.sleep(20)
   except Exception as e :
       print("Error Ocurred . log off:",e)
       time.sleep(20)
           # Step 9: Close browser session
       driver.quit()

       # driver.quit()

       # Save to Excel
       # pd.DataFrame(output).to_excel("car_service_output.xlsx", index=False)

