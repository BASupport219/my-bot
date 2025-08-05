from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import StaleElementReferenceException, NoSuchElementException, TimeoutException, UnexpectedAlertPresentException, InvalidSessionIdException
import time
import logging
import os
import pandas as pd

# Set up logging for essential messages only
logging.basicConfig(level=logging.INFO, format='✅ %(message)s')
logger = logging.getLogger()

# Set browser visibility: Change to "Yes" to show browser or "No" for headless mode
visibility = "No"  # Manually change this to "No" to run headless (no browser visible)

# File to upload and update (not used since script stops after login)
excel_file_path = r"C:\Users\USER\Desktop\JILI\Book4.xlsx"

# Constants
HOUR_IN_SECONDS = 3600  # 1 hour
HOURLY_RESTART_WAIT = 120  # 2 minutes wait for hourly restart
ERROR_RESTART_WAIT = 5  # 5 seconds wait for error restart
TABLE_CHECK_WAIT = 10  # Seconds to wait for table check or retry withdrawal page
EXECUTOR_CHECK_WAIT = 3  # Seconds to wait for Executor row check
CONFIRM_VERIFY_WAIT = 3  # Seconds to wait for Confirm Verify button

def setup_browser():
    """Set up and return Chrome driver with configured options."""
    options = webdriver.ChromeOptions()
    if visibility.lower() == "no":
        options.add_argument("--headless=new")  # Use new headless mode for better stability
        options.add_argument("--disable-gpu")  # Disable GPU for headless compatibility
        options.add_argument("--no-sandbox")  # Bypass OS security model
        options.add_argument("--disable-dev-shm-usage")  # Overcome limited resource problems
        options.add_argument("--window-size=1920,1080")  # Set window size to avoid rendering issues
        logger.info("Running in headless mode (browser not visible)")
    else:
        logger.info("Running with browser visible")
    
    # Suppress console errors and logging
    options.add_argument("--log-level=3")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    # Additional options for stability
    options.add_argument("--disable-extensions")  # Disable extensions to reduce crashes
    options.add_argument("--disable-popup-blocking")  # Allow pop-ups to prevent session issues
    options.add_argument("--start-maximized")  # Ensure consistent rendering
    options.add_argument("--disable-blink-features=AutomationControlled")  # Avoid detection as bot

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(60)  # Set page load timeout to prevent hanging
    wait = WebDriverWait(driver, 60)  # Main wait for critical actions
    return driver, wait

def cleanup_browser(driver):
    """Clean up and close all browser windows for the current WebDriver session."""
    if driver:
        logger.info("Initiating browser cleanup for current WebDriver session")
        try:
            # Switch to and close all open windows for this WebDriver instance
            for handle in driver.window_handles:
                try:
                    driver.switch_to.window(handle)
                    logger.info(f"Closing window with handle: {handle}")
                    driver.close()
                except (InvalidSessionIdException, Exception) as e:
                    logger.error(f"Error closing window {handle}: {str(e)}")
            # Quit the driver to terminate the session
            driver.quit()
            logger.info("Browser session terminated")
        except (InvalidSessionIdException, Exception) as e:
            logger.error(f"Error during browser cleanup: {str(e)}")

def main():
    driver, wait = None, None
    try:
        # Initialize browser once before the loop
        driver, wait = setup_browser()
        logger.info("Browser initialized successfully")
    except Exception as e:
        logger.error(f"Failed to initialize browser: {str(e)}")
        cleanup_browser(driver)
        return

    try:
        while True:  # Outer loop for login and hourly restarts
            start_time = time.time()  # Track start time for hourly restart
            try:
                # Store the original window handle
                original_window = driver.current_window_handle
                logger.info("Original window handle stored")

                # Step 1: Open the login page
                logger.info("Opening login page")
                driver.get("https://10superbo.com/page/mmg/login.jsp")
                wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/form")))
                logger.info("Login page loaded")
                time.sleep(2)

                # Step 2: Enter username
                logger.info("Entering username")
                username_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/form/div[2]/div/input")))
                username_field.clear()
                username_field.send_keys("superbabot7")
                wait.until(EC.text_to_be_present_in_element_value((By.XPATH, "/html/body/div[2]/div/form/div[2]/div/input"), "superbabot7"))
                logger.info("Username entered")
                time.sleep(1)

                # Step 3: Enter password
                logger.info("Entering password")
                password_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[2]/div/form/div[3]/div/input")))
                password_field.clear()
                password_field.send_keys("welcome1")
                wait.until(EC.text_to_be_present_in_element_value((By.XPATH, "/html/body/div[2]/div/form/div[3]/div/input"), "welcome1"))
                logger.info("Password entered")
                time.sleep(1)

                # Step 4: Click login button
                logger.info("Clicking login button")
                login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[2]/div/form/div[6]/a")))
                driver.execute_script("arguments[0].click();", login_button)
                logger.info("Login button clicked")
                time.sleep(2)

                # Verify login success
                try:
                    wait.until(EC.presence_of_element_located((By.TAG_NAME, "body")))
                    logger.info("Login successful")
                except:
                    logger.error("Login failed")
                    raise Exception("Login failed")

                while True:  # Inner loop for withdrawal table processing
                    try:
                        # Ensure we're in the original window before navigating
                        if driver.current_window_handle != original_window:
                            logger.info("Switching back to original window before navigating to withdrawal page")
                            driver.switch_to.window(original_window)

                        # Step 5: Navigate to withdrawal page and force refresh
                        logger.info("Navigating to withdrawal page")
                        try:
                            driver.get("https://10superbo.com/page/mmg/payment/withdrawal.jsp")
                            driver.refresh()  # Force page refresh to ensure fresh data
                            wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form")))
                            logger.info("Withdrawal page loaded")
                            time.sleep(2)
                        except Exception as e:
                            logger.error(f"Error in withdrawal page navigation: {str(e)}")
                            logger.info("Closing all windows and restarting from login after 5 seconds")
                            cleanup_browser(driver)
                            time.sleep(ERROR_RESTART_WAIT)  # Wait before restarting
                            driver, wait = setup_browser()  # Reinitialize browser
                            raise Exception("Withdrawal page navigation failed, triggering full restart")

                        # Step 6: Select 'Pending' from status dropdown using JS
                        logger.info("Selecting 'Pending' from status dropdown")
                        status_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form/div[2]/div/div[2]/select")))
                        driver.execute_script("arguments[0].value = '0'; arguments[0].dispatchEvent(new Event('change'));", status_dropdown)
                        wait.until(EC.text_to_be_present_in_element_value((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form/div[2]/div/div[2]/select"), "0"))
                        logger.info("Status set to Pending")
                        time.sleep(1)

                        # Step 7: Enter 100 in Withdraw From
                        logger.info("Entering 100 in Withdraw From")
                        withdraw_from_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form/div[6]/div[2]/div[2]/input")))
                        withdraw_from_field.clear()
                        withdraw_from_field.send_keys("100")
                        wait.until(EC.text_to_be_present_in_element_value((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form/div[6]/div[2]/div[2]/input"), "100"))
                        logger.info("Withdraw From set to 100")
                        time.sleep(1)

                        # Step 8: Enter 5000 in Withdraw To
                        logger.info("Entering 5000 in Withdraw To")
                        withdraw_to_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form/div[6]/div[2]/div[3]/input")))
                        withdraw_to_field.clear()
                        withdraw_to_field.send_keys("5000")
                        wait.until(EC.text_to_be_present_in_element_value((By.XPATH, "/html/body/div[5]/div[2]/div/div[10]/div/div/div[2]/form/div[6]/div[2]/div[3]/input"), "5000"))
                        logger.info("Withdraw To set to 5000")
                        time.sleep(1)

                        # Step 9: Click Search button
                        logger.info("Clicking Search button")
                        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[11]/div/div/input")))
                        driver.execute_script("arguments[0].click();", search_button)
                        logger.info("Search button clicked")
                        time.sleep(2)

                        # Step 10: Set records per page to 100 using JS
                        logger.info("Setting 100 records per page")
                        records_select = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#withdrawalSearchTable_length select")))
                        driver.execute_script("arguments[0].value = '100'; arguments[0].dispatchEvent(new Event('change'));", records_select)
                        logger.info("Records per page set to 100")
                        time.sleep(2)

                        # Check if withdrawal table exists
                        table_xpath = "/html/body/div[5]/div[2]/div/div[12]/div/div/div[2]/div/table/tbody"
                        try:
                            wait.until(EC.presence_of_element_located((By.XPATH, table_xpath + "/tr")))
                            table_elements = driver.find_elements(By.XPATH, table_xpath)
                            if not table_elements:
                                logger.warning("Withdrawal table not found or empty, retrying from withdrawal page")
                                time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                                continue  # Restart from Step 5
                            logger.info("Withdrawal table found, proceeding to process rows")
                        except Exception as e:
                            logger.error(f"Error checking withdrawal table: {str(e)}")
                            time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                            continue  # Restart from Step 5

                        # Step 11: Check Executor and click Username
                        processed_rows = []
                        row = 1
                        max_rows = 100  # Maximum rows to check

                        # Get the number of rows in the table
                        rows = driver.find_elements(By.XPATH, f"{table_xpath}/tr")
                        logger.info(f"Found {len(rows)} rows in withdrawal table")
                        max_rows = min(max_rows, len(rows))  # Adjust max_rows to actual table size

                        while row <= max_rows:
                            # Check for hourly restart
                            elapsed_time = time.time() - start_time
                            if elapsed_time >= HOUR_IN_SECONDS:
                                logger.info(f"One hour elapsed ({elapsed_time:.2f} seconds), restarting from login after 5-minute wait")
                                raise Exception("Hourly restart triggered")

                            executor_xpath = f"/html/body/div[5]/div[2]/div/div[12]/div/div/div[2]/div/table/tbody/tr[{row}]/td[27]"
                            brand_xpath = f"/html/body/div[5]/div[2]/div/div[12]/div/div/div[2]/div/table/tbody/tr[{row}]/td[2]"
                            username_xpath = f"/html/body/div[5]/div[2]/div/div[12]/div/div/div[2]/div/table/tbody/tr[{row}]/td[3]/a"

                            try:
                                # Check if row is a multiple of 20 (e.g., 20, 40, 60, etc.)
                                if row % 20 == 0:
                                    logger.info(f"Row {row} is a multiple of 20, clicking Search button")
                                    try:
                                        search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[11]/div/div/input")))
                                        driver.execute_script("arguments[0].click();", search_button)
                                        logger.info("Search button clicked")
                                        time.sleep(2)
                                        # Re-check table existence after search
                                        try:
                                            wait.until(EC.presence_of_element_located((By.XPATH, table_xpath + "/tr")))
                                            table_elements = driver.find_elements(By.XPATH, table_xpath)
                                            if not table_elements:
                                                logger.warning("Withdrawal table not found or empty after search, retrying from withdrawal page")
                                                time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                                                break  # Exit row processing loop and restart from Step 5
                                            logger.info("Withdrawal table found after search, continuing with row processing")
                                            # Update max_rows after search
                                            rows = driver.find_elements(By.XPATH, f"{table_xpath}/tr")
                                            max_rows = min(max_rows, len(rows))  # Adjust max_rows to new table size
                                        except Exception as e:
                                            logger.error(f"Error checking withdrawal table after search: {str(e)}")
                                            time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                                            break  # Exit row processing loop and restart from Step 5
                                    except Exception as e:
                                        logger.error(f"Error clicking Search button at row {row}: {str(e)}")
                                        time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                                        break  # Exit row processing loop and restart from Step 5

                                # Check if the row exists before proceeding
                                row_elements = driver.find_elements(By.XPATH, f"{table_xpath}/tr[{row}]")
                                if not row_elements:
                                    logger.info(f"No more rows found at row {row}, restarting from withdrawal page")
                                    time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                                    break  # Exit row processing loop and restart from Step 5

                                # Ensure we're in the original window before checking Executor
                                if driver.current_window_handle != original_window:
                                    logger.info("Switching back to original window before checking Executor")
                                    driver.switch_to.window(original_window)
                                    logger.info("Switched back to original window")

                                # Check Executor cell with a shorter timeout
                                executor_wait = WebDriverWait(driver, EXECUTOR_CHECK_WAIT)  # Use 3-second timeout
                                executor_cell = executor_wait.until(EC.presence_of_element_located((By.XPATH, executor_xpath)))
                                executor_text = executor_cell.text.strip()
                                logger.info(f"Row {row}: Executor text: '{executor_text}'")

                                # Extract Brand
                                brand = "Unknown"
                                try:
                                    brand_element = wait.until(EC.presence_of_element_located((By.XPATH, brand_xpath)))
                                    brand = brand_element.text.strip()
                                    logger.info(f"Brand: {brand}")
                                except Exception as e:
                                    logger.error(f"Error extracting Brand for row {row}: {str(e)}")
                                    # Continue processing the row even if Brand extraction fails

                                click_condition_met = False
                                if not executor_text or executor_text.isspace():  # If Executor is blank or contains only whitespace
                                    click_condition_met = True
                                    logger.info(f"Executor is blank for row {row}")
                                elif row == 1:  # Special case for row 1
                                    click_condition_met = True
                                    logger.info("Processing row 1 (special case)")

                                if click_condition_met:
                                    # Extract username
                                    username = "Unknown"
                                    try:
                                        username_element = wait.until(EC.presence_of_element_located((By.XPATH, username_xpath)))
                                        username = username_element.text.strip()
                                        logger.info(f"Username Row {row}: {username}")
                                    except Exception as e:
                                        logger.error(f"Error extracting username for row {row}: {str(e)}")

                                    try:
                                        username_link = wait.until(EC.element_to_be_clickable((By.XPATH, username_xpath)))
                                        driver.execute_script("arguments[0].click();", username_link)
                                        logger.info(f"Clicked Username for row {row}: {username}")
                                        time.sleep(2)

                                        # Wait for new window to open
                                        wait.until(EC.number_of_windows_to_be(2))
                                        all_windows = driver.window_handles
                                        new_window = [window for window in all_windows if window != original_window][0]
                                        driver.switch_to.window(new_window)
                                        logger.info("Switched to new window")

                                        # Step 12: Extract text from the specified XPath on the new page
                                        logger.info("Extracting text from new page")
                                        group_type = "No"  # Default
                                        try:
                                            target_element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[8]/div[1]/div/div/div[2]/table/tbody/tr[9]/td[1]")))
                                            text_value = target_element.text.strip()
                                            if text_value in ["None999"]:
                                                group_type = text_value
                                            logger.info(f"Extracted text/value from XPath: Group: {group_type}")
                                        except Exception as e:
                                            logger.info("Extracted text/value from XPath: Group: No")
                                            logger.error(f"Error locating group type element: {str(e)}")
                                            # Close all windows and restart from beginning
                                            cleanup_browser(driver)
                                            logger.info("Critical error in group type extraction, restarting from login after 5 seconds")
                                            time.sleep(ERROR_RESTART_WAIT)
                                            raise Exception("Group type extraction failed, triggering full restart")

                                        # Conditional action based on group_type
                                        if group_type == "No":
                                            # Extract currency
                                            logger.info("Extracting currency")
                                            currency = "Unknown"  # Default
                                            try:
                                                currency_element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[8]/div[1]/div/div/div[2]/table/tbody/tr[13]/td[1]")))
                                                currency = currency_element.text.strip()
                                                logger.info(f"Extracted currency: {currency}")
                                            except Exception as e:
                                                logger.info("Failed to extract currency, assuming not USD")
                                                logger.error(f"Error locating currency element: {str(e)}")
                                                # Close new window and switch back to original
                                                driver.close()
                                                driver.switch_to.window(original_window)
                                                logger.info("Closed new window and switched back to original window")
                                                time.sleep(TABLE_CHECK_WAIT)
                                                row += 1  # Move to next row
                                                continue  # Continue to next row

                                            if currency == "USD":
                                                logger.info("Currency is USD, skipping further checks")
                                                # Close new window and switch back to original
                                                try:
                                                    driver.close()
                                                    driver.switch_to.window(original_window)
                                                    logger.info("Closed new window")
                                                    logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row
                                                except Exception as window_e:
                                                    logger.error(f"Error switching windows: {str(window_e)}")
                                                    driver.switch_to.window(original_window)
                                                    logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                            # Continue with remaining checks if currency is not USD
                                            # Click down icon
                                            logger.info("Clicking down icon")
                                            down_icon = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[8]/div[1]/div/div/div[1]/div/div/span")))
                                            driver.execute_script("arguments[0].click();", down_icon)
                                            logger.info("Down icon clicked")
                                            time.sleep(2)

                                            # Extract deposit amount
                                            logger.info("Extracting deposit amount")
                                            deposit = "Unknown"  # Default
                                            try:
                                                deposit_element = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[8]/div[4]/div/div/div[2]/table/tbody/tr[1]/td[1]")))
                                                deposit = deposit_element.text.strip()
                                                logger.info(f"Extracted deposit amount: {deposit}")
                                            except Exception as e:
                                                logger.info("Failed to extract deposit amount, assuming not 0.00")
                                                logger.error(f"Error locating deposit element: {str(e)}")
                                                # Close new window and switch back to original
                                                driver.close()
                                                driver.switch_to.window(original_window)
                                                logger.info("Closed new window and switched back to original window")
                                                time.sleep(TABLE_CHECK_WAIT)
                                                row += 1  # Move to next row
                                                continue  # Continue to next row

                                            if deposit != "0.00":
                                                # Click Report if not 0.00
                                                logger.info("Deposit is not 0.00, clicking Report")
                                                report_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[1]/div[1]/div[1]/ul/li[3]/a")))
                                                driver.execute_script("arguments[0].click();", report_button)
                                                logger.info("Report button clicked")
                                                time.sleep(2)

                                                # Click Bonus after Report
                                                logger.info("Clicking Bonus")
                                                bonus_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/ul/li[2]/a")))
                                                driver.execute_script("arguments[0].click();", bonus_button)
                                                logger.info("Bonus button clicked")
                                                time.sleep(2)

                                                # Set records per page to 50 using JS
                                                logger.info("Setting records per page to 50")
                                                records_select = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[7]/div[1]/div/div/div/div/div[1]/div[4]/div/label/select")))
                                                driver.execute_script("arguments[0].value = '50'; arguments[0].dispatchEvent(new Event('change'));", records_select)
                                                logger.info("Records per page set to 50")
                                                time.sleep(2)

                                                # Step 13: Check New 50 Bonus Rows
                                                logger.info("Checking New 50 Bonus rows")
                                                bonus_keywords = [
                                                    "Rebate", "Cash Back", "Rebates", "Unlimited", "Deposit Bonus", "Rolling Reward", "Sports Daily CB", "Epic Win Rush", "JDB FORTUNES", "Sports Weekly CB", "Rebate- Final", "Welcome Back", "Welcome Call", "Bang Bang Bonus", "Lucky Tiger Bonus", "REAL TIME REBATE", "Refer A Friend Bonus", "REAL-TIME Bonus", "Slots-Manual", "Elements- Manual", "Super Elements", "Welcome Call Deposit", "REAL-TIME Bonus", " Tiger Free Spins", "EXCLUSIVE GAME", "০.৩০% স্পোর্টস রিবেট", "Real Time", "Rebate - Manual", "Manual", "RAF Monthly Achievement", " Major Vitality Contest", "Product Rewards", "No_Bonus", "NoBonus", "Birthday", "Real Time Bonus", "CashBack",
                                                    "Invite a Friend", "New Refer Bonus", "Deposit & Get 20 JILI Free Spin ", "Normal Deposit", "Daily Login", "Commission", "Unlimited Deposit"
                                                ]
                                                bonus_statuses = []
                                                for bonus_row in range(1, 26):  # Check rows 1 to 25
                                                    bonus_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[7]/div[1]/div/div/div/div/table/tbody/tr[{bonus_row}]/td[3]"
                                                    try:
                                                        elements = driver.find_elements(By.XPATH, bonus_xpath)
                                                        if not elements:
                                                            logger.info(f"Row {bonus_row}: No more rows found, stopping bonus check (processed {len(bonus_statuses)} rows)")
                                                            break
                                                        bonus_text = elements[0].text.strip()
                                                        logger.info(f"Row {bonus_row}: Extracted bonus text: {bonus_text}")
                                                        is_keyword_present = any(keyword.lower() in bonus_text.lower() for keyword in bonus_keywords)
                                                        status = "Yes" if is_keyword_present else "No"
                                                        logger.info(f"Row {bonus_row}: Status: {status}")
                                                        bonus_statuses.append(status)
                                                    except StaleElementReferenceException:
                                                        logger.warning(f"Stale element at bonus row {bonus_row}, retrying")
                                                        time.sleep(1)
                                                        continue
                                                    except Exception as e:
                                                        logger.error(f"Error checking bonus row {bonus_row}: {str(e)}")
                                                        # Close new window and switch back to original
                                                        try:
                                                            driver.close()
                                                            driver.switch_to.window(original_window)
                                                            logger.info("Closed new window and switched back to original window")
                                                        except Exception as window_e:
                                                            logger.error(f"Error switching windows: {str(window_e)}")
                                                            driver.switch_to.window(original_window)
                                                            logger.info("Switched back to original window")
                                                        time.sleep(TABLE_CHECK_WAIT)
                                                        row += 1  # Move to next row
                                                        continue  # Continue to next row

                                                # Check if all rows are "Yes" or table is empty
                                                if not bonus_statuses:
                                                    logger.info("No bonus rows found, proceeding to click Deposit tab")
                                                    all_yes = True
                                                else:
                                                    all_yes = all(status == "Yes" for status in bonus_statuses)
                                                    logger.info(f"Processed {len(bonus_statuses)} bonus rows. All rows are Yes: {all_yes}")

                                                # If all_yes is False, move to next row
                                                if not all_yes:
                                                    logger.info("Not all bonus rows are Yes, moving to next row")
                                                    try:
                                                        if len(driver.window_handles) > 1:
                                                            driver.close()
                                                            logger.info("Closed new window")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                        time.sleep(TABLE_CHECK_WAIT)
                                                        row += 1  # Move to next row
                                                        continue  # Continue to next row
                                                    except Exception as e:
                                                        logger.error(f"Error switching windows: {str(e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                        time.sleep(TABLE_CHECK_WAIT)
                                                        row += 1  # Move to next row
                                                        continue  # Continue to next row

                                                logger.info("Clicking Deposit tab")
                                                try:
                                                    deposit_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/ul/li[7]/a")))
                                                    driver.execute_script("arguments[0].click();", deposit_button)
                                                    logger.info("Deposit tab clicked")
                                                    time.sleep(1)
                                                except Exception as e:
                                                    logger.error(f"Error clicking Deposit tab: {str(e)}")
                                                    # Close new window and switch back to original
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                logger.info("Clicking Report in Deposit tab")
                                                try:
                                                    report_button_deposit = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/ul/li[10]/a")))
                                                    driver.execute_script("arguments[0].click();", report_button_deposit)
                                                    logger.info("Report in Deposit tab clicked")
                                                    time.sleep(2)
                                                except Exception as e:
                                                    logger.error(f"Error clicking Report in Deposit tab: {str(e)}")
                                                    # Close new window and switch back to original
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                logger.info("Selecting Transaction Type to Bet")
                                                try:
                                                    transaction_type_dropdown = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/form[1]/div[1]/div/div/select")))
                                                    driver.execute_script("arguments[0].value = '0'; arguments[0].dispatchEvent(new Event('change'));", transaction_type_dropdown)
                                                    wait.until(EC.text_to_be_present_in_element_value((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/form[1]/div[1]/div/div/select"), "0"))
                                                    logger.info("Transaction Type set to Bet")
                                                    time.sleep(1)
                                                except Exception as e:
                                                    logger.error(f"Error selecting Transaction Type to Bet: {str(e)}")
                                                    # Close new window and switch back to original
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                logger.info("Clearing Transaction Date Range")
                                                try:
                                                    date_range_field = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/form[1]/div[2]/div[2]/div[1]/input")))
                                                    date_range_field.clear()
                                                    logger.info("Transaction Date Range cleared")
                                                    time.sleep(1)
                                                except Exception as e:
                                                    logger.error(f"Error clearing Transaction Date Range: {str(e)}")
                                                    # Close new window and switch back to original
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                logger.info("Clicking Report tab again")
                                                try:
                                                    report_button_again = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/ul/li[10]/a")))
                                                    driver.execute_script("arguments[0].click();", report_button_again)
                                                    logger.info("Report tab clicked again")
                                                    time.sleep(1)
                                                except Exception as e:
                                                    logger.error(f"Error clicking Report tab again: {str(e)}")
                                                    # Close new window and switch back to original
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                logger.info("Clicking Search button")
                                                try:
                                                    search_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/form[1]/div[3]/div/div/input")))
                                                    driver.execute_script("arguments[0].click();", search_button)
                                                    logger.info("Search button clicked")
                                                    time.sleep(2)
                                                except Exception as e:
                                                    logger.error(f"Error clicking Search button: {str(e)}")
                                                    # Close new window and switch back to original
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                # Additional steps: Check View Icons for rows 1 to 5
                                                all_bet_yes_flags = []  # Track all_bet_yes for processed rows only
                                                processed_view_rows = 0  # Count of processed view rows
                                                for view_row in range(1, 6):  # Check rows 1 to 5
                                                    logger.info(f"Attempting to locate View Icon for row {view_row}")
                                                    view_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[1]/div/div/div/div/table/tbody/tr[{view_row}]/td[7]/a"
                                                    fallback_xpath = f"//table/tbody/tr[{view_row}]/td[7]/a"
                                                    # Check if view_row exists
                                                    elements = driver.find_elements(By.XPATH, view_xpath)
                                                    if not elements:
                                                        logger.info(f"View Icon for row {view_row} not found with primary XPath, trying fallback XPath")
                                                        elements = driver.find_elements(By.XPATH, fallback_xpath)
                                                        if not elements:
                                                            logger.info(f"View Icon for row {view_row} not found, stopping View Icon checks (processed {processed_view_rows} rows)")
                                                            break  # Exit View Icon loop to continue with next row

                                                    view_icon = wait.until(EC.element_to_be_clickable((By.XPATH, view_xpath)))
                                                    logger.info(f"View Icon for row {view_row} located, attempting to click")
                                                    driver.execute_script("arguments[0].scrollIntoView(true);", view_icon)
                                                    time.sleep(0.5)  # Brief pause after scrolling
                                                    driver.execute_script("arguments[0].click();", view_icon)
                                                    logger.info(f"View Icon clicked for row {view_row}")
                                                    time.sleep(2)

                                                    # Check Type rows 1 to 20 for view_row
                                                    logger.info(f"Checking Type rows 1 to 20 for View Icon row {view_row}")
                                                    all_bet_yes = True
                                                    for type_row in range(1, 21):
                                                        type_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/table/tbody/tr[{type_row}]/td[3]"
                                                        turnover_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/table/tbody/tr[{type_row}]/td[6]"
                                                        bet_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/table/tbody/tr[{type_row}]/td[4]"
                                                        win_loss_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/table/tbody/tr[{type_row}]/td[5]"
                                                        win_loss_negative_xpath = f"/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[2]/div/table/tbody/tr[{type_row}]/td[5]/span"
                                                        elements = driver.find_elements(By.XPATH, type_xpath)
                                                        if not elements:
                                                            logger.info(f"Row {type_row}: No more rows found, stopping type check for View Icon row {view_row}")
                                                            break
                                                        type_text = elements[0].text.strip().lower()
                                                        logger.info(f"Row {type_row}: Extracted type text: {type_text}")
                                                        
                                                        if 'slot' in type_text or 'free' in type_text or 'Slots' in type_text or 'fishing' in type_text:
                                                            logger.info(f"Row {type_row}: Yes (Type is slot, free, Slots, or fishing)")
                                                        elif 'crash' in type_text:
                                                            # Check turnover for crash type
                                                            try:
                                                                turnover_element = driver.find_element(By.XPATH, turnover_xpath)
                                                                turnover_text = turnover_element.text.strip()
                                                                try:
                                                                    # Remove commas from turnover value
                                                                    turnover_value = float(turnover_text.replace(',', ''))
                                                                    if turnover_value <= 200:
                                                                        logger.info(f"Row {type_row}: Yes (Turnover {turnover_value} is <= 200)")
                                                                    else:
                                                                        # For crash with turnover > 200, calculate (Win/Loss / Bet) * 100 or * (-100)
                                                                        try:
                                                                            bet_element = driver.find_element(By.XPATH, bet_xpath)
                                                                            bet_text = bet_element.text.strip()
                                                                            # Remove commas from bet value
                                                                            bet_value = float(bet_text.replace(',', ''))
                                                                            if bet_value == 0:
                                                                                logger.info(f"Row {type_row}: No (Bet is 0, cannot calculate ratio)")
                                                                                all_bet_yes = False
                                                                                continue

                                                                            # Check if Win/Loss is negative by looking for span element
                                                                            win_loss_negative_elements = driver.find_elements(By.XPATH, win_loss_negative_xpath)
                                                                            if win_loss_negative_elements:
                                                                                # Negative Win/Loss
                                                                                win_loss_element = driver.find_element(By.XPATH, win_loss_negative_xpath)
                                                                                win_loss_text = win_loss_element.text.strip()
                                                                                try:
                                                                                    # Clean Win/Loss: remove spaces, parentheses, and commas
                                                                                    cleaned_win_loss = win_loss_text.strip('() ').replace(',', '')
                                                                                    win_loss_value = -float(cleaned_win_loss)  # Negate for negative value
                                                                                    ratio = (win_loss_value / bet_value) * (-100)
                                                                                    logger.info(f"Row {type_row}: Calculated negative ratio for crash: {ratio:.2f}")
                                                                                    if ratio >= 20:
                                                                                        logger.info(f"Row {type_row}: Yes (Ratio {ratio:.2f} is >= 20)")
                                                                                    else:
                                                                                        logger.info(f"Row {type_row}: No (Ratio {ratio:.2f} is < 20)")
                                                                                        all_bet_yes = False
                                                                                except ValueError:
                                                                                    logger.error(f"Row {type_row}: Invalid negative Win/Loss value '{win_loss_text}', assuming No")
                                                                                    all_bet_yes = False
                                                                            else:
                                                                                # Positive or zero Win/Loss
                                                                                win_loss_element = driver.find_element(By.XPATH, win_loss_xpath)
                                                                                win_loss_text = win_loss_element.text.strip()
                                                                                try:
                                                                                    # Clean Win/Loss: remove commas
                                                                                    win_loss_value = float(win_loss_text.replace(',', ''))
                                                                                    ratio = (win_loss_value / bet_value) * 100
                                                                                    logger.info(f"Row {type_row}: Calculated positive ratio for crash: {ratio:.2f}")
                                                                                    if ratio >= 20:
                                                                                        logger.info(f"Row {type_row}: Yes (Turnover {turnover_value} is <= 200)")
                                                                                    else:
                                                                                        logger.info(f"Row {type_row}: No (Ratio {ratio:.2f} is < 20)")
                                                                                        all_bet_yes = False
                                                                                except ValueError:
                                                                                    logger.error(f"Row {type_row}: Invalid positive Win/Loss value '{win_loss_text}', assuming No")
                                                                                    all_bet_yes = False
                                                                        except Exception as e:
                                                                            logger.error(f"Row {type_row}: Error accessing Bet or Win/Loss: {str(e)}, assuming No")
                                                                            all_bet_yes = False
                                                                except ValueError:
                                                                    logger.error(f"Row {type_row}: Invalid turnover value '{turnover_text}', assuming No")
                                                                    all_bet_yes = False
                                                            except Exception as e:
                                                                logger.error(f"Row {type_row}: Error accessing turnover: {str(e)}, assuming No")
                                                                all_bet_yes = False
                                                        else:
                                                            # For non-slot/free/Slots/fishing/crash types, check turnover
                                                            try:
                                                                turnover_element = driver.find_element(By.XPATH, turnover_xpath)
                                                                turnover_text = turnover_element.text.strip()
                                                                try:
                                                                    # Remove commas from turnover value
                                                                    turnover_value = float(turnover_text.replace(',', ''))
                                                                    if turnover_value <= 200:
                                                                        logger.info(f"Row {type_row}: Yes (Turnover {turnover_value} is <= 200)")
                                                                    else:
                                                                        # Calculate (Win/Loss / Bet) * 100 for other types with turnover > 200
                                                                        try:
                                                                            bet_element = driver.find_element(By.XPATH, bet_xpath)
                                                                            bet_text = bet_element.text.strip()
                                                                            # Remove commas from bet value
                                                                            bet_value = float(bet_text.replace(',', ''))
                                                                            if bet_value == 0:
                                                                                logger.info(f"Row {type_row}: No (Bet is 0, cannot calculate ratio)")
                                                                                all_bet_yes = False
                                                                                continue

                                                                            # Check if Win/Loss is negative by looking for span element
                                                                            win_loss_negative_elements = driver.find_elements(By.XPATH, win_loss_negative_xpath)
                                                                            if win_loss_negative_elements:
                                                                                # Negative Win/Loss
                                                                                win_loss_element = driver.find_element(By.XPATH, win_loss_negative_xpath)
                                                                                win_loss_text = win_loss_element.text.strip()
                                                                                try:
                                                                                    # Clean Win/Loss: remove spaces, parentheses, and commas
                                                                                    cleaned_win_loss = win_loss_text.strip('() ').replace(',', '')
                                                                                    win_loss_value = -float(cleaned_win_loss)  # Negate for negative value
                                                                                    ratio = (win_loss_value / bet_value) * 100
                                                                                    logger.info(f"Row {type_row}: Calculated ratio for other type: {ratio:.2f}")
                                                                                    if -100 <= ratio <= -99:
                                                                                        logger.info(f"Row {type_row}: Yes (Ratio {ratio:.2f} is between -100 and -99)")
                                                                                    else:
                                                                                        logger.info(f"Row {type_row}: No (Ratio {ratio:.2f} is not between -100 and -99)")
                                                                                        all_bet_yes = False
                                                                                except ValueError:
                                                                                    logger.error(f"Row {type_row}: Invalid negative Win/Loss value '{win_loss_text}', assuming No")
                                                                                    all_bet_yes = False
                                                                            else:
                                                                                # Positive or zero Win/Loss
                                                                                win_loss_element = driver.find_element(By.XPATH, win_loss_xpath)
                                                                                win_loss_text = win_loss_element.text.strip()
                                                                                try:
                                                                                    # Clean Win/Loss: remove commas
                                                                                    win_loss_value = float(win_loss_text.replace(',', ''))
                                                                                    ratio = (win_loss_value / bet_value) * 100
                                                                                    logger.info(f"Row {type_row}: Calculated ratio for other type: {ratio:.2f}")
                                                                                    if -100 <= ratio <= -99:
                                                                                        logger.info(f"Row {type_row}: Yes (Ratio {ratio:.2f} is between -100 and -99)")
                                                                                    else:
                                                                                        logger.info(f"Row {type_row}: No (Ratio {ratio:.2f} is not between -100 and -99)")
                                                                                        all_bet_yes = False
                                                                                except ValueError:
                                                                                    logger.error(f"Row {type_row}: Invalid positive Win/Loss value '{win_loss_text}', assuming No")
                                                                                    all_bet_yes = False
                                                                        except Exception as e:
                                                                            logger.error(f"Row {type_row}: Error accessing Bet or Win/Loss: {str(e)}, assuming No")
                                                                            all_bet_yes = False
                                                                except ValueError:
                                                                    logger.error(f"Row {type_row}: Invalid turnover value '{turnover_text}', assuming No")
                                                                    all_bet_yes = False
                                                            except Exception as e:
                                                                logger.error(f"Row {type_row}: Error accessing turnover: {str(e)}, assuming No")
                                                                all_bet_yes = False

                                                    all_bet_yes_flags.append(all_bet_yes)
                                                    processed_view_rows += 1
                                                    if all_bet_yes:
                                                        logger.info(f"All type rows are Yes for View Icon row {view_row}")
                                                    else:
                                                        logger.info(f"Not all type rows are Yes for View Icon row {view_row}")
                                                        # Close the View Icon window and break to continue with next row
                                                        try:
                                                            close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/button")))
                                                            driver.execute_script("arguments[0].click();", close_button)
                                                            logger.info(f"Close button clicked for row {view_row}")
                                                            time.sleep(1)
                                                        except Exception as e:
                                                            logger.error(f"Error closing View Icon Button for row {view_row}: {str(e)}")
                                                        break  # Exit View Icon loop to continue with next row

                                                    # Close the View Icon Button
                                                    logger.info(f"Closing View Icon Button for row {view_row}")
                                                    try:
                                                        close_button = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div[5]/div[2]/div/div[3]/div/div/div/div/div/div[1]/div[2]/div/div/div[1]/button")))
                                                        driver.execute_script("arguments[0].click();", close_button)
                                                        logger.info(f"Close button clicked for row {view_row}")
                                                        time.sleep(1)
                                                    except Exception as e:
                                                        logger.error(f"Error closing View Icon Button for row {view_row}: {str(e)}")
                                                        break  # Exit View Icon loop to continue with next row

                                                # After processing View Icons (or early exit), close the new window
                                                try:
                                                    if len(driver.window_handles) > 1:
                                                        driver.close()
                                                        logger.info("Closed new window")
                                                    driver.switch_to.window(original_window)
                                                    logger.info("Switched back to original window")
                                                    time.sleep(2)
                                                except Exception as e:
                                                    logger.error(f"Error switching windows: {str(e)}")
                                                    try:
                                                        driver.close()
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Closed new window and switched back to original window")
                                                    except Exception as window_e:
                                                        logger.error(f"Error switching windows: {str(window_e)}")
                                                        driver.switch_to.window(original_window)
                                                        logger.info("Switched back to original window")
                                                    time.sleep(TABLE_CHECK_WAIT)
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                                # Verify only if at least one View Icon row was processed and all processed View Icon rows are True
                                                if processed_view_rows > 0 and all(all_bet_yes_flags):
                                                    logger.info(f"Attempting to click Verify button for withdrawal table row {row}: {username}")
                                                    verify_xpath = f"/html/body/div[5]/div[2]/div/div[12]/div/div/div[2]/div/table/tbody/tr[{row}]/td[28]/a[1]"
                                                    try:
                                                        verify_button = wait.until(EC.element_to_be_clickable((By.XPATH, verify_xpath)))
                                                        driver.execute_script("arguments[0].scrollIntoView(true);", verify_button)
                                                        time.sleep(0.5)  # Brief pause after scrolling
                                                        driver.execute_script("arguments[0].click();", verify_button)
                                                        logger.info(f"Verify button clicked for row {row}: {username}")

                                                        # Attempt to click Confirm Verify button
                                                        logger.info("Attempting to click Confirm Verify button")
                                                        confirm_verify_xpath = "/html/body/div[5]/div[2]/div/div[4]/div/div/div[3]/input[4]"
                                                        try:
                                                            confirm_verify_button = WebDriverWait(driver, CONFIRM_VERIFY_WAIT).until(
                                                                EC.element_to_be_clickable((By.XPATH, confirm_verify_xpath))
                                                            )
                                                            driver.execute_script("arguments[0].click();", confirm_verify_button)
                                                            logger.info("🎉🎉🎉 Confirm Verify button clicked🎉🎉🎉")
                                                            time.sleep(2)  # Allow time for confirmation to process
                                                            processed_rows.append(row)  # Mark row as processed
                                                            # Close any open new windows and switch back to original
                                                            try:
                                                                if len(driver.window_handles) > 1:
                                                                    driver.close()
                                                                    logger.info("Closed new window")
                                                                driver.switch_to.window(original_window)
                                                                logger.info("Switched back to original window")
                                                            except Exception as window_e:
                                                                logger.error(f"Error switching windows after Confirm Verify: {str(window_e)}")
                                                                driver.switch_to.window(original_window)
                                                                logger.info("Switched back to original window")
                                                        except (NoSuchElementException, TimeoutException) as e:
                                                            logger.warning(f"Confirm Verify button not found for row {row}, moving to next executor row: {str(e)}")
                                                            # Close any open new windows and switch back to original
                                                            try:
                                                                if len(driver.window_handles) > 1:
                                                                    driver.close()
                                                                    logger.info("Closed new window")
                                                                driver.switch_to.window(original_window)
                                                                logger.info("Switched back to original window")
                                                            except Exception as window_e:
                                                                logger.error(f"Error switching windows: {str(window_e)}")
                                                                driver.switch_to.window(original_window)
                                                                logger.info("Switched back to original window")
                                                            time.sleep(TABLE_CHECK_WAIT)
                                                            row += 1  # Move to next row
                                                            continue  # Continue to next executor row
                                                        except Exception as e:
                                                            logger.error(f"Unexpected error clicking Confirm Verify button for row {row}: {str(e)}")
                                                            # Close any open new windows and switch back to original
                                                            try:
                                                                if len(driver.window_handles) > 1:
                                                                    driver.close()
                                                                    logger.info("Closed new window")
                                                                driver.switch_to.window(original_window)
                                                                logger.info("Switched back to original window")
                                                            except Exception as window_e:
                                                                logger.error(f"Error switching windows: {str(window_e)}")
                                                                driver.switch_to.window(original_window)
                                                                logger.info("Switched back to original window")
                                                            time.sleep(TABLE_CHECK_WAIT)
                                                            row += 1  # Move to next row
                                                            continue  # Continue to next executor row
                                                    except Exception as e:
                                                        logger.error(f"Error clicking Verify button for row {row}: {str(e)}")
                                                        # Close new window and switch back to original
                                                        try:
                                                            driver.close()
                                                            driver.switch_to.window(original_window)
                                                            logger.info("Closed new window and switched back to original window")
                                                        except Exception as window_e:
                                                            logger.error(f"Error switching windows: {str(window_e)}")
                                                            driver.switch_to.window(original_window)
                                                            logger.info("Switched back to original window")
                                                        time.sleep(TABLE_CHECK_WAIT)
                                                        row += 1  # Move to next row
                                                        continue  # Continue to next row
                                                else:
                                                    logger.info(f"Either no View Icon rows processed or not all processed View Icon rows are Yes for row {row} (processed {processed_view_rows} rows), moving to next row")
                                                    row += 1  # Move to next row
                                                    continue  # Continue to next row

                                    except (NoSuchElementException, TimeoutException) as e:
                                        logger.error(f"Error accessing username link for row {row}: {str(e)}")
                                        # Close new window and switch back to original
                                        try:
                                            driver.close()
                                            driver.switch_to.window(original_window)
                                            logger.info("Closed new window and switched back to original window")
                                        except Exception as window_e:
                                            logger.error(f"Error switching windows: {str(window_e)}")
                                            driver.switch_to.window(original_window)
                                            logger.info("Switched back to original window")
                                        time.sleep(TABLE_CHECK_WAIT)
                                        row += 1  # Move to next row
                                        continue  # Continue to next row

                                # Increment row at the end of the loop iteration if click_condition_met is False
                                row += 1

                            except (NoSuchElementException, TimeoutException) as e:
                                logger.warning(f"Row {row} not found or timed out: {str(e)}, moving to next row")
                                row += 1  # Move to next row
                                continue  # Continue to next row
                            except StaleElementReferenceException:
                                logger.warning(f"Stale element at row {row}, retrying after {EXECUTOR_CHECK_WAIT} seconds")
                                time.sleep(EXECUTOR_CHECK_WAIT)  # Use shorter wait for stale element retry
                                continue  # Retry the same row
                            except Exception as e:
                                logger.error(f"Unexpected error processing row {row}: {str(e)}")
                                time.sleep(EXECUTOR_CHECK_WAIT)  # Use shorter wait for retry
                                row += 1  # Move to next row
                                continue  # Continue to next row

                        # After row processing, check if any rows were processed
                        if not processed_rows:
                            logger.info("No rows processed (no blank Executors or special row 1 condition met), retrying from withdrawal page")
                            time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                            continue  # Restart from Step 5

                        # If rows were processed, refresh the table and restart from Step 5
                        logger.info(f"Processed {len(processed_rows)} rows, refreshing withdrawal page")
                        time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                        continue  # Restart from Step 5

                    except UnexpectedAlertPresentException as e:
                        # Handle specific alert for being logged off
                        alert_text = str(e).lower()
                        if "you have been logged off because you have logged on at another location" in alert_text:
                            logger.error(f"Logged off alert detected: {str(e)}")
                            logger.info("Closing all windows and restarting from login immediately")
                            cleanup_browser(driver)
                            driver, wait = setup_browser()  # Reinitialize browser
                            continue  # Restart from login (outer loop)
                        else:
                            logger.error(f"Unexpected alert in withdrawal page processing: {str(e)}")
                            # Close any open new windows
                            try:
                                if len(driver.window_handles) > 1:
                                    driver.close()
                                    logger.info("Closed new window")
                                driver.switch_to.window(original_window)
                                logger.info("Switched back to original window")
                            except Exception as window_e:
                                logger.error(f"Error switching windows: {str(window_e)}")
                                driver.switch_to.window(original_window)
                                logger.info("Switched back to original window")
                            time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                            continue  # Restart from Step 5
                    except Exception as e:
                        if str(e) == "Hourly restart triggered":
                            raise  # Propagate hourly restart to outer loop
                        elif str(e) == "Withdrawal page navigation failed, triggering full restart":
                            raise  # Propagate withdrawal page navigation failure to outer loop
                        logger.error(f"Error in withdrawal page processing: {str(e)}")
                        logger.info(f"Retrying from withdrawal page after {TABLE_CHECK_WAIT} seconds")
                        # Ensure any open new windows are closed
                        try:
                            if len(driver.window_handles) > 1:
                                driver.close()
                                logger.info("Closed new window")
                            driver.switch_to.window(original_window)
                            logger.info("Switched back to original window")
                        except Exception as e:
                            logger.error(f"Error switching windows: {str(e)}")
                            driver.switch_to.window(original_window)
                            logger.info("Switched back to original window")
                        time.sleep(TABLE_CHECK_WAIT)  # Wait before retry
                        continue  # Restart from Step 5

            except Exception as e:
                logger.error(f"An error occurred, handling restart: {str(e)}")
                cleanup_browser(driver)
                if str(e) == "Hourly restart triggered":
                    logger.info(f"Hourly restart triggered, waiting {HOURLY_RESTART_WAIT} seconds before restarting from login")
                    time.sleep(HOURLY_RESTART_WAIT)  # 5-minute wait for hourly restart
                else:
                    logger.info(f"Non-hourly error, waiting {ERROR_RESTART_WAIT} seconds before restarting from login")
                    time.sleep(ERROR_RESTART_WAIT)  # 5-second wait for other errors
                driver, wait = setup_browser()  # Reinitialize browser
                continue  # Restart from login (outer loop)

    finally:
        cleanup_browser(driver)

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Script terminated by user")
        cleanup_browser(None)
    except Exception as e:
        logger.error(f"Unexpected error in main: {str(e)}")
        cleanup_browser(None)
    finally:

        cleanup_browser(None)
