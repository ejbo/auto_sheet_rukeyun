from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time

user_data_dir = r'C:\Users\jzl19\AppData\Local\Microsoft\Edge\User Data'
edge_options = webdriver.EdgeOptions()
edge_options.add_argument('--user-data-dir=' + user_data_dir)
edge_driver_path = r'C:\Coding\Environment\edgedriver_win64\msedgedriver.exe'

driver = webdriver.Edge(executable_path = edge_driver_path, options=edge_options)


driver.get('https://www.google.com/')
# driver.get('https://b.keruyun.com/bui-link/#/mind-ui/#/order/cashRegisterDetails')

time.sleep(10)

button_download_xpath = '//button[@class = "ant-btn ant-btn-primary ant-btn-background-ghost"]'
button_download_element = driver.find_element(By.XPATH, button_download_xpath)
button_download_element.click()

time.sleep(5)

main_window_handle = driver.current_window_handle
popup_window_handle = None
wait = WebDriverWait(driver, 5)  # Adjust the timeout as needed

for handle in driver.window_handles:
    if handle != main_window_handle:
        popup_window_handle = handle
        break

if popup_window_handle:
    driver.switch_to.window(popup_window_handle)

    # Interact with elements in the popup window
    popup_button_xpath = '//button[@class = "ant-btn ant-btn-primary"]'  # Replace with the appropriate XPath
    popup_button_element = wait.until(EC.presence_of_element_located((By.XPATH, popup_button_xpath)))
    popup_button_element.click()

    # Switch back to the main window
    driver.switch_to.window(main_window_handle)

time.sleep(5)
driver.quit()