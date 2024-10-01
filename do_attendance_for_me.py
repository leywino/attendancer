import time
import re
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime, timedelta
import calendar

# Load details from needy.txt
def load_credentials():
    with open('needy.txt', 'r') as file:
        data = file.readlines()
    credentials = {}
    for line in data:
        key, value = line.strip().split("=")
        credentials[key] = value.strip('"')
    return credentials

# Check if a specific day is one of the days to avoid
def is_attendance_needed(day, days_to_avoid):
    avoid_days = [int(day) for day in days_to_avoid.split(',')]
    return day not in avoid_days

# Initialize Selenium WebDriver
service = Service('C:/Users/leywi/OneDrive/Documents/development/chromedriver-win64/chromedriver.exe')  # Update this to your ChromeDriver path
driver = webdriver.Chrome(service=service)

def login(credentials):
    driver.get('https://www.resoluteonline.in/Login')

    # Enter username and password
    driver.find_element(By.ID, "MainContent_UserName").send_keys(credentials['website_username'])
    driver.find_element(By.ID, "MainContent_Password").send_keys(credentials['website_password'])
    
    # Click login
    driver.find_element(By.ID, "MainContent_LoginButton").click()

    # Wait for OTP input
    print("Waiting for OTP input...")
    WebDriverWait(driver, 300).until(EC.url_contains("/Home"))

def navigate_to_schedule():
    # Replace 'Home' with 'Schedule' in the current URL
    current_url = driver.current_url
    schedule_url = re.sub(r"/Home$", "/Schedule", current_url)
    driver.get(schedule_url)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_Create")))

def create_and_edit_schedule(for_date):
    # Click on "New Schedule"
    driver.find_element(By.ID, "MainContent_Create").click()
    
    # Wait for the "Save New Form" button to appear (indicating the schedule form is loaded)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_SaveNewForm")))

    # Save the new schedule
    driver.find_element(By.ID, "MainContent_SaveNewForm").click()
    
    # Wait for the edit icon of the first schedule to become clickable (indicating that the list is reloaded)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Edit this entry']")))
    
    # Edit the first schedule
    driver.find_element(By.XPATH, "//img[@alt='Edit this entry']").click()

    # Fill in schedule details (dynamically update based on for_date)
    schedule_from = driver.find_element(By.ID, "MainContent_NewScheduleDate")
    from_time = f"{for_date.strftime('%d-%b-%Y')} 09:30"
    driver.execute_script(f"arguments[0].setAttribute('value','{from_time}')", schedule_from)
    
    # Instead of using staleness_of, wait for the element to be interactable again after postback
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_NewScheduleDate")))

    # Once the page has reloaded, locate the new "schedule_to" field
    schedule_to = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_NewScheduleTo")))
    to_time = f"{for_date.strftime('%d-%b-%Y')} 17:30"
    driver.execute_script(f"arguments[0].setAttribute('value','{to_time}')", schedule_to)

    # Set the schedule type to "In Office"
    schedule_type_dropdown = Select(driver.find_element(By.ID, "MainContent_cmbScheduleType"))
    schedule_type_dropdown.select_by_visible_text("In Office")

    # Set the status to "Closed"
    current_date = datetime.now().date()
    if for_date.date() <= current_date:
        # Set the status to "Closed" only if the date is today or in the past
        status_dropdown = Select(driver.find_element(By.ID, "MainContent_cmbScheduleStatus"))
        status_dropdown.select_by_visible_text("Closed")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_cmbScheduleStatus")))
    else:
        print(f"Skipping status change for future date: {for_date.strftime('%d-%b-%Y')}")
    
    # Wait for the save button to be clickable again before clicking it
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_SaveNewForm"))).click()


def main():
    credentials = load_credentials()

    # Get current month and year
    current_date = datetime.now()
    month_days = calendar.monthrange(current_date.year, current_date.month)[1]  # Total days in the current month

    # Log in once
    login(credentials)
    navigate_to_schedule()

    # Loop through each day of the current month
    for day in range(1, month_days + 1):
        # Create a datetime object for the current day
        for_date = datetime(current_date.year, current_date.month, day)

        # Check if the current day needs attendance
        if is_attendance_needed(day, credentials["what_days_to_avoid_attendance"]):
            print(f"Marking attendance for {for_date.strftime('%d-%b-%Y')}")
            create_and_edit_schedule(for_date)
        else:
            print(f"Skipping attendance for {for_date.strftime('%d-%b-%Y')} (Day to avoid)")
    
    # Quit the driver after all attendance is completed
    driver.quit()

if __name__ == "__main__":
    main()