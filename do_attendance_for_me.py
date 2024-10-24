import re
from selenium import webdriver  # type: ignore
from selenium.webdriver.common.by import By  # type: ignore
from selenium.webdriver.chrome.service import Service  # type: ignore
from selenium.webdriver.support.ui import Select  # type: ignore
from selenium.webdriver.support.ui import WebDriverWait  # type: ignore
from selenium.webdriver.support import expected_conditions as EC  # type: ignore
from datetime import datetime
import calendar

def load_credentials():
    with open('needy.txt', 'r') as file:
        data = file.readlines()
    credentials = {}
    for line in data:
        key, value = line.strip().split("=")
        credentials[key] = value.strip('"')
    return credentials

def is_attendance_needed(for_date, days_to_mark):
    mark_days = [datetime.strptime(day.strip(), "%d-%b") for day in days_to_mark.split(',')]
    
    mark_days = [day.replace(year=for_date.year) for day in mark_days]
    
    return for_date in mark_days

service = Service('C:/Users/leywi/OneDrive/Documents/development/chromedriver-win64/chromedriver.exe')  # Change this path to your chromedriver path
driver = webdriver.Chrome(service=service)

def login(credentials):
    driver.get('https://www.resoluteonline.in/Login')

    driver.find_element(By.ID, "MainContent_UserName").send_keys(credentials['website_username'])
    driver.find_element(By.ID, "MainContent_Password").send_keys(credentials['website_password'])
    
    driver.find_element(By.ID, "MainContent_LoginButton").click()

    print("Waiting for OTP input...")
    WebDriverWait(driver, 300).until(EC.url_contains("/Home"))

def navigate_to_schedule():
    current_url = driver.current_url
    schedule_url = re.sub(r"/Home$", "/Schedule", current_url)
    driver.get(schedule_url)
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_Create")))

def create_and_edit_schedule(for_date):
    driver.find_element(By.ID, "MainContent_Create").click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_SaveNewForm")))

    driver.find_element(By.ID, "MainContent_SaveNewForm").click()
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, "//img[@alt='Edit this entry']")))
    
    driver.find_element(By.XPATH, "//img[@alt='Edit this entry']").click()

    schedule_from = driver.find_element(By.ID, "MainContent_NewScheduleDate")
    from_time = f"{for_date.strftime('%d-%b-%Y')} 09:30"
    driver.execute_script(f"arguments[0].setAttribute('value','{from_time}')", schedule_from)
    
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_NewScheduleDate")))

    schedule_to = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_NewScheduleTo")))
    to_time = f"{for_date.strftime('%d-%b-%Y')} 17:30"
    driver.execute_script(f"arguments[0].setAttribute('value','{to_time}')", schedule_to)

    schedule_type_dropdown = Select(driver.find_element(By.ID, "MainContent_cmbScheduleType"))
    schedule_type_dropdown.select_by_visible_text("In Office")

    current_date = datetime.now().date()
    if for_date.date() <= current_date:
        status_dropdown = Select(driver.find_element(By.ID, "MainContent_cmbScheduleStatus"))
        status_dropdown.select_by_visible_text("Closed")
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "MainContent_cmbScheduleStatus")))
    else:
        print(f"Skipping status change for future date: {for_date.strftime('%d-%b-%Y')}")
    
    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "MainContent_SaveNewForm"))).click()


def main():
    credentials = load_credentials()

    month_to_mark = credentials['month_to_mark']
    year_to_mark = int(credentials['year_to_mark'])


    month_number = datetime.strptime(month_to_mark, "%b").month
    month_days = calendar.monthrange(year_to_mark, month_number)[1]

    login(credentials)
    navigate_to_schedule()

    for day in range(1, month_days + 1):
        for_date = datetime(year_to_mark, month_number, day)
        if is_attendance_needed(for_date, credentials["what_days_to_mark_attendance"]):
            print(f"Marking attendance for {for_date.strftime('%d-%b-%Y')}")
            create_and_edit_schedule(for_date)
        else:
            print(f"Skipping attendance for {for_date.strftime('%d-%b-%Y')}")
    
    driver.quit()

if __name__ == "__main__":
    main()
