# selenium 4
import shutil
import time
from selenium.webdriver.common.by import By
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.options import Options
import datetime
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os
from datetime import datetime
from zoneinfo import ZoneInfo
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.edge.options import Options


options = Options()
options.use_chromium = True
service = EdgeService(executable_path=r"C:\Users\kiran.bs\Downloads\edgedriver_win64\msedgedriver.exe")  # Or full path to msedgedriver

# ---- Download folder & file setup ----
download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
counter = 1
renamed_file = f"Demographics Report{counter}.xlsx"

driver = webdriver.Edge(service=service, options=options)
driver.get("https://admin.benefitsjunction.com/")   # Replace with your login URL

driver.maximize_window()

driver.implicitly_wait(3)
USN = ""
PAS = ""
driver.maximize_window()
driver.find_element(By.XPATH, "//*[@id='TxtUserName']").send_keys(USN) #username 
driver.find_element(By.XPATH, "//*[@id='TxtPassword']").send_keys(PAS) 
time.sleep(1)
driver.find_element(By.XPATH, "//*[@id='LoginButton']").click()
time.sleep(1)

driver.switch_to.frame("FrmTabs")
print("Switched to frame: FrmTabs")

dropdown_element = driver.find_element(By.ID, "EmployeeUniversalSearchUC_ddlClients")
Select(dropdown_element).select_by_visible_text("Metal Pros LLC")   # Replace Company Name here
#print("W.T.Rich")   # Print the selected value
driver.find_element(By.ID, "EmployeeUniversalSearchUC_btnGo").click()
print("Clicked Go button")
time.sleep(1)  

#navigate to report tab
driver.switch_to.default_content()
driver.switch_to.frame("FrmClientAdminTabs")
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "HLReports"))).click()
print("Navigated to Reports tab")


#navigate to custom report
driver.switch_to.default_content()
driver.switch_to.frame("FrmClientAdminLeft")
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='TrvwPayrollt47' and text()='Custom Reports']"))
).click()
print("Navigated to Custom Report")


driver.switch_to.default_content()
driver.switch_to.frame("FrmClientAdminMiddle")
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_btnAdd"))).click()
print("Clicked Add button for new report")


WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_TxtRptName"))).send_keys("WFJ Metal Pros LLC Demographics Report")


WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ImgEmpPerInfo"))).click()
print("Opened Employee Personal Info section")


# Select Check Boxes
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_1"))).click()
print("Checked First Name")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_2"))).click()
print("Checked Last Name")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_3"))).click()
print("Checked Middle Name")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_4"))).click()
print("Checked Employee SSN")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_5"))).click()
print("Checked Date of Birth")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_7"))).click()
print("Checked Gender")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_8"))).click()
print("Checked Marital Status")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_10"))).click()
print("Checked Home Addr1")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_11"))).click()
print("Checked Home Addr2")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_12"))).click()
print("Checked Home City")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_13"))).click()
print("Checked Home State")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_14"))).click()
print("Checked Home ZIP")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_17"))).click()
print("Checked Home Phone")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_18"))).click()
print("Checked Cell Phone")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_19"))).click()
print("Checked Personal Email")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpPer_24"))).click()
print("Checked User")


WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ImgEmpWorkInfo"))).click()
print("Opened Employee Work Info section")
time.sleep(1)

# Work Info Checkboxes
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_2"))).click()
print("Checked Hire Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_3"))).click()
print("Checked ReHire Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_4"))).click()
print("Checked Status")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_5"))).click()
print("Checked Status Effective Date")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_8"))).click()

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_9"))).click()
print("Checked Div Effective Date")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_10"))).click()
print("Checked Location")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_11"))).click()
print("Checked Location Effective Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_12"))).click()
print("Checked Class")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_13"))).click()
print("Checked Class Effective Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_14"))).click()
print("Checked Department")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_15"))).click()
print("Checked Department Effective Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_18"))).click()
print("Checked Job Title")
time.sleep(1)
WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_19"))).click()
print("Checked Job Title Effective Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_20"))).click()
print("Checked Salary")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_21"))).click()
print("Checked Salary Period")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_22"))).click()
print("Checked Job Effective Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_23"))).click()
print("Checked Pay Cycle")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_24"))).click()
print("Checked Pay Cycle Effective Date")

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "CBLEmpWork_37"))).click()
print("Checked Username")


WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_btnNext"))).click()
print("Clicked Next Button")
time.sleep(1)

dropdown = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "lbxFields"))
)
print("Located dropdown: lbxFields")

# Define custom order
custom_order = [
    "First Name", "Middle Name", "Last Name", "Employee SSN", "Personal Email",
    "Date Of Birth", "Gender", "Home Addr1", "Home Addr2", "Home Addr State",
    "Home Addr City", "Home Addr Zip", "Home Phone", "Cell Phone", "Marital Status",
    "Job Title", "Job Title Effective Date", "Status", "Status Eff Date", "Hire Date",
    "Rehire Date", "Work Email", "Location", "Location Effective Date", "Division",
    "Div Effective Date", "Department", "Dept Effective Date", "Class", "Class Effective Date",
    "Pay Cycle", "Pay Cycle Effective Date", "Salary", "Salary Period", "Salary Effective Date",
    "Username"
]
print("Custom order defined.")

# Get all current options
all_options = {opt.text: opt.get_attribute("value") for opt in dropdown.find_elements(By.TAG_NAME, "option")}
print("Captured all dropdown options:", list(all_options.keys()))

# Filter options
options_to_use = [field for field in custom_order if field in all_options]
print("Filtered options in custom order:", options_to_use)

# Replace dropdown options
driver.execute_script("""
    var select = arguments[0];
    var values = arguments[1];
    var valuesMap = arguments[2];
    select.innerHTML = "";
    for (var i=0; i<values.length; i++) {
        var opt = document.createElement("option");
        opt.text = values[i];
        opt.value = valuesMap[values[i]];
        select.appendChild(opt);
    }
""", dropdown, options_to_use, all_options)
print("Dropdown reordered successfully.")

driver.execute_script("arguments[0].selectedIndex = 1;", dropdown)
print("First option selected in dropdown.")
time.sleep(1)

WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "btnSubmit"))).click()
print("Finalize Button Clicked")
time.sleep(1)


#navigate to generate option
driver.switch_to.default_content()
driver.switch_to.frame("FrmClientAdminLeft")
WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='TrvwPayrollt47' and text()='Custom Reports']"))
).click()
print("Navigated to Custom Report")
time.sleep(1)

driver.switch_to.default_content()
driver.switch_to.frame("FrmClientAdminMiddle")

report_name = "WFJ Metal Pros LLC Demographics Report"
print("Looking for report:", report_name)

try:
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.ID, "ctl00_ctl00_body_body_GVPrevCustRpts"))
    )
    xpath = f"//table[@id='ctl00_ctl00_body_body_GVPrevCustRpts']//tr[td[normalize-space(.)='{report_name}']]//a[contains(text(),'Generate')]"
    generate_link = WebDriverWait(driver, 10).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    driver.execute_script("arguments[0].scrollIntoView(true);", generate_link)
    time.sleep(0.5)
    generate_link.click()
    print(f" Generate clicked for: {report_name}")
    time.sleep(2)

    # ---- Set the date ----
    driver.switch_to.default_content()
    driver.switch_to.frame("FrmClientAdminMiddle")
    date_input = driver.find_element(By.ID, "ctl00_ctl00_body_body_TxtEffDate")
    driver.execute_script("""
        arguments[0].value = '10/1/2025';
        arguments[0].dispatchEvent(new Event('change'));
    """, date_input)
    time.sleep(1)

    dropdown_input = driver.find_element(By.XPATH, "//input[@type='text' and @value='Select']")
    dropdown_input.click()
    time.sleep(1)
    driver.find_element(By.XPATH, "//li[text()='All']").click()
    time.sleep(1)

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_chkIncTerm"))).click()
    print("Checked Include Terminations")
    time.sleep(1)

    dropdown_element = driver.find_element(By.ID, "ctl00_ctl00_body_body_ddlEmpType")
    Select(dropdown_element).select_by_visible_text("All")
    print("EE type ALL selected")

    WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_btnGenerate"))).click()
    print("Clicked generate button")
    

    # ---- Navigate to previously Generated tab ----
    driver.switch_to.default_content()
    driver.switch_to.frame("FrmClientAdminLeft")
    WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.XPATH, "//*[@id='TrvwPayrollt50' and text()='Previously Generated']"))).click()
    print("Clicked previously Generated button")
    

except Exception as e: 
    print("Error:", str(e))


# ---- Check status & download ----
driver.switch_to.default_content()
driver.switch_to.frame("FrmClientAdminMiddle")

refresh_btn_id = "ctl00_ctl00_body_body_btnRefresh"
report_link_xpath = "//table[@id='ctl00_ctl00_body_body_GVPrevCustGene']/tbody/tr[2]/td[2]/a[contains(text(),'WFJ')]"
status_xpath = "//table[@id='ctl00_ctl00_body_body_GVPrevCustGene']/tbody/tr[2]/td[5]"

max_retries = 20
retry_count = 0
while retry_count < max_retries:
    try:
        status_td = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, status_xpath)))
        status_text = status_td.text.strip()
        print(f"Attempt {retry_count+1}: Status = {status_text}")

        if status_text.lower() == "completed":
            report_link = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, report_link_xpath)))
            driver.execute_script("arguments[0].scrollIntoView(true);", report_link)
            time.sleep(0.5)
            report_link.click()
            print(" Report link clicked!")

            # --- Wait for download and rename ---
            filename_partial = "Demographics Report"
            download_wait = 0
            max_wait = 120
            while download_wait < max_wait:
                files = [f for f in os.listdir(download_dir)
                         if filename_partial in f and not f.endswith(".crdownload")]
                if files:
                    original_file = os.path.join(download_dir, files[0])
                    renamed_path = os.path.join(download_dir, renamed_file)
                    if os.path.exists(renamed_path):
                        os.remove(renamed_path)
                    shutil.move(original_file, renamed_path)
                    print(f" File downloaded and renamed to: {renamed_file}")
                    break
                time.sleep(2)
                download_wait += 2
            else:
                print(" Download did not complete within 2 minutes.")
            break
        else:
            refresh_btn = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, refresh_btn_id)))
            refresh_btn.click()
            print("⏳ Status pending, clicked refresh. Waiting 30s...")
            time.sleep(30)
            retry_count += 1

    except Exception as e:
        print(f"❌ Error: {e}, retrying in 5s...")
        time.sleep(5)
        retry_count += 1
    




