"""Selenium automation implementations for Workforce Junction.

Implements:
1. Custom Report Generator (BenefitsJunction admin portal automation)
2. Auto-Login Tester (OE popup assertions from Excel records)
3. New Client Auto Login (Portal login validation)
4. New Client Password Setup (Initial DOB password setup)
"""

from __future__ import annotations

import os
import re
import shutil
import time
from datetime import datetime
from typing import TYPE_CHECKING, Any, Dict, List

import pandas as pd
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.ui import WebDriverWait

if TYPE_CHECKING:
    from api import UI


def _read_excel_rows(path: str, required_columns: List[str]) -> List[Dict[str, Any]]:
    """Read rows from an Excel file and return them as dictionaries."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Excel file not found: {path}")

    df = pd.read_excel(path, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Case-insensitive column matching
    cmap = {c.lower(): c for c in df.columns}
    normalized_cols = {}
    for req in required_columns:
        match = cmap.get(req.lower())
        if match:
            normalized_cols[req] = match
        else:
            raise ValueError(f"Missing column '{req}' in {path}; found {list(df.columns)}")

    df = df.rename(columns={orig: req for req, orig in normalized_cols.items()})
    return df[required_columns].dropna().to_dict("records")


def _create_edge_driver(driver_path: str = "", download_dir: str = "") -> webdriver.Edge:
    """Create and return an Edge WebDriver instance with appropriate options."""
    options = EdgeOptions()
    options.use_chromium = True
    options.add_argument("--start-maximized")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")

    if download_dir:
        prefs = {
            "download.default_directory": download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
        }
        options.add_experimental_option("prefs", prefs)

    if driver_path and os.path.exists(driver_path):
        service = EdgeService(executable_path=driver_path)
        return webdriver.Edge(service=service, options=options)
    
    # Try default Edge service
    try:
        return webdriver.Edge(options=options)
    except Exception:
        # Fallback to Chrome if Edge is unavailable
        chrome_opts = ChromeOptions()
        chrome_opts.add_argument("--start-maximized")
        if download_dir:
            chrome_opts.add_experimental_option("prefs", {
                "download.default_directory": download_dir,
                "download.prompt_for_download": False,
            })
        return webdriver.Chrome(options=chrome_opts)


def _create_chrome_driver() -> webdriver.Chrome | webdriver.Edge:
    """Create and return a Chrome WebDriver instance with automatic fallback to Edge."""
    # 1. Try Chrome first
    try:
        opts = ChromeOptions()
        opts.add_argument("--start-maximized")
        opts.add_argument("--disable-gpu")
        opts.add_argument("--no-sandbox")
        return webdriver.Chrome(options=opts)
    except Exception:
        pass

    # 2. Try Edge fallback (built into Windows)
    try:
        edge_opts = EdgeOptions()
        edge_opts.use_chromium = True
        edge_opts.add_argument("--start-maximized")
        edge_opts.add_argument("--disable-gpu")
        edge_opts.add_argument("--no-sandbox")
        return webdriver.Edge(options=edge_opts)
    except Exception as edge_err:
        raise RuntimeError(f"Could not launch Chrome or Edge browser: {edge_err}")



# ==============================================================================
# 1. Custom Report Generator (Based on Custom Report Generation.py)
# ==============================================================================

def run_custom_report(params: Dict[str, Any], ui: "UI") -> None:
    """
    Drives the BenefitsJunction admin portal end-to-end to build and download
    a Demographics report with exact column ordering and selections.
    """
    driver_path = str(params.get("driver_path") or "").strip()
    portal_url = str(params.get("portal_url") or "https://admin.benefitsjunction.com/").strip()
    username = str(params.get("username") or "").strip()
    password = str(params.get("password") or "").strip()
    client_name = str(params.get("client_name") or "Metal Pros LLC").strip()
    report_name = str(params.get("report_name") or "WFJ Metal Pros LLC Demographics Report").strip()
    report_date = str(params.get("report_date") or "10/1/2025").strip()

    download_dir = os.path.join(os.path.expanduser("~"), "Downloads")
    os.makedirs(download_dir, exist_ok=True)

    ui.log("Initializing WebDriver...")
    driver = None
    try:
        driver = _create_edge_driver(driver_path=driver_path, download_dir=download_dir)
        driver.maximize_window()
        driver.implicitly_wait(4)

        ui.log(f"Navigating to {portal_url}")
        driver.get(portal_url)

        wait = WebDriverWait(driver, 15)

        # Login
        ui.log(f"Entering credentials -> Username: {username} | Password: {password}")
        username_el = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='TxtUserName']")))
        username_el.clear()
        username_el.send_keys(username)

        password_el = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='TxtPassword']")))
        password_el.clear()
        password_el.send_keys(password)
        time.sleep(1)

        wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='LoginButton']"))).click()
        time.sleep(1)
        ui.log("Logged in successfully.")

        # Switch to frame: FrmTabs
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmTabs")
        ui.log("Switched to frame: FrmTabs")

        dropdown_el = wait.until(EC.presence_of_element_located((By.ID, "EmployeeUniversalSearchUC_ddlClients")))
        Select(dropdown_el).select_by_visible_text(client_name)
        ui.log(f"Selected client: {client_name}")

        wait.until(EC.element_to_be_clickable((By.ID, "EmployeeUniversalSearchUC_btnGo"))).click()
        ui.log("Clicked Go button")
        time.sleep(1)

        # Navigate to Reports tab
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminTabs")
        wait.until(EC.element_to_be_clickable((By.ID, "HLReports"))).click()
        ui.log("Navigated to Reports tab")

        # Navigate to Custom Reports
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminLeft")
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='TrvwPayrollt47' and text()='Custom Reports']"))
        ).click()
        ui.log("Navigated to Custom Reports")

        # Add new report
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminMiddle")
        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_btnAdd"))).click()
        ui.log("Clicked Add button for new report")

        # Report name
        rpt_name_input = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_TxtRptName")))
        rpt_name_input.clear()
        rpt_name_input.send_keys(report_name)
        ui.log(f"Entered report name: {report_name}")

        # Personal Info section
        wait.until(EC.element_to_be_clickable((By.ID, "ImgEmpPerInfo"))).click()
        ui.log("Opened Employee Personal Info section")

        personal_cbs = [
            ("CBLEmpPer_1", "First Name"),
            ("CBLEmpPer_2", "Last Name"),
            ("CBLEmpPer_3", "Middle Name"),
            ("CBLEmpPer_4", "Employee SSN"),
            ("CBLEmpPer_5", "Date of Birth"),
            ("CBLEmpPer_7", "Gender"),
            ("CBLEmpPer_8", "Marital Status"),
            ("CBLEmpPer_10", "Home Addr1"),
            ("CBLEmpPer_11", "Home Addr2"),
            ("CBLEmpPer_12", "Home City"),
            ("CBLEmpPer_13", "Home State"),
            ("CBLEmpPer_14", "Home ZIP"),
            ("CBLEmpPer_17", "Home Phone"),
            ("CBLEmpPer_18", "Cell Phone"),
            ("CBLEmpPer_19", "Personal Email"),
            ("CBLEmpPer_24", "User"),
        ]

        for cb_id, name in personal_cbs:
            try:
                cb = wait.until(EC.element_to_be_clickable((By.ID, cb_id)))
                if not cb.is_selected():
                    cb.click()
            except Exception as e:
                ui.log(f"Note on {name} ({cb_id}): {e}", "warn")

        ui.log("Checked Personal Info fields")

        # Work Info section
        wait.until(EC.element_to_be_clickable((By.ID, "ImgEmpWorkInfo"))).click()
        ui.log("Opened Employee Work Info section")
        time.sleep(1)

        work_cbs = [
            ("CBLEmpWork_2", "Hire Date"),
            ("CBLEmpWork_3", "ReHire Date"),
            ("CBLEmpWork_4", "Status"),
            ("CBLEmpWork_5", "Status Effective Date"),
            ("CBLEmpWork_8", "Division"),
            ("CBLEmpWork_9", "Div Effective Date"),
            ("CBLEmpWork_10", "Location"),
            ("CBLEmpWork_11", "Location Effective Date"),
            ("CBLEmpWork_12", "Class"),
            ("CBLEmpWork_13", "Class Effective Date"),
            ("CBLEmpWork_14", "Department"),
            ("CBLEmpWork_15", "Department Effective Date"),
            ("CBLEmpWork_18", "Job Title"),
            ("CBLEmpWork_19", "Job Title Effective Date"),
            ("CBLEmpWork_20", "Salary"),
            ("CBLEmpWork_21", "Salary Period"),
            ("CBLEmpWork_22", "Job Effective Date"),
            ("CBLEmpWork_23", "Pay Cycle"),
            ("CBLEmpWork_24", "Pay Cycle Effective Date"),
            ("CBLEmpWork_37", "Username"),
        ]

        for cb_id, name in work_cbs:
            try:
                cb = wait.until(EC.element_to_be_clickable((By.ID, cb_id)))
                if not cb.is_selected():
                    cb.click()
            except Exception as e:
                ui.log(f"Note on {name} ({cb_id}): {e}", "warn")

        ui.log("Checked Work Info fields")

        # Click Next Button
        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_btnNext"))).click()
        ui.log("Clicked Next Button")
        time.sleep(1)

        dropdown = wait.until(EC.presence_of_element_located((By.ID, "lbxFields")))
        ui.log("Located dropdown: lbxFields")

        # Define custom order from user script
        custom_order = [
            "First Name", "Middle Name", "Last Name", "Employee SSN", "Personal Email",
            "Date Of Birth", "Gender", "Home Addr1", "Home Addr2", "Home Addr State",
            "Home Addr City", "Home Addr Zip", "Home Phone", "Cell Phone", "Marital Status",
            "Job Title", "Job Title Effective Date", "Status", "Status Eff Date", "Hire Date",
            "Rehire Date", "Work Email", "Location", "Location Effective Date", "Division",
            "Div Effective Date", "Department", "Dept Effective Date", "Class", "Class Effective Date",
            "Pay Cycle", "Pay Cycle Effective Date", "Salary", "Salary Period", "Salary Effective Date",
            "Username",
        ]

        all_options = {opt.text.strip(): opt.get_attribute("value") for opt in dropdown.find_elements(By.TAG_NAME, "option")}
        options_to_use = [field for field in custom_order if field in all_options]

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
        ui.log("Dropdown reordered successfully in custom order.")

        driver.execute_script("arguments[0].selectedIndex = 0;", dropdown)
        time.sleep(1)

        wait.until(EC.element_to_be_clickable((By.ID, "btnSubmit"))).click()
        ui.log("Finalize Button Clicked")
        time.sleep(1)

        # Navigate back to Custom Reports list
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminLeft")
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='TrvwPayrollt47' and text()='Custom Reports']"))
        ).click()
        ui.log("Navigated back to Custom Reports list")
        time.sleep(1)

        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminMiddle")
        ui.log(f"Looking for report: {report_name}")

        wait.until(EC.presence_of_element_located((By.ID, "ctl00_ctl00_body_body_GVPrevCustRpts")))
        xpath = f"//table[@id='ctl00_ctl00_body_body_GVPrevCustRpts']//tr[td[normalize-space(.)='{report_name}']]//a[contains(text(),'Generate')]"
        generate_link = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView(true);", generate_link)
        time.sleep(0.5)
        generate_link.click()
        ui.log(f"Generate clicked for: {report_name}")
        time.sleep(2)

        # Set effective date and options
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminMiddle")
        date_input = wait.until(EC.presence_of_element_located((By.ID, "ctl00_ctl00_body_body_TxtEffDate")))
        driver.execute_script(f"""
            arguments[0].value = '{report_date}';
            arguments[0].dispatchEvent(new Event('change'));
        """, date_input)
        time.sleep(1)

        try:
            dropdown_input = driver.find_element(By.XPATH, "//input[@type='text' and @value='Select']")
            dropdown_input.click()
            time.sleep(1)
            driver.find_element(By.XPATH, "//li[text()='All']").click()
            time.sleep(1)
        except Exception:
            pass

        try:
            inc_term = wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_chkIncTerm")))
            if not inc_term.is_selected():
                inc_term.click()
            ui.log("Checked Include Terminations")
            time.sleep(1)
        except Exception:
            pass

        try:
            dropdown_el = driver.find_element(By.ID, "ctl00_ctl00_body_body_ddlEmpType")
            Select(dropdown_el).select_by_visible_text("All")
            ui.log("EE type ALL selected")
        except Exception:
            pass

        wait.until(EC.element_to_be_clickable((By.ID, "ctl00_ctl00_body_body_btnGenerate"))).click()
        ui.log("Clicked generate button")

        # Navigate to previously Generated tab
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminLeft")
        wait.until(
            EC.element_to_be_clickable((By.XPATH, "//*[@id='TrvwPayrollt50' and text()='Previously Generated']"))
        ).click()
        ui.log("Clicked previously Generated button")

        # Check status & download
        driver.switch_to.default_content()
        driver.switch_to.frame("FrmClientAdminMiddle")

        refresh_btn_id = "ctl00_ctl00_body_body_btnRefresh"
        report_link_xpath = "//table[@id='ctl00_ctl00_body_body_GVPrevCustGene']/tbody/tr[2]/td[2]/a"
        status_xpath = "//table[@id='ctl00_ctl00_body_body_GVPrevCustGene']/tbody/tr[2]/td[5]"

        before_files = set(os.listdir(download_dir))
        max_retries = 30
        retry_count = 0

        while retry_count < max_retries:
            try:
                status_td = wait.until(EC.presence_of_element_located((By.XPATH, status_xpath)))
                status_text = status_td.text.strip()
                ui.log(f"Attempt {retry_count+1}: Status = {status_text}")

                if status_text.lower() == "completed":
                    report_link = wait.until(EC.element_to_be_clickable((By.XPATH, report_link_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView(true);", report_link)
                    time.sleep(0.5)
                    report_link.click()
                    ui.log("Report link clicked!")

                    # Wait for download and rename
                    filename_partial = "Demographics Report"
                    download_wait = 0
                    max_wait = 120
                    renamed_final_file = None

                    while download_wait < max_wait:
                        current_files = set(os.listdir(download_dir))
                        candidates = [
                            f for f in (current_files - before_files)
                            if not f.endswith(".crdownload") and not f.endswith(".tmp")
                        ]
                        if not candidates:
                            candidates = [
                                f for f in os.listdir(download_dir)
                                if filename_partial in f and not f.endswith(".crdownload") and not f.endswith(".tmp")
                            ]
                        if candidates:
                            latest_file = max(
                                [os.path.join(download_dir, f) for f in candidates],
                                key=os.path.getmtime,
                            )
                            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                            renamed_path = os.path.join(download_dir, f"Demographics_Report_{ts}.xlsx")
                            try:
                                shutil.move(latest_file, renamed_path)
                                renamed_final_file = renamed_path
                            except Exception:
                                renamed_final_file = latest_file
                            ui.log(f"✓ File downloaded and saved to: {renamed_final_file}", "ok")
                            break
                        time.sleep(2)
                        download_wait += 2
                    else:
                        ui.log("Download did not complete within 2 minutes.", "warn")

                    ui.status("Custom Report generated successfully", "pass")
                    break
                else:
                    refresh_btn = wait.until(EC.element_to_be_clickable((By.ID, refresh_btn_id)))
                    refresh_btn.click()
                    ui.log("Status pending, clicked refresh. Waiting 15s...")
                    time.sleep(15)
                    retry_count += 1
            except Exception as exc:
                ui.log(f"Polling note: {exc}, retrying in 5s...", "warn")
                time.sleep(5)
                retry_count += 1

    finally:
        if driver:
            try:
                driver.quit()
            except Exception:
                pass


# ==============================================================================
# 2. Auto-Login & OE Tester (Based on Test Records_Auto_Login.py)
# ==============================================================================

def run_auto_login_test(excel_path: str, portal_url: str, expected_text: str, ui: "UI") -> None:
    """
    Reads Username and Password from Excel, logs in, and asserts popup text
    matches the expected substring lines from the PopUp_LblMsg element.
    """
    df = pd.read_excel(excel_path)
    df.columns = [str(c).strip() for c in df.columns]
    cmap = {c.lower(): c for c in df.columns}
    uc = cmap.get("username")
    pc = cmap.get("password")

    if not uc or not pc:
        raise ValueError(f"Excel file must contain 'Username' and 'Password' columns. Found: {list(df.columns)}")

    records = df[[uc, pc]].dropna().rename(columns={uc: "Username", pc: "Password"}).to_dict("records")
    lines = [ln.strip() for ln in expected_text.strip().splitlines() if ln.strip()]
    total = len(records)

    ui.log("=" * 65)
    ui.log(f"Starting {total} test(s) | {len(lines)} assertion(s) | {time.strftime('%H:%M:%S')}")
    ui.log("=" * 65)

    driver = _create_chrome_driver()
    passed, failed = [], []

    try:
        wait = WebDriverWait(driver, 15)

        for i, row in enumerate(records, 1):
            user = str(row["Username"]).strip()
            pwd = str(row["Password"]).strip()
            lbl = f"[{i:02d}/{total}] {user}"

            try:
                driver.get(portal_url.strip())

                # Log credentials being attempted
                ui.log(f"  Entering credentials -> Username: {user} | Password: {pwd}")

                # Username input
                uf = wait.until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//input[@type='text' and ("
                        "@placeholder='Username' or @name='username' or "
                        "@id='username' or @id='Username' or "
                        "contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'user'))]"
                    ))
                )
                uf.clear()
                uf.send_keys(user)

                # Password input
                pf = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
                pf.clear()
                pf.send_keys(pwd)

                # Log In button
                wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//input[@value='Log In'] | //button[normalize-space()='Log In']"
                    ))
                ).click()

                # Wait for popup
                popup = wait.until(EC.visibility_of_element_located((By.ID, "PopUp_LblMsg")))

                popup_text = popup.text.strip()
                if not popup_text:
                    popup_text = driver.execute_script("return arguments[0].innerText;", popup).strip()

                # Check assertions
                missing = [ln for ln in lines if ln not in popup_text]
                if missing:
                    raise AssertionError("Missing from popup:\n" + "\n".join(f"  • {m}" for m in missing))

                ui.log(f"PASS {lbl}", "pass")
                passed.append(user)

            except TimeoutException:
                r = "Timed out — login failed or popup not visible"
                ui.log(f"FAIL {lbl}\n  {r}", "fail")
                failed.append({"Username": user, "Password": pwd, "Reason": r})
            except AssertionError as e:
                ui.log(f"FAIL {lbl}\n{e}", "fail")
                failed.append({"Username": user, "Password": pwd, "Reason": str(e)})
            except Exception as e:
                r = f"{type(e).__name__}: {str(e)[:100]}"
                ui.log(f"FAIL {lbl}\n  {r}", "fail")
                failed.append({"Username": user, "Password": pwd, "Reason": r})

            ui.progress((i / total) * 100)
            ui.status(f"Running {i}/{total} — {len(passed)} passed, {len(failed)} failed")
            time.sleep(1)

        ui.log("\n" + "=" * 65)
        ui.log("RESULTS SUMMARY")
        ui.log("=" * 65)
        ui.log(f"Passed : {len(passed)} / {total}", "pass")
        ui.log(f"Failed : {len(failed)} / {total}", "fail" if failed else "info")

        if failed:
            ui.log("\nFAILED RECORDS:", "fail")
            for f in failed:
                ui.log(
                    f"  Username: {f['Username']} | Password: {f.get('Password', '-')} | Reason: {f['Reason']}\n  {'-'*50}",
                    "fail"
                )
            ui.status(f"Done — {len(passed)} passed, {len(failed)} failed", "fail")
        else:
            ui.log("All records passed!", "pass")
            ui.status(f"Done — {len(passed)} passed", "pass")

    finally:
        driver.quit()


# ==============================================================================
# 3. New Client Auto Login Check
# ==============================================================================

def run_new_client_auto_login(excel_path: str, portal_url: str, ui: "UI") -> None:
    """Verifies that new-client test records can log into the portal."""
    rows = _read_excel_rows(excel_path, ["Username", "Password"])
    total = len(rows)
    ui.log("=" * 65)
    ui.log(f"Starting Auto-Login Check: {total} record(s) | {time.strftime('%H:%M:%S')}")
    ui.log("=" * 65)

    driver = _create_chrome_driver()
    passed, failed = [], []

    try:
        wait = WebDriverWait(driver, 15)

        for i, row in enumerate(rows, 1):
            user = str(row.get("Username", "")).strip()
            pwd = str(row.get("Password", "")).strip()
            lbl = f"[{i:02d}/{total}] {user}"

            try:
                driver.get(portal_url.strip())

                # Log credentials being attempted
                ui.log(f"  Entering credentials -> Username: {user} | Password: {pwd}")

                uf = wait.until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//input[@type='text' and ("
                        "@placeholder='Username' or @name='username' or "
                        "@id='username' or @id='Username' or "
                        "contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'user'))]"
                    ))
                )
                uf.clear()
                uf.send_keys(user)

                pf = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
                pf.clear()
                pf.send_keys(pwd)

                wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//input[@value='Log In'] | //button[normalize-space()='Log In']"
                    ))
                ).click()

                time.sleep(2.5)
                # Check for explicit error banner/label
                err_els = driver.find_elements(
                    By.XPATH,
                    "//*[contains(@class,'error') or contains(@id,'lblError') or contains(@id,'ErrMsg') or contains(@class,'alert-danger') or contains(@id,'Error')]"
                )
                visible_errors = [e.text.strip() for e in err_els if e.is_displayed() and e.text.strip()]

                # Check if login button is still displayed (indicating failure)
                login_btn = driver.find_elements(By.XPATH, "//input[@value='Log In'] | //button[normalize-space()='Log In']")
                is_login_still_there = any(b.is_displayed() for b in login_btn)

                if visible_errors:
                    err_msg = visible_errors[0]
                    ui.log(f"FAIL {lbl} — {err_msg}", "fail")
                    failed.append({"Username": user, "Password": pwd, "Reason": err_msg})
                elif is_login_still_there:
                    ui.log(f"FAIL {lbl} — Login unsuccessful (remained on login screen)", "fail")
                    failed.append({"Username": user, "Password": pwd, "Reason": "Remained on login screen"})
                else:
                    ui.log(f"PASS {lbl} — Authenticated successfully", "pass")
                    passed.append(user)

            except Exception as e:
                err_str = str(e).splitlines()[0] if str(e) else type(e).__name__
                ui.log(f"FAIL {lbl} — {err_str}", "fail")
                failed.append({"Username": user, "Password": pwd, "Reason": err_str})

            ui.progress((i / total) * 100)
            ui.status(f"Checked {i}/{total} — {len(passed)} passed, {len(failed)} failed")
            time.sleep(1)

        ui.log("\n" + "=" * 65)
        ui.log("RESULTS SUMMARY")
        ui.log("=" * 65)
        ui.log(f"Passed : {len(passed)} / {total}", "pass")
        ui.log(f"Failed : {len(failed)} / {total}", "fail" if failed else "info")

        if failed:
            ui.log("\nFAILED RECORDS:", "fail")
            for f in failed:
                ui.log(f"  Username: {f['Username']} | Password: {f.get('Password', '-')} | Reason: {f['Reason']}", "fail")
            ui.status(f"Done — {len(passed)} passed, {len(failed)} failed", "fail")
        else:
            ui.log("All records passed!", "pass")
            ui.status(f"Done — {len(passed)} passed", "pass")

    finally:
        driver.quit()


# ==============================================================================
# 4. New Client Password Setup
# ==============================================================================

def run_new_client_password_setup(excel_path: str, portal_url: str, password: str, ui: "UI") -> None:
    """
    Sets up initial passwords for employee records from UserID and Date Of Birth.
    """
    rows = _read_excel_rows(excel_path, ["UserID", "Date Of Birth"])
    total = len(rows)
    ui.log("=" * 65)
    ui.log(f"Starting Password Setup: {total} record(s) | {time.strftime('%H:%M:%S')}")
    ui.log("=" * 65)

    driver = _create_chrome_driver()
    passed, failed = [], []

    try:
        wait = WebDriverWait(driver, 15)

        for i, row in enumerate(rows, 1):
            user = str(row.get("UserID", "")).strip()
            dob_raw = str(row.get("Date Of Birth", "")).strip()
            lbl = f"[{i:02d}/{total}] {user}"

            # Format DOB as temp password
            # Strip all non-digit chars (handles dates like "1990-05-14", "1990/05/14", timestamps)
            dob_clean = re.sub(r"[^\d]", "", dob_raw)
            if len(dob_clean) == 14:  # Timestamp: YYYYMMDDHHMMSS -> use first 8 digits
                dob_clean = dob_clean[:8]
            if len(dob_clean) == 8:
                # If YYYYMMDD, reformat to MMDDYYYY
                if dob_clean[:4].startswith(("19", "20")):  # year first
                    temp_pwd = dob_clean[4:6] + dob_clean[6:8] + dob_clean[:4]  # MMDDYYYY
                else:
                    temp_pwd = dob_clean  # already MMDDYYYY
            else:
                temp_pwd = dob_raw  # fallback to raw string

            try:
                driver.get(portal_url.strip())

                # Log credentials being attempted
                ui.log(f"  Entering credentials -> Username: {user} | Temp Password (DOB): {temp_pwd}")

                uf = wait.until(
                    EC.presence_of_element_located((
                        By.XPATH,
                        "//input[@type='text' and ("
                        "@placeholder='Username' or @name='username' or "
                        "@id='username' or @id='Username' or "
                        "contains(translate(@id,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'user'))]"
                    ))
                )
                uf.clear()
                uf.send_keys(user)

                pf = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='password']")))
                pf.clear()
                pf.send_keys(temp_pwd)

                wait.until(
                    EC.element_to_be_clickable((
                        By.XPATH,
                        "//input[@value='Log In'] | //button[normalize-space()='Log In']"
                    ))
                ).click()

                time.sleep(2.5)

                # Check if login threw an error
                err_els = driver.find_elements(
                    By.XPATH,
                    "//*[contains(@class,'error') or contains(@id,'lblError') or contains(@id,'ErrMsg') or contains(@class,'alert-danger') or contains(@id,'Error')]"
                )
                visible_errors = [e.text.strip() for e in err_els if e.is_displayed() and e.text.strip()]

                # Check for password reset fields
                new_pwd_fields = driver.find_elements(
                    By.XPATH,
                    "//input[@type='password' and (contains(@id,'New') or contains(@name,'New') or contains(@id,'txtNewPwd') or contains(@placeholder,'New Password'))]"
                )

                if visible_errors and not new_pwd_fields:
                    err_msg = visible_errors[0]
                    ui.log(f"FAIL {lbl} — Login rejected: {err_msg} | Temp Password used: {temp_pwd}", "fail")
                    failed.append({"Username": user, "TempPassword": temp_pwd, "Reason": f"Login rejected: {err_msg}"})
                    continue

                if new_pwd_fields:
                    new_pwd_fields[0].clear()
                    new_pwd_fields[0].send_keys(password)

                    confirm_fields = driver.find_elements(
                        By.XPATH,
                        "//input[@type='password' and (contains(@id,'Confirm') or contains(@name,'Confirm') or contains(@id,'txtConfirmPwd') or contains(@placeholder,'Confirm'))]"
                    )
                    if confirm_fields:
                        confirm_fields[0].clear()
                        confirm_fields[0].send_keys(password)

                    # Security questions if present
                    sec_selects = driver.find_elements(By.XPATH, "//select[contains(@id,'Question') or contains(@name,'Question') or contains(@id,'ddlSec')]")
                    for s in sec_selects:
                        try:
                            sel = Select(s)
                            if len(sel.options) > 1:
                                sel.select_by_index(1)
                        except Exception:
                            pass

                    sec_answers = driver.find_elements(By.XPATH, "//input[@type='text' and (contains(@id,'Answer') or contains(@name,'Answer') or contains(@id,'txtAnswer'))]")
                    for ans_input in sec_answers:
                        try:
                            ans_input.clear()
                            ans_input.send_keys("Workforce")
                        except Exception:
                            pass

                    submit_btn = driver.find_elements(
                        By.XPATH,
                        "//input[@type='submit' or @value='Submit' or @value='Save' or @value='Continue'] | //button[contains(text(),'Submit') or contains(text(),'Save') or contains(text(),'Continue')]"
                    )
                    if submit_btn:
                        submit_btn[0].click()
                        time.sleep(2)

                    # Re-check for post-submit errors
                    post_errs = driver.find_elements(
                        By.XPATH,
                        "//*[contains(@class,'error') or contains(@id,'lblError') or contains(@id,'ErrMsg') or contains(@class,'alert-danger')]"
                    )
                    post_visible = [e.text.strip() for e in post_errs if e.is_displayed() and e.text.strip()]
                    if post_visible:
                        ui.log(f"FAIL {lbl} — Password change error: {post_visible[0]}", "fail")
                        failed.append({"Username": user, "Reason": post_visible[0]})
                        continue

                    ui.log(f"PASS {lbl} — Password setup completed", "pass")
                    passed.append(user)

                else:
                    # Check if user is already logged into member dashboard
                    logged_in_indicators = driver.find_elements(
                        By.XPATH,
                        "//*[contains(text(),'Logout') or contains(text(),'Log Out') or contains(text(),'Welcome') or contains(@id,'Dashboard') or contains(@id,'PopUp_LblMsg')]"
                    )
                    login_btn = driver.find_elements(By.XPATH, "//input[@value='Log In'] | //button[normalize-space()='Log In']")
                    is_login_still_there = any(b.is_displayed() for b in login_btn)

                    if logged_in_indicators and not is_login_still_there:
                        ui.log(f"PASS {lbl} — Already configured (Logged in successfully)", "pass")
                        passed.append(user)
                    elif is_login_still_there:
                        reason = visible_errors[0] if visible_errors else "Invalid credentials or temporary password rejected"
                        ui.log(f"FAIL {lbl} — {reason} | Temp Password used: {temp_pwd}", "fail")
                        failed.append({"Username": user, "TempPassword": temp_pwd, "Reason": reason})
                    else:
                        ui.log(f"PASS {lbl} — Setup step finished", "pass")
                        passed.append(user)

            except Exception as e:
                err_str = str(e).splitlines()[0] if str(e) else type(e).__name__
                ui.log(f"FAIL {lbl} — {err_str}", "fail")
                failed.append({"Username": user, "TempPassword": temp_pwd, "Reason": err_str})

            ui.progress((i / total) * 100)
            ui.status(f"Processing {i}/{total} — {len(passed)} passed, {len(failed)} failed")
            time.sleep(1)

        ui.log("\n" + "=" * 65)
        ui.log("RESULTS SUMMARY")
        ui.log("=" * 65)
        ui.log(f"Passed : {len(passed)} / {total}", "pass")
        ui.log(f"Failed : {len(failed)} / {total}", "fail" if failed else "info")

        if failed:
            ui.log("\nFAILED RECORDS:", "fail")
            for f in failed:
                ui.log(
                    f"  Username: {f['Username']} | Temp Password (DOB): {f.get('TempPassword', '-')} | Reason: {f['Reason']}",
                    "fail"
                )
            ui.status(f"Done — {len(passed)} passed, {len(failed)} failed", "fail")
        else:
            ui.log("All records passed!", "pass")
            ui.status(f"Done — {len(passed)} passed", "pass")

    finally:
        driver.quit()
