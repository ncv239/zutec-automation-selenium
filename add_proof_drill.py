from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

import pandas as pd
import time

from credentials import url, zutec_user_email, zutec_user_pwd, scs_email, signature_pin, fld_construction_package_code, fld_primary_asset_name, fld_construction_package_description


def read_data(path):
    dtype = {
        "ring_number": str,
        "segment_type": str,
        "segment_id": str,
        "location_drilled": str,
        "grout_thickness": str,
        "water_flow": str,
        "voids": str,
        "sika_swell_oring_installed": bool,
        "secondary_grouting_required": bool,
        "shift_eng_name": str,
        "proofdrill_by": str,
        "proofdrill_date": str,
        "proofdrill_hour": str,
        "shift": str,
        "drive": str,
        "subcontractor_itp_inspection_requirements": str,
        "scs_itp_inspection_requirements": str,
        "scs_checked_by": str,
        "scs_date": str,
        "scs_date_hour": str,
    }
    df = pd.read_excel(path, dtype=dtype)
    return df


def element_to_be_clickable(driver, timeout: int, by=By.ID, selector: str = ""):
    for i in range(3):  # try three times
        try:
            return WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, selector)))
        except:
            time.sleep(1)



def login(driver, url):
    ''' Open the ZUTEC url and perform a log-in'''
    driver.get(url)
    # login
    element_to_be_clickable(driver, 10, By.ID, "email-input").send_keys(zutec_user_email)
    element_to_be_clickable(driver, 10, By.ID, "password-input").send_keys(zutec_user_pwd)
    element_to_be_clickable(driver, 10, By.ID, "login-submit-button").click()


def navigate_to_proof_drill(driver):
    ''' After logged in to ZUTEC, navigate in the side bar to the proof drill checksheet section'''
    # click the logo
    #element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "#zone-item-2-143 > div > div > div").click()

    # click the menu to the Proof Drilling page
    WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#folderframe")))

    # 04. Quality and Inspection Management
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "li[id='24'] > a").click()
    # 04.3 West
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "li[id='34'] > a").click()
    # 04.3.07 Tunneling
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, "li[id='758'] > a"))).click()
    #West-Insitu Tunnel Segment Repair
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "li[id='5239'] > a").click()
    # West-Proof Drilling Checksheet-TEM Required
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "li[id='5240'] > a").click()
    time.sleep(1)


def add_proof_drill_record(driver, data, save_checksheet=False):
    ''' After navigated to the proofdrill checksheet section,
    Create a ZUTEC Checksheet for ProofDriling  (!) UPLINE (!) based on a data from a single row of excel table
    '''
    ring_number = data["ring_number"]
    segment_type = data["segment_type"]
    segment_id = data["segment_id"]
    location_drilled = data["location_drilled"]
    grout_thickness = data["grout_thickness"]
    water_flow = data["water_flow"]
    voids = data["voids"]
    sika_swell_oring_installed = data["sika_swell_oring_installed"]
    secondary_grouting_required = data["secondary_grouting_required"]
    shift_eng_name = data["shift_eng_name"]
    proofdrill_by = data["proofdrill_by"]
    proofdrill_date = data["proofdrill_date"]
    proofdrill_hour = data["proofdrill_hour"]
    shift = data["shift"]
    drive = data["drive"]
    subcontractor_itp_inspection_requirements = data["subcontractor_itp_inspection_requirements"]
    scs_itp_inspection_requirements = data["scs_itp_inspection_requirements"]
    scs_checked_by = data["scs_checked_by"]
    scs_date = data["scs_date"]
    scs_date_hour = data["scs_date_hour"]

    driver.switch_to.default_content()
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#sectionframe")))

    element_to_be_clickable(driver, 10, By.ID, "btn_add_record").click()

    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='fld_scsjv_project_area']").click()
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='fld_scsjv_project_area']").send_keys("WEST")

    time.sleep(1) # wait to load
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='fld_primary_asset_name']").send_keys(fld_primary_asset_name)

    time.sleep(1) # wait to load
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='fld_construction_package_description']").send_keys(fld_construction_package_description)
    time.sleep(2) # wait to load
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='fld_construction_package_code']").send_keys(fld_construction_package_code)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_construction_schedule_id']").send_keys("R"+ring_number)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_design_schedule_id']").send_keys("R"+ring_number)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_element_id']").send_keys("TBC")
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_hs2_unique_asset_identifier_uaid']").send_keys("TBC")
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_itp_reference']").send_keys("TBC")
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_bim_link']").send_keys("TBC")
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_subcontractor_name']").send_keys("Self-Delivered")
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_inspection_team_scs']").send_keys(shift_eng_name)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_shift_engineer']").send_keys(shift_eng_name)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_proof_drill_done_by']").send_keys(proofdrill_by)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='frmdate_show_fld_drill_date']").send_keys(proofdrill_date)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='frmtime_hours_fld_drill_date']").send_keys(proofdrill_hour)

    shift = shift.upper()
    if (shift == "DS"):
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_shift0']").click()
    elif (shift == "BS"):
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_shift1']").click()
    elif (shift == "NS"):
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_shift2']").click()


    if (drive.lower() == "upline"):
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_drive0']").click()
    else:
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_drive1']").click()


    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_ring_number_1']").send_keys(ring_number)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_segment_type_1']").send_keys(segment_type)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_segment_id_1']").send_keys(segment_id)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_location_of_segment_drilled_1']").send_keys(location_drilled)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_grout_thickness_mm_1']").send_keys(grout_thickness)



    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_water_flow_1']").send_keys(water_flow)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_voids_1']").send_keys(voids)

    if sika_swell_oring_installed:
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_sika_swell_o_ring_grout_plug_10']").click()
    else:
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_sika_swell_o_ring_grout_plug_11']").click()


    if secondary_grouting_required:
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_secondary_grouting_required_10']").click()
    else:
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='button_fld_secondary_grouting_required_11']").click()


    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='fld_2_scs_checked_by']").send_keys(scs_checked_by)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "input[id='frmdate_show_fld_2_scs_date']").send_keys(scs_date)
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "select[id='frmtime_hours_fld_2_scs_date']").send_keys(scs_date_hour)


    #Subcontractor - ITP Inspection Requirements
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "div[id='fld_1_sub_con_itp_inspection_requirementsvalue'] iframe")))
    inputField = element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "body")
    driver.execute_script("arguments[0].innerHTML = arguments[1]", inputField, subcontractor_itp_inspection_requirements)

    # SCS - ITP Inspection Requirements
    driver.switch_to.default_content()
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#sectionframe")))
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "div[id='cke_frm_fld_2_scs_itp_inspection_requirements'] iframe")))
    inputField = element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "body")
    driver.execute_script("arguments[0].innerHTML = arguments[1]", inputField, scs_itp_inspection_requirements)

    # now click the signature button
    driver.switch_to.default_content()
    WebDriverWait(driver, 30).until(EC.frame_to_be_available_and_switch_to_it((By.CSS_SELECTOR, "#sectionframe")))
    # click the button
    signField = element_to_be_clickable(driver, 10, By.XPATH, "/html/body/form/table/tbody/tr/td/table/tbody/tr/td/div[7]/div[17]/table/tbody/tr/td/div[19]")
    signFieldId = signField.get_attribute("id").split("_")[1]  # this id is dynamic, changes on every refresh
    signField.click()

    time.sleep(3)
    # find email field
    emailField = element_to_be_clickable(driver, 10, By.CSS_SELECTOR, f"input[id='sig_{signFieldId}_email']")
    emailField.click()
    emailField.send_keys(scs_email)
    time.sleep(1)
    emailField.send_keys(Keys.ARROW_DOWN)
    time.sleep(1)
    emailField.send_keys(Keys.ENTER)
    time.sleep(1)
    # WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CSS_SELECTOR, f"#ui-id-19 > li > div"))).click()

    #enter PIN
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, f"input[id='sig_{signFieldId}_pin_pinlogin_0']").send_keys(signature_pin[0])
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, f"input[id='sig_{signFieldId}_pin_pinlogin_1']").send_keys(signature_pin[1])
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, f"input[id='sig_{signFieldId}_pin_pinlogin_2']").send_keys(signature_pin[2])
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, f"input[id='sig_{signFieldId}_pin_pinlogin_3']").send_keys(signature_pin[3])
    time.sleep(3)
    # click SAVE button
    element_to_be_clickable(driver, 10, By.CSS_SELECTOR, f"div[aria-describedby='sig_{signFieldId}_dialog'] div.ui-dialog-buttonset > button:nth-child(2)").click()

    if save_checksheet:  # click the save button
        element_to_be_clickable(driver, 10, By.CSS_SELECTOR, "#savebutton_top").click()
    else:  # click the cancel button or small cancell button
        try:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#cancelbutton"))).click()
        except:
            WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#cancelbutton_small"))).click()
        
        # if confirmation alert pops up
        try:
            WebDriverWait(driver, 10).until(EC.alert_is_present())
            alert = driver.switch_to.alert
            alert.accept()
        except:
            pass
    time.sleep(5)


def print_info():
    info_msg = '''
Running script to automatically create Zutec Proofdrill checksheet
Important Notes!
- prior to use download and place the chromedriver.exe file in the script directory
- when running, the script will open a browser, make it full-screen (otherwise css-selectors wont work)
- make sure to change credentials in the credentials.py file
- make sure to fill the data excel sheet properly (e.g. text-type for dates). Read comments in headers.
'''
    print(info_msg)


def main():
    print_info()
    path_to_excel = "./test_data.xlsx"
    data = read_data(path_to_excel)

    driver = webdriver.Chrome()  # Optional argument, if not specified will take from current directory

    login(driver, url)
    navigate_to_proof_drill(driver)

    for data_row in data.iterrows():
        add_proof_drill_record(driver, data_row[1], save_checksheet=False)
        print("R", data_row[1]["ring_number"], " - successfully created check-sheet", )

    time.sleep(60)
    driver.quit()


main()