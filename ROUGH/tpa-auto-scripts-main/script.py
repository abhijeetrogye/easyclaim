import re
import os
import time
import pandas as pd 


from playwright.sync_api import Playwright, sync_playwright, expect

# --- CONFIGURATION ---
EXCEL_FILE = "NEW PORTAL.xlsx"
DOWNLOAD_DIR = "DataPDF"


def run(playwright: Playwright) -> None:
    
    
    # browser = playwright.chromium.launch(headless=False)
    # context = browser.new_context()
    # page = context.new_page()
    
    user_data_dir = "chrome-profile"
    context = playwright.chromium.launch_persistent_context(
        user_data_dir,
        headless=False
    )

    page = context.new_page()
    page.goto("https://provider.nha.gov.in/")

    try:
        df = pd.read_excel(EXCEL_FILE)

        # FIX: Explicitly convert 'Status' and 'Case_ID' to string
        if 'Status' not in df.columns:
            df['Status'] = ""
        else:
            df['Status'] = df['Status'].astype(str)

        if 'Case_ID' not in df.columns:
            df['Case_ID'] = ""
        else:
            df['Case_ID'] = df['Case_ID'].astype(str)

        
        print("👉 Please complete LOGIN manually in the browser.")
        print("👉 After login is successful, press ENTER here to continue...")
        input()
        
        page.get_by_role("button", name="CLOSE").click()
        print("Login Success")

        #===============
        #PREAUTH PROCESS STARTS
        #===============
        registration_id = "1015795522"

        search_box = page.locator(
       'xpath= //*[@id="root"]/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/input')
            
        search_box.click()
        # search_box.fill("")
        search_box.type(registration_id)

        # Click search button inside the same bar
        search_button = page.locator('xapth= //*[@id="root"]/div[2]/div[1]/div[2]/div[3]/div[1]/div[2]/div/div[4]/div/span')

        page.locator('xpath =//*[@id="root"]/div[2]/div[1]/div[2]/div[3]/div[2]/div/div/div/div").wait_for(state="visible')

        page.locator("#Path_98789").first.click()

        #===============
        #Medical inforamtion STAGE 1
        #===============
        print("Stage 1 Started")

        print("Medical Information entering")
        page.get_by_role("button", name="Medical Information").click()
        print("Entering the general finding")
        page.get_by_role("button", name="General Findings").click()
        print("Entered general findings")
        page.get_by_role("group").filter(has_text="Cyanosis Yes No").locator("label").nth(2).click()
        print("Data feeded in general finding")
        page.get_by_role("button", name="SAVE").click()
        print("Data Saved")
        page.get_by_role("button", name="General Findings").click()
        print("General Findings Completed")

        print("Entering Family History")
        page.get_by_role("button", name="Family History").click()
        page.get_by_role("button", name="Family History").click()
        print("Done Family History")




        #===============
        #ADMISSION inforamtion STAGE 2
        #===============
        print("Stage 2 Started")

        ## Admission Details 

        print("Entering the Admission Information")
        page.get_by_role("button", name="Admission Information").click()
        print("Admission Information Success")
        print("Entering the Admission Details")
        page.get_by_role("button", name="Admission Details").click()
        print("Admission Details Success")
        time.sleep(2)
        
        save_btn = page.get_by_role("button", name="SAVE").click()
        print("Save Button ")
        edit_btn = page.get_by_role("button", name="Edit").click()
        print("Edit button")
        
        # -------- Ensure Edit Mode ----------
        if save_btn.is_visible():
            print("Already in edit mode (SAVE button visible)")
        else:
            edit_btn.click()
            print("Edit button clicked")
            page.wait_for_timeout(2000)

        print("Date Imported from the excel")
        time.sleep(2)
        #page.locator("div").filter(has_text=re.compile(r"^Surgery Treatment Start Date \*Clear28/12\/2025$")).get_by_role("img").click()
        date_picker = page.locator("div").filter(has_text=re.compile(r"^Surgery Treatment Start Date \*DD/MM\/YYYY$")).locator("#Rectangle_1171").dblclick()
        page.get_by_role("button", name="Select Sunday, December 28,").click()

        print("Date Field Locator Success")

        from datetime import datetime

        time.sleep(1)
   
        print("Date Successfully Added")


        page.get_by_role("group").locator("label").filter(has_text="No").locator("span").dblclick()
        print("Medico Selected Successfull")

        admission_type = "PLANNED"
        print("Admission type imported")
        page.locator("#AdmissionType > .css-31y905-control > .css-14ntbqx > .css-ackcql").click()
        print("selected the dropdown")

        if admission_type =="PLANNED":
            page.get_by_text("PLANNED", exact=True).click()
        elif admission_type == "EMERGENCY":
            page.get_by_text("EMERGENCY", exact=True).click()
        else:
            raise Exception(f"Unknown Admission_Type: {admission_type}")

        page.get_by_role("button", name="SAVE").click()
        print("Admission Type Added Successfully")


        #===============
        #TREATMENT inforamtion STAGE 3
        #===============
        print("Stage 3 Started")

    
        page.get_by_role("button", name="Treatment").click()
        print("Entered the Treatment Protocols")

        ## Referral Details
        page.locator(".NJ6Bg2V2FKKAXwJJJSvT.ml-auto > .accordion-button").first.click()
        page.locator("div").filter(has_text=re.compile(r"^Manual$")).click()
        
        #e-Referral Details
        page.get_by_text("e-Referral").click()
        page.locator("[id=\"1002\"]").check()
        page.get_by_role("textbox", name="Type here").dblclick()
        page.get_by_role("button", name="Search").click()
    
        #Manual Details
        page.get_by_text("Manual").click()
        page.locator("[id=\"1001\"]").check()
        page.locator("input[name=\"Referralnumber\"]").click()
        page.locator("input[name=\"Referralnumber\"]").type("123R24353")
        page.locator(".col-lg-3.col-md-6 > .formgroup > .css-b62m3t-container > .css-31y905-control > .css-14ntbqx > .css-ackcql").click()
        page.locator("#react-select-30-input").press("ArrowDown")
        page.get_by_text("AGARTALA(AG1)", exact=True).click()
        page.locator("#react-select-30-input").type("23")
        page.get_by_text("DD/MM/YYYY").click()
        page.get_by_role("button", name="Currently selected Saturday,").click()
        page.locator("input[name=\"Referralendoresement\"]").click()
        page.locator("input[name=\"Referralendoresement\"]").type("TYT")
        page.locator("input[name=\"endoresementhosp\"]").click()
        page.locator("input[name=\"endoresementhosp\"]").type("MAX")
        page.get_by_role("textbox", name="Mobile Number").click()
        page.get_by_role("button", name="Drag your file here  Or").click()
        page.get_by_role("button", name="Drag your file here  Or").set_input_files("NSCR78356.pdf")
        page.locator("div:nth-child(2) > .Qnc3lxFrZByXRLyQliMJ > .accordion-item > .UmpLaTf6zoWiOScAnnMK > .row.XkJXDZHMARVubcuOItEA > .NJ6Bg2V2FKKAXwJJJSvT > .accordion-button").click()
    
    
    #Diagonse Details
        page.get_by_role("textbox", name="Enter Diagnosis").click()
        print("Clicked the Diagnose box")
        diagnosis = "XRAY"
        print("Data Of Daignosis is loaded")

        if diagnosis:
         page.get_by_role("textbox", name="Enter Diagnosis").type(diagnosis)
        print("Searched Diagnosis from the excell")

        selected_diag = "Other specified syndromic intellectual disability"
        print("Data Imported from  excel")

        page.get_by_role("textbox", name="Selected Diagnosis").click()
        print("Diagnose box selected")
        page.get_by_role("textbox", name="Selected Diagnosis").type(selected_diag)
        print("Typing the Exact Name From the excel")
        page.get_by_text(selected_diag).click()
        print("Done Selected Diagnosis")
        #Diagnosis Type
        page.get_by_role("radio", name="Manual Primary").check()
        print("Diagnosis Type Selected Properly")
        page.get_by_role("button", name="ADD").click()
        print("Diagnosis Added Successully")


        #Treatment Details
        page.locator(".NJ6Bg2V2FKKAXwJJJSvT.rQOcJnFMaOZStyAKvox1 > .accordion-button").click()
        print("Entered the Treatment Plan Successfull")

        #Specialty
        checkup_type = "Annual Health Check-up"
        print("Data Loaded fromthe Excel")
        if checkup_type:
                page.locator(".col-lg-10 > .formgroup > .css-b62m3t-container > .css-31y905-control > .css-14ntbqx > .css-ackcql").first.click()
                try:
                 page.get_by_text(checkup_type, exact=True).click()
                 print("Exact Match Found")
                except:
                        page.get_by_text(checkup_type, exact=False).first.click()
                        print("Exact Match Not Found so Fallback")
        else:
         print("Checkup_Type is empty in Excel")
        print("Speacialty Success")

        #Procedures
        page.locator(".col-lg-10 > .formgroup > .css-b62m3t-container > .css-31y905-control > .css-14ntbqx > .css-ackcql").click()
        page.locator("#react-select-32-option-0").click()
        
        #No of Days/Units
        text_value = "2"
        print("Days & Units data imported")
        if text_value:
            page.get_by_role("textbox", name="Type here", exact=True).click()
            print("Textbox Selected")
        page.get_by_role("textbox", name="Type here", exact=True).type(text_value)
        print("No of Days & Units entered")
        
        #Add Treatment Plan Button
        page.locator(".m9FzljqXbDJyFhzambbf").click()
        print("Treatment Plan Added Successfully")

        #Investigation/Documents Details
        page.locator(".M8Nwg9n1mRYxBEjrNzE5 > .UmpLaTf6zoWiOScAnnMK > .row.XkJXDZHMARVubcuOItEA > .NJ6Bg2V2FKKAXwJJJSvT > .accordion-button").first.click()
        print("Entered Investigatin/Documents Success")
        # page.locator(".col-lg-4 > div > .css-b62m3t-container > .css-31y905-control > .css-1wy0on6 > .css-tlfecz-indicatorContainer > .css-8mmkcg").click()
        
        #CGHS Card Copy 
        bill_no = "NSCR78356"
        print("Bill no extracted from the excel")
        BASE_PDF_DIR = r"C:\Users\Stunning\Desktop\CGHS NEW\DataPDF"
        print("Folder path given")
        pdf_path = os.path.join(BASE_PDF_DIR, f"{bill_no}.pdf")
        print("Fetchs the data from the folder")
        if os.path.exists(pdf_path):
            page.get_by_role("row", name="CGHS Card Copy NA").get_by_role("button").click()
            page.get_by_role("row", name="CGHS Card Copy NA").get_by_role("button").set_input_files("NSCR78356.pdf")
            print(f"Uploaded Card Copy for Bill No: {bill_no}")
        else:
            print(f"Card Copy PDF not found for Bill No: {bill_no}")

        #Patient Photo
        bill_no = "NSCR78356"
        print("Bill no extracted from the excel")
        BASE_PDF_DIR = r"C:\Users\Stunning\Desktop\CGHS NEW\DataPDF"
        print("Folder path given")
        pdf_path = os.path.join(BASE_PDF_DIR, f"{bill_no}.pdf")
        print("Fetchs the data from the folder")
        if os.path.exists(pdf_path):
            page.get_by_role("button", name="Drag your file here  Or").click()
            page.get_by_role("button", name="Drag your file here  Or").set_input_files("NSCR78356.pdf")
            print(f"Uploaded Patient photo for Bill No: {bill_no}")
        else:
            print(f"Patient Photo PDF not found for Bill No: {bill_no}")


        #Care Team Doc Fields
        page.locator("div:nth-child(5) > .Qnc3lxFrZByXRLyQliMJ > .M8Nwg9n1mRYxBEjrNzE5 > .UmpLaTf6zoWiOScAnnMK > .row.XkJXDZHMARVubcuOItEA > .NJ6Bg2V2FKKAXwJJJSvT > .accordion-button").click()
        print("Entered the Care Team")

        
        doctor_name = "Dr S A Agarwal"
        print("Data Loaded")
        # ---- Dropdown selection ----
        page.locator(".my-auto > div > .css-b62m3t-container > .css-31y905-control > .css-14ntbqx > .css-ackcql").click()
        print("Selected the dropdown")

        page.get_by_text("Other", exact=True).click()
        
        # ---- Doctor name textbox ----
        if doctor_name:
            page.get_by_role("textbox", name="Doctor's Name").click()
        page.get_by_role("textbox", name="Doctor's Name").type(doctor_name)
        
        #------Doctor Registration ID---------
        page.get_by_role("textbox", name="Registration ID / HPR ID").click()
        print("Selected the Textbox")
        page.get_by_role("textbox", name="Registration ID / HPR ID").type("1234")
        print("Registration Id Success")

        #------Qualification Details---------

        page.locator("#qualification svg").click()
        #page.locator(".css-cv8flm-control > .css-14ntbqx > .css-ackcql").click()
        page.locator(".css-cv8flm-control > .css-14ntbqx > .css-ackcql").click()
        print("Text box Selected")
        page.get_by_text("M.B.B.S.", exact=True).click()
        print("Qualification Success")
        
        #--------Mobile Number-------

        page.get_by_role("textbox", name="Contact No.").click()
        print("Selected the textbox")
        page.get_by_role("textbox", name="Contact No.").type("9000000000")
        print("Number Added Success")
        page.get_by_role("button", name="ADD").nth(2).click()
        print("Added Doc Name Success")
        
        #page.get_by_role("button", name="Treatment").click()
        #page.get_by_role("button", name="Finance").click()
        
        page.get_by_role("button", name="Preview & Validate").click()
        page.locator("div").filter(has_text="validate EDIT").nth(3).click()
        page.get_by_role("button", name="validate", exact=True).click()
        page.locator("div").filter(has_text="Initiate Pre-Authorization").nth(3)
        page.locator("div").filter(has_text="Initiate Pre-Authorization").nth(3)
        page.get_by_role("button", name="EDIT").click()
        page.get_by_role("button", name="Home").click()
        page.get_by_role("textbox", name="Search").click()
        page.get_by_role("textbox", name="Search").fill("1015795522")
        page.locator(".icon").click()
        page.locator("#Path_98789").click()
    
    except Exception as e:
        print("Script failed:", e)
        
    finally:
        if page:
            try:
                print("Attempting logout...")
                page.get_by_role("button", name="Dinesh Kumar Sharma").click()
                page.get_by_role("button", name="Logout").click()
                print("Logged out successfully")
            except Exception as logout_err:
                print("Logout failed:", logout_err)

        try:
            print('to_excel')
            # df.to_excel(EXCEL_FILE, index=False)
        except Exception as save_err:
            print("Failed to save excel:", save_err)

        # if context:
        #     try: context.close()
        #     except: pass
        # if browser:
        #     try: browser.close()
        #     except: pass


with sync_playwright() as playwright:
    run(playwright)
