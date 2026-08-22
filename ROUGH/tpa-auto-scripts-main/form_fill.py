import re
import os
import time
import pandas as pd 
from datetime import datetime
from playwright.sync_api import sync_playwright, expect

# user name: dineshdwarka
def run_automation():
    
    with sync_playwright() as playwright:
        try:
            print("connecting to browser...")
            browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
            print("connected successfully!")
            
            context = browser.contexts[0]
            pages = context.pages
            
            if not pages:
                print("❌ No pages found. Creating new page...")
                page = context.new_page()
                page.goto("https://provider.nha.gov.in/")
            else:
                page = pages[0]
                print(f"✅ Using existing page: {page.url}")
            
            print("\n📊 Loading Excel data...")
            excel_path = "entries.xlsx"

            if not os.path.exists(excel_path):
                print(f"❌ Error: {excel_path} not found!")
                return

            df = pd.read_excel(excel_path)
            total_rows = len(df)
            print(f"✅ Found {total_rows} entries to process\n")

            # Track results for output Excel
            results = []

            login_first(page)

            for idx, row in df.iterrows():
                print(f"\n{'='*60}")
                print(f"📋 Processing Entry {idx + 1}/{total_rows}")
                print(f"{'='*60}")

                try:
                    # Extract data from Excel row
                    reg_id = str(row['Registration ID*'])
                    diagnosis = str(row['Diagnose*'])
                    doctor_name = str(row['Doctor Name*'])
                    doctor_id = str(row['Doctor Id *'])

                    # Parse procedures and amounts (separated by |)
                    procedures_str = str(row['Procedures*'])
                    amounts_str = str(row['Amounts'])

                    procedures = [p.strip() for p in procedures_str.split('|')]
                    amounts = [a.strip() for a in amounts_str.split('|')]

                    # File paths
                    card_file = str(row['Card File*'])
                    patient_photo = str(row['Patient Photo*'])

                    print(f"  Registration ID: {reg_id}")
                    print(f"  Diagnosis: {diagnosis}")
                    print(f"  Doctor: {doctor_name}")
                    print(f"  Procedures: {len(procedures)}")
                    for i, (proc, amt) in enumerate(zip(procedures, amounts)):
                        print(f"    {i+1}. {proc[:50]}... - ₹{amt}")

                    # Build details dictionary
                    details = {
                        'reg_id': reg_id,
                        'diagnosis': diagnosis,
                        'doctor_name': doctor_name,
                        'contact_no': doctor_id,
                        'procedures': procedures,
                        'amounts': amounts,
                        'card_file': card_file,
                        'patient_photo': patient_photo
                    }

                    # Process this claim
                    search_claim_entry(page, details['reg_id'])
                    process_medical_info(page)

                    input("\n⏸️  Complete ADMISSION INFO manually, then press Enter to continue...")

                    process_treatment(page, details)

                    print(f"\n✅ Entry {idx + 1} completed successfully!")

                    # Record success
                    results.append({
                        'Registration ID': reg_id,
                        'Status': 'SUCCESS',
                        'Error': '',
                        'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })

                except Exception as row_error:
                    print(f"\n  ❌ Error processing entry {idx + 1}: {row_error}")
                    import traceback
                    traceback.print_exc()

                    # Record failure
                    results.append({
                        'Registration ID': reg_id if 'reg_id' in locals() else 'Unknown',
                        'Status': 'FAILED',
                        'Error': str(row_error),
                        'Timestamp': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                    })

                    # Ask user to manually navigate to home page
                    print("\n" + "="*50)
                    print("⚠️  ENTRY FAILED - MANUAL ACTION REQUIRED")
                    print("="*50)
                    print("Please manually navigate to the HOME PAGE (claims list)")
                    print("Make sure you can see the search box and RefreshList button")
                    input("\n✅ Press Enter when you are on the HOME PAGE to continue with next entry...")

                    continue  # Continue to next entry

            # Save results to output Excel
            print("\n" + "=" * 60)
            print("📊 Saving results to output Excel...")
            output_df = pd.DataFrame(results)
            output_file = f"output_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            output_df.to_excel(output_file, index=False)
            print(f"✅ Results saved to: {output_file}")

            # Print summary
            success_count = len([r for r in results if r['Status'] == 'SUCCESS'])
            failed_count = len([r for r in results if r['Status'] == 'FAILED'])
            print(f"\n📈 Summary: {success_count} SUCCESS, {failed_count} FAILED out of {total_rows} entries")
            print("=" * 60)
            
        except ConnectionRefusedError:
            print("could not connect to browser")

        except Exception as e:
            print(f"\n script failed: {e}")
            import traceback
            traceback.print_exc()

def login_first(page):
    print("login sucess")
    #page.get_by_role("button", name="CLOSE").click()
    #print("initial prompt closed")

def search_claim_entry(page, reg_id):
    page.locator("div > svg + p:has-text('RefreshList')").locator("..").click()


    search_box = page.locator(".form-group:has(input[placeholder='Search'])")
    search_box.wait_for(state="visible")

    input_box = search_box.locator("input")

    input_box.click()
    input_box.clear()
    input_box.type(reg_id, delay=100)
    print("reg id entered")

    search_box.locator("span.icon").click(force=True)
    print("reg id searched")
    
    
    page.locator("div > svg + p:has-text('RefreshList')").locator("..").click()

    page.wait_for_timeout(2000)

    # Click the claim entry icon using XPath
    claim_icon = page.locator("xpath=//*[@id='Path_98789']").first
    claim_icon.wait_for(state="visible", timeout=10000)
    claim_icon.click()

    page.wait_for_timeout(2000)
    print(f"  ✓ Opened claim entry")




def process_medical_info(page):
    print("\n📝 Processing Medical Information section...")

    medical_btn = page.locator(
        "button:has(h4:has-text('Medical Information'))"
    ).first
    medical_btn.wait_for(state="visible", timeout=10000)
    medical_btn.scroll_into_view_if_needed()
    page.wait_for_timeout(500)

    # Always click to expand Medical Information section
    print("  Clicking Medical Information to expand...")
    medical_btn.click(force=True)
    page.wait_for_timeout(2000)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)

    # Find Personal History section - use text to identify correct one
    print("  Looking for Personal History section...")

    # Find the container that has "Personal History" text inside it
    personal_section = page.locator("div.Qnc3lxFrZByXRLyQliMJ:has(span:text('Personal History'))").first
    personal_section.wait_for(state="visible", timeout=15000)
    personal_section.scroll_into_view_if_needed()
    print("  Found Personal History container")

    # Find the accordion header button inside
    personal_header_btn = personal_section.locator("h2.accordion-header button").first
    personal_header_btn.wait_for(state="visible", timeout=10000)

    is_expanded = personal_header_btn.get_attribute("aria-expanded")
    print(f"  Personal History expanded: {is_expanded}")
    if is_expanded == "false":
        print("  Expanding Personal History...")
        # Try clicking the span text directly
        span_text = personal_section.locator("span:has-text('Personal History')").first
        span_text.click(force=True)
        page.wait_for_timeout(3000)

        # Check again if expanded
        is_expanded = personal_header_btn.get_attribute("aria-expanded")
        print(f"  After click, expanded: {is_expanded}")

    # Wait for accordion body
    print("  Waiting for accordion content to load...")
    page.wait_for_timeout(2000)

    # Check if already in edit mode (SAVE button visible) or need to click EDIT
    accordion_body = personal_section.locator(".accordion-body").first
    save_btn = accordion_body.locator("button:has-text('SAVE')").first
    edit_btn = accordion_body.locator("button.QoeYZqaZIdTkfax4H_yz").first

    # Check if SAVE button is already visible (already in edit mode)
    if save_btn.is_visible():
        print("  Already in edit mode (SAVE button visible)...")
    else:
        # Need to click EDIT button first
        print("  Looking for Edit button...")
        if not edit_btn.is_visible():
            print("  Edit button not visible, clicking header again...")
            personal_header_btn.click(force=True)
            page.wait_for_timeout(2000)

        print("  Clicking Edit button...")
        edit_btn.wait_for(state="visible", timeout=10000)
        edit_btn.scroll_into_view_if_needed()
        edit_btn.click(force=True)
        page.wait_for_timeout(2000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1000)

    # Select "No" for Known Allergies
    print("  Selecting Known Allergies: No...")
    known_allergies_no = page.locator("input#KnownAllergies[value='No']").first
    known_allergies_no.scroll_into_view_if_needed()
    page.wait_for_timeout(500)
    known_allergies_no.click(force=True)

    # Select "No" for Habits Addictions
    print("  Selecting Habits Addictions: No...")
    habits_no = page.locator("input#HabitsAddictions[value='No']").first
    habits_no.scroll_into_view_if_needed()
    page.wait_for_timeout(500)
    habits_no.click(force=True)

    page.wait_for_timeout(1000)

    # Click SAVE button
    print("  Clicking SAVE button...")
    save_btn = personal_section.locator("button:has-text('SAVE')").first
    save_btn.wait_for(state="visible", timeout=10000)
    save_btn.click()
    page.wait_for_timeout(2000)

    # Close Personal History accordion
    if personal_header_btn.get_attribute("aria-expanded") == "true":
        personal_header_btn.click()
        page.wait_for_timeout(1000)

    # Close Medical Information section
    if medical_btn.get_attribute("aria-expanded") == "true":
        medical_btn.click()
        page.wait_for_timeout(1000)

    print("✅ Medical Information completed!")


def process_admission_info(page, details):
    print("Opening ADMISSION INFORMATION section...")
    admission_btn = page.locator(
        "button:has(h4:has-text('ADMISSION INFORMATION'))"
    ).first
    admission_btn.scroll_into_view_if_needed()
    admission_btn.click()

    page.wait_for_timeout(1000)

    print("Opening Admission Details accordion...")
    admission_header_btn = page.locator(
        "button.accordion-button:has(span:has-text('Admission Details'))"
    ).first

    admission_header_btn.wait_for(state="visible")

    if admission_header_btn.get_attribute("aria-expanded") == "false":
        admission_header_btn.click()

    page.wait_for_timeout(2000)

    print("Clicking EDIT button...")
    edit_btn = page.get_by_role("button", name="EDIT").first
    edit_btn.wait_for(state="visible")
    edit_btn.click()

    page.wait_for_timeout(2000)
    
    medico_case_radio = page.locator(
        "fieldset:has(label:has-text('Medico Legal Case')) input[value='No']"
    )

    medico_case_radio.wait_for(state="attached")
    medico_case_radio.check(force=True)
    
    page.wait_for_timeout(2000)
    
    page.locator("//div[@id='AdmissionType']//input[@role='combobox']").click()
    page.wait_for_selector("//*[contains(@id,\"listbox\")]")
    page.locator(
        '//*[contains(@id,"listbox")]//div[normalize-space()="PLANNED"]'
    ).click()
    
    
    date_string = '29/12/2025'
    day, month, year = date_string.split('/')
    day = int(day)
    month = int(month)
    year = int(year)
    
    months = ['January', 'February', 'March', 'April', 'May', 'June', 
              'July', 'August', 'September', 'October', 'November', 'December']
    target_month_name = months[month - 1]
    
    print("Clicking Surgery Treatment Start Date field...")
    date_label = page.locator("label:has-text('Surgery Treatment Start Date')").first
    date_label.scroll_into_view_if_needed()
    date_input = date_label.locator("..").locator("input").first
    date_input.click()

    page.wait_for_selector(".sdp--square-btn", state="visible", timeout=10000)
    
    max_attempts = 24
    for attempt in range(max_attempts):
        current_display = page.locator("p.sdp--month-name").text_content().strip()
        print(f"Current: {current_display}, Target: {target_month_name} {year}")
        
        if f"{target_month_name} {year}" == current_display:
            print("Found target month/year!")
            break
        
        try:
            current_parts = current_display.split()
            current_month_name = current_parts[0]
            current_year = int(current_parts[1])
            current_month = months.index(current_month_name) + 1
        except:
            print(f"Could not parse: {current_display}")
            break
        
        current_date = current_year * 12 + current_month
        target_date = year * 12 + month
        
        if target_date > current_date:
            next_btn = page.locator("button.sdp--next-month-btn")
            next_btn.click()
        elif target_date < current_date:
            prev_btn = page.locator("button.sdp--previous-month-btn")
            prev_btn.click()
        else:
            break
        
        page.wait_for_timeout(300)
    
    day_buttons = page.locator(f"button.sdp--date:not(.sdp--disabled):has-text('{day}')")
    
    count = day_buttons.count()
    if count > 1:
        for i in range(count):
            btn = day_buttons.nth(i)
            classes = btn.get_attribute("class")
            if "disabled" not in classes and "text-muted" not in classes:
                btn.click()
                break
    else:
        day_buttons.first.click()
    
    page.wait_for_timeout(500)
    
    hidden_input = page.locator("#datehidden-ProposedSurgeryTreatment")
    actual = hidden_input.input_value()
    
    print(f"✓ Date set to: {actual}")
    assert actual == date_string, f"Expected {date_string}, got {actual}"
    

    page.wait_for_timeout(2000)
    save_btn = page.get_by_role("button", name="SAVE").first
    save_btn.wait_for(state="visible")
    save_btn.click()
    


def process_treatment(page, details):
    treatment_btn = page.locator("button:has(h4:text('Treatment'))").first
    treatment_btn.wait_for(state="visible", timeout=10000)
    treatment_btn.scroll_into_view_if_needed()
    treatment_btn.click()
    
    page.wait_for_timeout(1000)
    
    print("Opening Diagnosis section...")
    page.wait_for_timeout(2000)

    # Find the Diagnosis toggle button using the text label in the same row
    diagnosis_toggle = page.locator(
        "xpath=//div[contains(@class,'zKnrJTgkPSZ9htEXwleT')][contains(text(),'Diagnosis')]"
        "/ancestor::div[contains(@class,'accordion-item')][1]"
        "//button[contains(@class,'accordion-button')]"
    ).first

    diagnosis_toggle.wait_for(state="visible", timeout=10000)
    diagnosis_toggle.scroll_into_view_if_needed()

    # Check if diagnosis accordion is closed (aria-expanded="false")
    is_expanded = diagnosis_toggle.get_attribute("aria-expanded")
    print(f"Diagnosis expanded: {is_expanded}")
    if is_expanded == "false":
        diagnosis_toggle.click()
        page.wait_for_timeout(1000)

    print("Looking for diagnosis input field...")
    diagnosis_input = page.locator("input[placeholder='Enter Diagnosis']").first
    diagnosis_input.wait_for(state="visible", timeout=10000)
    print("Found diagnosis input, typing...")
    diagnosis_input.click()
    diagnosis_input.press("Control+A")
    diagnosis_input.press("Backspace")
    diagnosis_input.type(details['diagnosis'], delay=100)
    print(f"Typed diagnosis: {details['diagnosis']}")
    
    page.wait_for_timeout(500)
    
    first_item = page.locator('li[data-entityindex="0"] .entityTitlePointer')
    first_item.scroll_into_view_if_needed()
    first_item.click(force=True)

    section = page.locator("div.col-md-1").first
    add_btn = section.locator("button:has-text('ADD'):not([disabled])")

    add_btn.wait_for(state="visible")
    add_btn.click()

    page.wait_for_timeout(1000)
    diagnosis_toggle.click()

    page.wait_for_timeout(1000)

    # ========== TREATMENT PLAN SECTION ==========
    print("Opening Treatment Plan section...")
    treatment_plan_toggle = page.locator(
        "xpath=//div[contains(@class,'zKnrJTgkPSZ9htEXwleT')][contains(text(),'Treatment Plan')]"
        "/ancestor::div[contains(@class,'accordion-item')][1]"
        "//button[contains(@class,'accordion-button')]"
    ).first

    treatment_plan_toggle.wait_for(state="visible", timeout=10000)
    treatment_plan_toggle.scroll_into_view_if_needed()

    is_expanded = treatment_plan_toggle.get_attribute("aria-expanded")
    print(f"Treatment Plan expanded: {is_expanded}")
    if is_expanded == "false":
        treatment_plan_toggle.click()
        page.wait_for_timeout(1000)

    # Get procedures and amounts from details
    procedures = details.get('procedures', [])
    amounts = details.get('amounts', [])

    # List of Unspecified Codes to cycle through when procedure not found
    unspecified_codes = [
        "Unspecified Code(US0010U100-None)",
        "Unspecified Code(US001U100-None)",
        "Unspecified Code(US002U100-None)",
        "Unspecified Code(US003U100-None)",
        "Unspecified Code(US004U100-None)",
        "Unspecified Code(US005U100-None)",
        "Unspecified Code(US006U100-None)",
        "Unspecified Code(US007U100-None)",
        "Unspecified Code(US008U100-None)",
        "Unspecified Code(US009U100-None)"
    ]
    unspecified_idx = 0  # Track which unspecified code to use next

    print(f"Adding {len(procedures)} procedure(s)...")

    # Loop through each procedure
    for proc_idx, (procedure_name, amount) in enumerate(zip(procedures, amounts)):
        print(f"\n  Adding Procedure {proc_idx + 1}/{len(procedures)}: {procedure_name[:40]}...")

        # Re-locate the treatment plan section after each add (page reloads)
        treatment_plan_section = page.locator(
            "xpath=//div[contains(@class,'zKnrJTgkPSZ9htEXwleT')][contains(text(),'Treatment Plan')]"
            "/ancestor::div[contains(@class,'accordion-item')][1]"
        ).first

        procedure_dropdown = treatment_plan_section.locator("input[role='combobox']").nth(1)
        procedure_dropdown.wait_for(state="visible", timeout=10000)
        procedure_dropdown.click()
        page.wait_for_timeout(500)

        # Type the procedure name
        procedure_dropdown.type(procedure_name, delay=100)
        page.wait_for_timeout(1500)

        # Check if any option appeared in dropdown
        option_exists = page.locator('[id*="listbox"] [id*="option"]').first.is_visible()

        if option_exists:
            # Select the first matching option
            first_option = page.locator('[id*="listbox"] [id*="option"]').first
            first_option.click()
            page.wait_for_timeout(500)
            print(f"    ✓ Found matching procedure in dropdown")
        else:
            # Procedure not found - use Unspecified Code from list
            current_unspecified = unspecified_codes[unspecified_idx % len(unspecified_codes)]
            print(f"    ⚠ Procedure not found, using {current_unspecified}...")

            # Clear and type the specific Unspecified Code
            procedure_dropdown.press("Control+A")
            procedure_dropdown.press("Backspace")
            procedure_dropdown.type(current_unspecified, delay=100)
            page.wait_for_timeout(1000)

            unspecified_option = page.locator('[id*="listbox"] [id*="option"]').first
            unspecified_option.wait_for(state="visible", timeout=5000)
            unspecified_option.click()
            page.wait_for_timeout(1000)

            # Move to next unspecified code for next time
            unspecified_idx += 1

            # Fill the Procedure Name field
            procedure_name_input = treatment_plan_section.locator('input[placeholder="Type here"]').nth(1)
            procedure_name_input.wait_for(state="visible", timeout=5000)
            procedure_name_input.click()
            procedure_name_input.type(procedure_name, delay=50)
            page.wait_for_timeout(500)

            # Fill the Amount field
            amount_input = treatment_plan_section.locator('input[placeholder="Type here"]').nth(2)
            amount_input.wait_for(state="visible", timeout=5000)
            amount_input.click()
            amount_input.type(str(amount), delay=50)
            page.wait_for_timeout(500)

        # Click the + button to add the procedure
        print(f"    Clicking + button...")
        plus_btn = treatment_plan_section.locator("div.col-lg-1 img.m9FzljqXbDJyFhzambbf").first
        plus_btn.wait_for(state="visible", timeout=10000)
        plus_btn.scroll_into_view_if_needed()
        page.wait_for_timeout(500)
        plus_btn.evaluate("el => el.click()")

        # Wait for page to load after adding procedure
        print(f"    Waiting for page to load...")
        page.wait_for_timeout(3000)
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(2000)

    print(f"\n✓ All {len(procedures)} procedures added!")

    # Close Treatment Plan section
    treatment_plan_toggle = page.locator(
        "xpath=//div[contains(@class,'zKnrJTgkPSZ9htEXwleT')][contains(text(),'Treatment Plan')]"
        "/ancestor::div[contains(@class,'accordion-item')][1]"
        "//button[contains(@class,'accordion-button')]"
    ).first
    treatment_plan_toggle.scroll_into_view_if_needed()
    page.wait_for_timeout(500)
    treatment_plan_toggle.click(force=True)
    page.wait_for_timeout(1000)

    # ========== INVESTIGATIONS/DOCUMENTS SECTION ==========
    print("Opening Investigations/Documents section...")

    # Scroll up to find the section
    page.evaluate("window.scrollBy(0, -500)")
    page.wait_for_timeout(1000)

    investigations_toggle = page.locator(
        "xpath=//div[contains(@class,'zKnrJTgkPSZ9htEXwleT')][contains(text(),'Investigations')]"
        "/ancestor::div[contains(@class,'accordion-item')][1]"
        "//h2[contains(@class,'accordion-header')]//button"
    ).first

    investigations_toggle.wait_for(state="visible", timeout=15000)
    investigations_toggle.scroll_into_view_if_needed()

    is_expanded = investigations_toggle.get_attribute("aria-expanded")
    print(f"Investigations/Documents expanded: {is_expanded}")
    if is_expanded == "false":
        investigations_toggle.click()
        page.wait_for_timeout(2000)

    # Upload files by document name using paths from details
    card_file = details.get('card_file', '')
    patient_photo = details.get('patient_photo', '')

    file_uploads = {
        "CGHS Card Copy": card_file,
        "Patient Photo": patient_photo
    }

    for doc_name, doc_file_path in file_uploads.items():
        if doc_file_path and os.path.exists(doc_file_path):
            print(f"Uploading {doc_name}...")
            row = page.locator('tr').filter(has_text=doc_name)
            file_input = row.locator('input[type="file"]')
            file_input.set_input_files(doc_file_path)
            print(f"✓ {doc_name}: uploaded")
            page.wait_for_timeout(1000)
        else:
            print(f"⚠ {doc_name}: File not found or path empty - {doc_file_path}")

    print("Files upload completed!")

    # Keep Investigations/Documents section OPEN for manual verification
    input("\n⏸️  Verify file uploads are complete, then press Enter to continue to Care Team Details...")

    # Close Investigations/Documents section after verification
    investigations_toggle.click()
    page.wait_for_timeout(1000)

    # ========== CARE TEAM DETAILS SECTION ==========
    print("Opening Care Team Details section...")

    # Find the Care Team accordion section container
    care_team_section = page.locator(
        "xpath=//div[contains(@class,'zKnrJTgkPSZ9htEXwleT')][contains(text(),'Care Team Details')]"
        "/ancestor::div[contains(@class,'accordion-item')][1]"
    ).first

    # The toggle button is in the h2.accordion-header
    care_team_button = care_team_section.locator("h2.accordion-header button").first

    care_team_button.wait_for(state="visible", timeout=10000)
    care_team_button.scroll_into_view_if_needed()

    is_expanded = care_team_button.get_attribute("aria-expanded")
    print(f"Care Team Details expanded: {is_expanded}")

    if is_expanded == "false":
        care_team_button.click()
        page.wait_for_timeout(1000)

    page.wait_for_timeout(1000)

    # Select "Other" from Doctor Type dropdown (in top row with class ktN2z_0w5Pjb2qqvloRq)
    print("Selecting Doctor Type: Other...")
    top_row = care_team_section.locator("div.row.ktN2z_0w5Pjb2qqvloRq").first
    doctor_type_dropdown = top_row.locator("input[role='combobox']").first
    doctor_type_dropdown.wait_for(state="visible", timeout=10000)
    doctor_type_dropdown.click()
    page.wait_for_timeout(500)

    doctor_type_dropdown.type("Other", delay=100)
    page.wait_for_timeout(500)

    # Select from dropdown options
    other_option = page.locator('[id*="listbox"] [id*="option"]').first
    other_option.wait_for(state="visible", timeout=5000)
    other_option.click()
    page.wait_for_timeout(1000)

    # After selecting "Other", the form fields appear
    # Use specific IDs/placeholders for the inputs
    print("Filling doctor details...")
    doctor_name_input = page.locator('input[id="doctor name"]').first
    doctor_name_input.wait_for(state="visible", timeout=10000)
    doctor_name_input.type(details['doctor_name'], delay=50)

    reg_id_input = page.locator('input[id="registration no."]').first
    reg_id_input.wait_for(state="visible", timeout=10000)
    reg_id_input.type(details['reg_id'], delay=50)

    # Qualification dropdown - use the container ID
    print("Selecting Qualification dropdown...")
    qualification_container = page.locator('#qualification').first
    qualification_container.wait_for(state="visible", timeout=10000)
    qualification_container.click()
    page.wait_for_timeout(500)

    page.keyboard.type("M.B.B.S", delay=100)
    page.wait_for_timeout(500)

    mbbs_option = page.locator('[id*="listbox"] [id*="option"]').first
    mbbs_option.wait_for(state="visible", timeout=5000)
    mbbs_option.click()
    page.wait_for_timeout(500)

    contact_input = page.locator('input[id="contact no."]').first
    contact_input.wait_for(state="visible", timeout=10000)
    contact_input.type("6789067890", delay=50)

    # Click ADD button (should be enabled after filling all fields)
    print("Clicking ADD button...")
    page.wait_for_timeout(500)
    add_btn = care_team_section.locator("button:has-text('ADD'):not([disabled])").first
    add_btn.wait_for(state="visible", timeout=10000)
    add_btn.click()
    page.wait_for_timeout(1000)

    # Close Care Team section
    care_team_button.click()
    page.wait_for_timeout(1000)
    print("Care Team Details completed!")

    # ========== PREVIEW & VALIDATE ==========
    print("Clicking Preview & Validate button...")
    preview_btn = page.locator("button:has-text('Preview & Validate')").first
    preview_btn.wait_for(state="visible", timeout=10000)
    preview_btn.scroll_into_view_if_needed()
    preview_btn.click()
    page.wait_for_timeout(2000)
    print("Preview & Validate clicked!")

    # Wait for the modal to appear and click the validate button
    print("Waiting for modal to appear...")
    page.wait_for_timeout(3000)

    # Find Validate button using XPath
    print("  Looking for Validate button...")
    validate_btn = page.locator("xpath=/html/body/div[6]/div/div/div[3]/div/div[1]/button")
    validate_btn.wait_for(state="visible", timeout=15000)
    validate_btn.click(force=True)
    page.wait_for_timeout(3000)
    print("Validate button clicked!")

    # Click "Initiate Pre-Authorization" button
    print("  Looking for Initiate Pre-Authorization button...")
    initiate_btn = page.locator("button:has-text('Initiate Pre-Authorization')").first
    initiate_btn.wait_for(state="visible", timeout=15000)
    initiate_btn.click(force=True)
    page.wait_for_timeout(2000)
    print("Initiate Pre-Authorization clicked!")

    # Click "YES" in confirmation modal
    print("  Looking for YES confirmation button...")
    yes_btn = page.locator("xpath=/html/body/div[8]/div/div/div[2]/div/div[1]/button")
    yes_btn.wait_for(state="visible", timeout=15000)
    yes_btn.click(force=True)
    page.wait_for_timeout(3000)
    print("✅ Pre-Authorization submitted successfully!")

    # Navigate back to home for next entry
    print("  Clicking Home button...")
    page.evaluate("window.scrollTo(0, 0)")
    page.wait_for_timeout(1000)
    home_btn = page.locator("xpath=//*[@id='noun-home-5510410']/ancestor::svg")
    home_btn.wait_for(state="visible", timeout=10000)
    home_btn.click(force=True)
    page.wait_for_timeout(3000)
    page.wait_for_load_state("networkidle")
    page.wait_for_timeout(2000)
    print("  ✓ Ready for next entry")


if __name__ == "__main__":
    run_automation()