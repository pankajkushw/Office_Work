def WorkinProgress_Report():
    w_FY_type = ["2016-2017", "2017-2018", "2018-2019", "2019-2020", "2020-2021", "2022-2023", "2025-2026"]
    #Manual intervention required for clicking download button due to timeout issues.
    #w_FY_type = ["2024-2025"]
    print("Navigating to Work in Progress Report...")
    mySleepFunction(WAIT_SECONDS)
    # Selcting A2 Report
    try: 
        link = driver.find_element(By.PARTIAL_LINK_TEXT, "Work Progress for PMAY-G")
        link.click()
    except Exception as e:
        print("Error while navigating to Work in Progress Report:", e)
        driver.quit()
        return
    finally:
        print("Navigation to Work in Progress Report attempted.")

    print("Work in Progress Report Selected")
    for fy in data_type:
        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlyear"))
        dropdown.select_by_visible_text(fy)
        print(fy)
        
        for scheme in w_Scheme_type:
            dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_rblScheme"))
            dropdown.select_by_visible_text(scheme)
            print(scheme)
            
            for fy in w_FY_type:
                dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlFinYear"))
                dropdown.select_by_visible_text(fy)
                print(fy)
                for block in block_type:
                    dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlBlock"))
                    dropdown.select_by_visible_text(block)
                    print(block)
                    for panchayat in Panchayat_type:
                        dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlPanch"))
                        dropdown.select_by_visible_text(panchayat)
                        print(panchayat)
                        for category_type in Category_type:
                            dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlCat"))
                            dropdown.select_by_visible_text(category_type)
                            print(category_type)
                            for progress_type in Progress_type:
                                dropdown = Select(driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_ddlSanction"))
                                dropdown.select_by_visible_text(progress_type)
                                print(progress_type)
                                #download block data for
                                button = driver.find_element(By.ID, "ctl00_ContentPlaceHolder1_btnExportexcel")
                                button.click()
                                mySleepFunction(WAIT_SECONDS)
                                new_file_name = f"WorkinProgress_{block}_{fy}"
                                rename_latest_download(new_file_name)  
                                mySleepFunction(2)                             

                                # # NEW: Dynamically wait instead of using a static sleep timer
                                # if wait_for_download_complete(DOWNLOAD_DIR, timeout=15):
                                #     new_file_name = f"WorkinProgress_{block}_{fy}"
                                #     rename_latest_download(new_file_name)
                                # else:
                                #     print(f"Skipping rename for {block} due to download timeout.")
                                print("Data Downloaded for " + fy + " " + scheme + " " + fy + " " + block + " " + panchayat + " " + category_type + " " + progress_type)

def rename_latest_download(new_filename):
    # 1. Automatically locate the default user Downloads folder
    download_dir = Path.home() / "Downloads"
    move_dir = Path.home() / "Downloads/WIP"  # Adjust this path if your downloads go to a different folder
    
    # 2. Gather all items in the directory that are files
    files = [f for f in download_dir.iterdir() if f.is_file()]
    
    if not files:
        print("No files found in the Downloads folder.")
        return

    # 3. Find the most recently modified file
    latest_file = max(files, key=os.path.getmtime)
    
    # Keep the original file extension (e.g., .pdf, .zip, .csv)
    file_extension = latest_file.suffix
    
    # 4. Construct the new path
    new_file_path = move_dir / f"{new_filename}{file_extension}"
    
    # 5. Rename the file safely
    try:
        if not new_file_path.exists():
            latest_file.rename(new_file_path)
            print(f"Successfully renamed:\nFrom: '{latest_file.name}'\nTo:   '{new_file_path.name}'")
        else:
            print(f"Error: A file named '{new_file_path.name}' already exists.")
    except Exception as e:
        print(f"An error occurred: {e}")



login()
mySleepFunction(WAIT_SECONDS)
driver.quit()
