from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from datetime import datetime
import os
import pandas as pd
import re
from openpyxl.styles import PatternFill, NamedStyle, Alignment, Font, Border, Side
from openpyxl.utils.cell import get_column_letter

# Script definitions
PREVIOUS_PATCH_WEDNESDAY = "2024-08-14"

# URLs used for Microsoft Catalog patch gathering

catalog_urls = [
    ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+Server+2016+for+x64","Windows Server 2016","version 1607"),
    ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+Server+2019+for+x64","Windows Server 2019","version 1809"),
    ("https://www.catalog.update.microsoft.com/Search.aspx?q=Microsoft+Server+Operating+System-21H2+for+x64","Windows Server 2022","version 21H2")
]

# catalog_urls = [
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+Server+2016+for+x64+2024","Windows Server 2016 (2024)","version 1607"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+Server+2019+for+x64+2024","Windows Server 2019 (2024)","version 1809"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+Server+2016+for+x64","Windows Server 2016","version 1607"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+Server+2019+for+x64","Windows Server 2019","version 1809"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Microsoft+Server+Operating+System-21H2+for+x64","Windows Server 2022","version 21H2"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+10+Version+22H2+for+x64","Windows 10","version 22H2"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Windows+11+Version+23H2+for+x64","Windows 11","version 23H2"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=SQL+Server+2016+Service+Pack+3","SQL Server 2016","SP3 (13.0.6419.1)"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=SQL+Server+2019","SQL Server 2019","version RTM"),
#     ("https://www.catalog.update.microsoft.com/Search.aspx?q=Microsoft+Office+2016+32-bit","Office 2016","x32 bits")
# ] 

def catalog_scrape(urls,cutoff_date,system1_csv,patchstatus_csv, header_mappings=None, remove_columns=None, filter_terms=None, screenshot=True):

    cutoff_date = datetime.strptime(cutoff_date, '%Y-%m-%d')
    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service)

    sys1AllRows = []
    system1_exists = os.path.isfile(system1_csv)
    patchstatus_exists = os.path.isfile(patchstatus_csv)
    try:
        for url, label, version in urls:

            # Go to catalog search
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "ctl00_catalogBody_updateMatches"))
            )
            # Find and click Last Updated to sort from newest to oldest
            driver.find_element(By.ID,"ctl00_catalogBody_updateMatches_ctl02_dateHeaderLink").click()
            
            # Take and save screenshots
            if screenshot: driver.save_screenshot(f'{label}_{datetime.now().strftime('%Y_%m_%d')}.png')

            #Locate table of patches
            table = driver.find_element(By.ID, "ctl00_catalogBody_updateMatches")
            rows = table.find_elements(By.TAG_NAME, "tr")

            # Extract the header row
            header = [th.text.strip() for th in rows[0].find_elements(By.TAG_NAME, 'th')]

            # Add the product label column to the header
            header[0] = '#' 

            # Mapping column indices based on headers
            column_indices = {name: index for index, name in enumerate(header)}
            # Adjust header based on mappings provided
            if header_mappings:
                # Swap and reorder headers as per header_mappings
                for key, value in header_mappings.items():
                    if key in column_indices and value in column_indices:
                        header[column_indices[key]], header[column_indices[value]] = header[column_indices[value]], header[column_indices[key]]
            
            # Remove unnecessary columns (download, size, etc.)
            if remove_columns:
               header = [col for col in header if col not in remove_columns]

            # List to store the System 1 and Patch Status data
            system1_rows = []
            patchstatus_rows = []
            patchNumber = 1
            # Process each row and filter based on the date
            for row in rows[1:]: # Skip the header
                cells = row.find_elements(By.TAG_NAME, 'td')
                if len(cells) > 0:

                    title_text = cells[column_indices['Title']].text # Patch name / description

                    # Extract the date
                    date_str = cells[column_indices['Last Updated']].text
                    try:
                        row_date = datetime.strptime(date_str,'%m/%d/%Y')
                        if row_date >= cutoff_date:
                            #print(title_text)
                            try:
                                kb_number = re.search(r"KB\d{6,}", title_text).group()
                            except ValueError:
                                print(f"Invalid KB number in row {title_text}")
                                kb_number = "KB error"
                            #print(kb_number.group())
                            #Create a row with the new order excluding unwanted columns
                            new_row = [patchNumber] + [cells[column_indices[col]].text for col in header[1:] if col in column_indices] + ["Microsoft"] + [label] + [version] + [f"{kb_number}"] + [f"https://support.microsoft.com/kb/{kb_number}"]
                            patchNumber += 1
                            # Check if it's System 1 row or Patch Status row
                            if filter_terms and any(term.lower() in title_text.lower() for term in filter_terms):
                                # Add row to Patch Status rows
                                patchstatus_rows.append(new_row)
                                patchNumber -= 1
                            else:
                                # Add row to System 1 rows
                                system1_rows.append(new_row)
                                sys1AllRows.append(new_row)
                    except ValueError:
                            print(f"Invalid date format found in row {title_text} {date_str}")
            
            header.extend(['Vendor Name','Vendor Product','Model/Version', 'Patch Name', 'Patch Link'])
            header[1] = 'Patch Description'
            header[2] = 'Release Date'
            header[3] = 'Update Type'



            # Write the System 1 rows to the CSV
            # with open(system1_csv, 'a',newline='',encoding='utf-8') as csvfile:
            #     csvwriter = csv.writer(csvfile)
            #     # Write the header only if the file does not already exist
            #     if not system1_exists:
            #         csvwriter.writerow(header)
            #         system1_exists = True # Ensure headers are not written again
            #     # Write the rows
            #     csvwriter.writerows(system1_rows)

            # # Write the Patch Status rows to the CSV
            # with open(patchstatus_csv, 'a',newline='',encoding='utf-8') as csvfile:
            #     csvwriter = csv.writer(csvfile)
            #     # Write the header only if the file does not already exist
            #     if not patchstatus_exists:
            #         csvwriter.writerow(header)
            #         patchstatus_exists = True # Ensure headers are not written again
            #     # Write the rows
            #     csvwriter.writerows(patchstatus_rows)
            # print(f"Filtered data has been written to {system1_csv} and {patchstatus_csv}")
    finally:
        sys1df = pd.DataFrame(data=sys1AllRows,columns=header)
        driver.quit()
        return sys1df


# URLs used for AV definitions gathering
av_urls = [
    ("https://www.trellix.com/downloads/security-updates",
     "McAfee","McAfee Endpoint Security 10.7","AMCore Content for McAfee Security (V3 DAT)",
     "//*[starts-with(.,'V3_')]"),
    ("https://www.broadcom.com/support/security-center/definitions/download/detail?gid=sep14#sep-14.3-to-14.3-ru8-x64",
     "Broadcom","Symantec 14.3 MP1","Symantec Endpoint Protection Client Installations on Windows Platforms (64-bit) (SEP 14.3 Dark-Network Client only)",
     "//*[starts-with(.,'core15sdssepv5i64')]"),
    ("https://www.microsoft.com/en-us/wdsi/defenderupdates",
     "Microsoft","Windows Defender","Security intelligence updates for Microsoft Defender Antivirus and other Microsoft antimalware",
     "//*[starts-with(.,'Version: 1.')]")
]


def av_scrape():
    service = Service(executable_path="chromedriver.exe")
    driver = webdriver.Chrome(service=service)
    headers = ["#", "Patch Description", "Release Date", "Update Type", "Vendor Name", "Vendor Product", "Model/Version"]
    av_rows = []
    try:
        for url,vendor, product, description, searchfor in av_urls:
            # Go to AV defintion site
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH(searchfor)))
            )
          
            new_row = ["1"] + [description] + ["todays date"] + ["Virus Def."] + [vendor] + [product] + ["ver"]
            av_rows.append(new_row)
    finally:
        avdf = pd.DataFrame(data=av_rows, columns=headers)
        driver.quit()
        print(avdf)
        #return avdf

def build_excel():
    # Get Catalog Dataframe
    catalogDf = catalog_scrape(
    urls=catalog_urls,
    cutoff_date=PREVIOUS_PATCH_WEDNESDAY,
    system1_csv=f"system1_{datetime.now().strftime('%Y_%m_%d')}.csv",
    patchstatus_csv=f"patch_status_{datetime.now().strftime('%Y_%m_%d')}.csv",
    header_mappings={'Classification':'Last Updated'}, # Swap these columns
    remove_columns=['Products','Version','Size','Download'], # Remove these columns
    filter_terms=[
        'Preview','Dynamic','Azure Stack HCI','Office 2019',
        'Access','Project','Outlook','PowerPoint','Visio','Publisher'
        '3.5, 4.8 and 4.8.1','3.5, 4.7.2 and 4.8',
        '.NET Framework 3.5 and 4.8.1 for Windows 10 Version 22H2 for x64'
        ], # Filter rows based on patch criteria
    screenshot=False)

    masterDf = catalogDf # This is temporary, will have to do .concat() with AV dataframe at some point

    # Patch name (KB number), Patch link (KB with link), Approved by BN (initially empty), Test Status (intially empty), Comment (also empty)
    header_list = ["#","Vendor Name","Vendor Product","Model/Version","Patch Name","Patch Description","Patch Link", "Release Date", "Update Type", "Approved by BN", "Test Status", "Comment"] 
    #header_list = ["#", "Patch Link", "Patch Name", "Vendor Name", "Vendor Product", "Model/Version", "Patch Description", "Release Date", "Update Type", "Approved by BN", "Test Status", "Comment"] 
    masterDf = masterDf.reindex(columns = header_list)
    # masterDf.style.set_properties(**{'border': '1px solid black'})

    with pd.ExcelWriter("System 1 Patches.xlsx", engine="openpyxl") as writer:
        masterDf.to_excel(writer, sheet_name="September-2024", index=False, freeze_panes=(2,0))
        # Workbook
        workbook = writer.book
        # Worksheet
        worksheet = workbook["September-2024"]

        # AutoFit column width
        for column in worksheet.columns:
            max_length = max(len(str(cell.value)) for cell in column)
            adjusted_width = (max_length + 2) * 1.0
            worksheet.column_dimensions[column[0].column_letter].width = adjusted_width
        # Adjust the size of column "A"
        worksheet.column_dimensions["F"].width = 45.0

        for row in worksheet.iter_rows(min_row=2, max_row=2, min_col=2, max_col=worksheet.max_column):
            for cell in row:
                cell.fill = PatternFill(fill_type="solid", start_color="DDEBF7", end_color="DDEBF7")
                cell.font = Font(color="000000", size=8, bold=True)
                cell.alignment = Alignment(horizontal="center",vertical="center")

build_excel()
