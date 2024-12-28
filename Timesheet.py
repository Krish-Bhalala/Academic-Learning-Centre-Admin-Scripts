import os
import openpyxl
import http.client
from datetime import datetime, timedelta
import ssl
import urllib.parse
import gzip
import io
from bs4 import BeautifulSoup
from email_generator import emailDataGenerator  # Importing the function from email_generator.py

ssl._create_default_https_context = ssl._create_unverified_context

directory_path = r"PATH TO FOLDER WITH TIMESHEET.XLXS"
file_path = r"PATH TO FILE TIMESHEET.XLXS"
TUTOR_ID = <ADD YOUR TUTOR ID HERE>
EMPLOYEE_ENDPOINT = <ADD EMPLOYEE ENDPOINT FROM INSPECTOR TOOLS>

FIRSTROW = 15
startDate = ""
endDate = ""
COLUMN_SEQUENCE = "DEFGIK"

def fetchRawData(tutor_ID,appointment_type,start_Date,end_Date):
    start_Date = datetime.strptime(start_Date,"%d-%b-%y")
    start_Date = start_Date.strftime("%Y-%m-%d")
    end_Date = datetime.strptime(end_Date,"%d-%b-%y")
    end_Date = end_Date.strftime("%Y-%m-%d")
    print(start_Date,end_Date)
    conn = http.client.HTTPSConnection("manitoba.mywconline.com")
    payload = f'rid%5B%5D={tutor_ID}&sdate={start_Date}&edate={end_Date}'

    # REMOVED COOKIES AND HID (EMPLOYEE ONLY) ENDPOINTS FOR PRIVACY
    headers = {
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/png,image/svg+xml,*/*;q=0.8',
      'Accept-Language': 'en-US,en;q=0.5',
      'Accept-Encoding': 'gzip, deflate, br, zstd',
      'Content-Type': 'application/x-www-form-urlencoded',
      'Origin': 'https://manitoba.mywconline.com',
      'DNT': '1',
      'Sec-GPC': '1',
      'Connection': 'keep-alive',
      'Referer': f'https://manitoba.mywconline.com/{EMPLOYEE_ENDPOINT}?reset=1&sid={TUTOR_ID}',
      'Upgrade-Insecure-Requests': '1',
      'Sec-Fetch-Dest': 'document',
      'Sec-Fetch-Mode': 'navigate',
      'Sec-Fetch-Site': 'same-origin',
      'Sec-Fetch-User': '?1',
      'Priority': 'u=0, i'
    }
    conn.request("POST",EMPLOYEE_ENDPOINT, payload, headers)
    res = conn.getresponse()
    data = res.read()
    # Check if the response is gzip-encoded
    if res.getheader('Content-Encoding') == 'gzip':
        # Decompress gzip data
        buffer = io.BytesIO(data)
        with gzip.GzipFile(fileobj=buffer) as f:
            decompressed_data = f.read()
    else:
        decompressed_data = data

    try:
        # Try to decode as UTF-8
        decoded_data = decompressed_data.decode('utf-8')

        # Parse the HTML content
        soupData = BeautifulSoup(decoded_data, 'html.parser')
        return soupData
    except UnicodeDecodeError:
        # If UTF-8 decoding fails, print the raw data for inspection
        print("Failed to decode as UTF-8. Raw data:")

def parseRawData(soup):
    names = soup.find_all("p",class_="modal_cursor fw-bold theblue fs-5 m-0")
    names = [[name.next_element] for name in names]
    dates = soup.find_all("span",class_="fw-bold fs-5")
    times = [date.parent.br.next_element for date in dates]
    times = [time.split(" to ") for time in times]

    for i in range(len(dates)):
        temp = dates[i].next_element
        dates[i] = temp.split(", ",maxsplit=1)

    data = format_data(names,times,dates)
    return data

def format_data(names,times,dates):
    data = []
    for name,time,date in zip(names,times,dates):
        print(name,time,date)
        data.append(date[:1]+time+[name[0]+" dated: "+date[1]])
    return data

def editRow(sheet,rowNumber,rowData):
    for i in range(len(COLUMN_SEQUENCE)):
        cellID = COLUMN_SEQUENCE[i] + str(rowNumber)
        sheet[cellID] = rowData[i]

def updateDate():
    global startDate
    global endDate
    # Get the current date and time
    today = datetime.now()

    # Calculate the number of days to subtract to get to the previous Saturday
    days_to_subtract = (today.weekday() + 2) % 7
    previous_saturday = today - timedelta(days=days_to_subtract)
    following_friday = previous_saturday + timedelta(days=6)

    # Format the date as yymmdd
    formatted_date = previous_saturday.strftime("%d-%b-%y")
    startDate = formatted_date
    endDate = following_friday.strftime("%d-%b-%y")

def fillDate(sheet):
    START_DATE_CELLID = "F3"        #2024-06-08 date format
    global startDate
    global endDate
    sheet[START_DATE_CELLID] = startDate

def saveWorkbook(workbook,newFileName):
    new_filename = newFileName
    if os.path.exists(new_filename):
        print(f"Warning: {new_filename} already exists. Choose a different name.")
    else:
        workbook.save(directory_path+new_filename)
        print(f"Workbook saved as {new_filename}")

def editWorkBook(data,sheet,workbook):
    transactionID = "1030"
    description = "tutoring"
    currRow = FIRSTROW
    for element in data:
        rows = element[:1]+[transactionID]+element[1:3]+[description]+[element[3]]
        editRow(sheet,currRow,rows)
        currRow += 1

    fillDate(sheet)
    print("Excel file has been edited successfully.")

def initiateWorkBook():
    try:
        # Load the existing Excel file
        workbook = openpyxl.load_workbook(file_path)
        return workbook
    except FileNotFoundError:
        print(f"Error: The file at {file_path} was not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

def initiateSheet(workbook):
    try:
        return workbook.active       # Select the active sheet (or you can select a specific sheet by name)
    except Exception as e:
        print(e)

def main():
    workbook = initiateWorkBook()
    sheet = initiateSheet(workbook)
    updateDate()
    data = fetchRawData(TUTOR_ID,"canceled",startDate,endDate)

    parsed_data = parseRawData(data)
    editWorkBook(parsed_data,sheet,workbook)
    newFileName = f"Krish Bhalala Timesheetssss {startDate} to {endDate}.xlsx"
    # Save the changes to the Excel file

main()
