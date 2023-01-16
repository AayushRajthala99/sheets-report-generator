from googleapiclient.discovery import build  # for building service
# to work with service service_account
from google.oauth2 import service_account
from service_creator import Create_Service, Create_Service2
from googleapiclient.errors import HttpError
from datetime import datetime

# downloaded from console and stored in the working discovery
# SERVICE_ACCOUNT_FILE = 'service-account.json'
SERVICE_ACCOUNT_FILE = '.\credentials\service-account.json'
DRIVE_SCOPE = ['https://www.googleapis.com/auth/drive']
SHEET_SCOPE = ['https://www.googleapis.com/auth/spreadsheets']
SHEETS_API_NAME = 'sheets'
SHEETS_API_VERSION = 'v4'
DRIVE_API_NAME = 'drive'
DRIVE_API_VERSION = 'v3'

# creds = service_account.Credentials.from_service_account_file(
#     SERVICE_ACCOUNT_FILE, scopes=SHEET_SCOPE)

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SHEET_SCOPE)


def writer(SPREADSHEET_ID, name, headers, result):
    try:
        sheetservice = build(
            SHEETS_API_NAME, SHEETS_API_VERSION, credentials=creds)
        sheet = sheetservice.spreadsheets()  # just making it short

        # --------------writing values to the cells of Sheet2
        sheetValues = []
        # timestamp = f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

        # Dynamic Assignment of Spreadsheet Column Range...
        headerLength = len(headers)
        rangeValue = f"Sheet1!A1:A{headerLength}"

        # Formatting Sheets Headers & Titles...
        # testNameValue = ['Test Name :', sheet_title, '',
        #  'Report Generation Time', timestamp]
        # lineBreakValue = []

        # for x in range(0, headerLength):
        #     if (x < (headerLength-5)):
        #         testNameValue.append('')

        #     lineBreakValue.append('----------')

        # sheetValues.append(testNameValue)
        # sheetValues.append(headers)

        # Creating A List containing only Values from A List of Dictionaries...
        # resultValues = [list(result.values()) for result in resultArray]
        resultValues = [list(result.values())]

        # Concatenate Operation between Two Lists...
        sheetValues = sheetValues + resultValues

        print("\nPopulating Google Spreadsheet for ::: [ {0} ]... OK!".format(
            name))

        # write data values into the cells starting from A1 of Sheet1
        sheet.values().append(spreadsheetId=SPREADSHEET_ID, range=f"{rangeValue}",
                              valueInputOption="RAW", body={"values": sheetValues}).execute()
        return

    except HttpError as error:
        print("Error in Sheet Writer Code::: ", error)
        pass


def sharer(ID):
    try:
        # share-link code starts here.
        file_id = ID

        # setting the permission attributes
        request_body = {
            'role': 'reader',
            'type': 'anyone'
        }
        # driveservice = Create_Service2(
        #     SERVICE_ACCOUNT_FILE, DRIVE_API_NAME, DRIVE_API_VERSION, DRIVE_SCOPE)

        driveservice = Create_Service2(
            SERVICE_ACCOUNT_FILE, DRIVE_API_NAME, DRIVE_API_VERSION, DRIVE_SCOPE)

        response_permission = driveservice.permissions().create(
            fileId=file_id, body=request_body, supportsAllDrives=True).execute()
        # print(response_permission)

        # Print Sharing Link
        response_share_link = driveservice.files().get(
            fileId=file_id,
            fields='webViewLink',
            supportsAllDrives=True).execute()

        # print("The share link is : ",response_share_link)
        # print("\nGenerating the URL... OK!")

        # getting just the link from the object response_share_link
        link = response_share_link.get("webViewLink")
        return link
    except Exception as error:
        print("Error in Link Generation [Sharer] Code::: ", error)
        pass
