import io
import os
import csv
import json
import time
import numpy as np
from os import path
import pandas as pd
from sheet_writer import sharer, writer, shareUrlParse
from service_creator import Create_Service, Create_Service2, gspreadService
from gspread_dataframe import set_with_dataframe

# --------------------------------------------------------------------------------------------

CLIENT_SECRET_FILE = os.path.normpath('.\dependencies\client-secret.json')
SERVICE_ACCOUNT_FILE = os.path.normpath('.\dependencies\service-account.json')
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

# subDirectories = ['Enter the directories to be listed in the report']

# Example of subDirectories...
subDirectories = ['recordings', 'responses',
                  'screenshots', 'packetcaptures', 'urlInfo.csv']

# Drive ID of the Shared Drive... Change this to choose the Directory to upload the files to...
driveID = '--Enter The Drive ID Here--'

# Parent Directory ID of the shared drive....
parent_directory_id = '--Enter The Parent Directory ID Here--'

# Spreadsheet ID to populate the results...
spreadsheet_id = '--Enter The Spreadsheet ID Here--'

# ----------------------------------------------------------------------------------------------------------------------------------

# service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
service = Create_Service2(SERVICE_ACCOUNT_FILE, API_NAME, API_VERSION, SCOPES)


def idreturner(id, df, column):
    try:
        if f"{id}" in df[f"{column}"].values:
            index = df[f"{column}"][df[f"{column}"] == f"{id}"].index[0]
            return (df.loc[index, 'id'])
        else:
            return None
    except Exception as error:
        print("Error in ID Returner Function::: ", error)
        pass


def lister(service, driveId, parentId, sleeptime=None):

    if (sleeptime):
        time.sleep(int(sleeptime))

    try:
        query = (f"parents ='{parentId}' and trashed=False")

        response = service.files().list(q=query, orderBy='name', fields="files(id, name, mimeType)", includeItemsFromAllDrives=True,
                                        corpora='drive', driveId=driveId, supportsAllDrives=True, pageSize=500).execute()

        files = response.get('files')
        df = pd.DataFrame(files)

        # Replacing all NaN values with 'N/A'...
        df = df.replace(np.nan, 'N/A')

        df = df.to_dict('records')
        return (df)

    except Exception as error:
        print("Error in lister Function::: ", error)
        pass


def directoryInfo(directoryObject, directoryName):
    try:
        value = [obj for obj in directoryObject if directoryName in obj['name']]

        if (len(value) > 0):
            return value[0]
        else:
            return None

    except Exception as error:
        print("Error in directoryInfo Function::: ", error)
        pass


def subDirectoryInfo(parentDirectory):
    try:
        test_name = parentDirectory['name']
        test_folder_id = parentDirectory['id']
        # test_name = directory['name']
        # print("\nTest Name:::", test_name)

        # Getting Sub-Directories & Files within Test Results [recordings,responses,screenshots,packetcaptures,urlInfo.csv]
        test_subDirectories = lister(service, driveID, test_folder_id, 1)

        recordingInfo = directoryInfo(test_subDirectories, subDirectories[0])
        if (recordingInfo):
            recordingInfo = recordingInfo['id']
        # print("Recording Directory ID:::", recordingInfo)

        responseInfo = directoryInfo(test_subDirectories, subDirectories[1])
        if (responseInfo):
            responseInfo = responseInfo['id']
        # print("Response Directory ID:::", responseInfo)

        screenshotInfo = directoryInfo(test_subDirectories, subDirectories[2])
        if (screenshotInfo):
            screenshotInfo = screenshotInfo['id']
        # print("Screenshot Directory ID:::", screenshotInfo)

        pcapInfo = directoryInfo(test_subDirectories, subDirectories[3])
        if (pcapInfo):
            pcapInfo = pcapInfo['id']
        # print("Packet Capture Directory ID:::", pcapInfo)

        urlInfo = directoryInfo(test_subDirectories, subDirectories[4])
        if (urlInfo):
            urlInfo = urlInfo['id']
        # print("Packet Capture Directory ID:::", urlInfo)

        resultObject = {
            "test_name": test_name,
            "test_folder_id": test_folder_id,
            "recordingFolderID": recordingInfo,
            "responseFolderID": responseInfo,
            "screenshotFolderID": screenshotInfo,
            "pcapFolderID": pcapInfo,
            "urlInfoID": urlInfo
        }

        return resultObject

    except Exception as error:
        print("Error in subDirectoryInfo Function::: ", error)
        pass


resultDirectories = lister(service, driveID, parent_directory_id)

if path.isfile(os.path.normpath('.\dependencies\generatedReports.txt')) is False:
    print("--generatedReports.txt File Created!--")
    File = open(os.path.normpath('.\dependencies\generatedReports.txt'), "x")
    File.close()

reportFile = open(os.path.normpath(
    '.\dependencies\generatedReports.txt'), 'rb')
alreadyGeneratedReportList = reportFile.read().splitlines()
reportFile.close()

for i, value in enumerate(alreadyGeneratedReportList):
    if (len(value) > 0):
        alreadyGeneratedReportList[i] = value.decode('ASCII')

reportGenerationCount = 0
generatedReports = []

# GSPREAD CLIENT INITIALIZATION...
gspreadClient = gspreadService(SERVICE_ACCOUNT_FILE)
spreadsheetObject = gspreadClient.open_by_key(spreadsheet_id)

for directory in resultDirectories:
    try:
        directoryID = str(directory["id"])
        if (not (directoryID in alreadyGeneratedReportList) and (directory["mimeType"] == 'application/vnd.google-apps.folder')):
            reportGenerationCount += 1

            testDirectory = subDirectoryInfo(directory)

            # Gathering IDs for further Operation...
            testDirectoryID = testDirectory['test_folder_id']
            test_name = testDirectory['test_name']
            recordingFolderID = testDirectory['recordingFolderID']
            responseFolderID = testDirectory['responseFolderID']
            screenshotFolderID = testDirectory['screenshotFolderID']
            pcapFolderID = testDirectory['pcapFolderID']
            urlInfoID = testDirectory['urlInfoID']

            # Retrieving List of Files in Respective Directories as Pandas Dataframes...
            recordingList = lister(service, driveID, recordingFolderID, 1)
            recordingList = pd.DataFrame(recordingList)

            responseList = lister(service, driveID, responseFolderID, 1)
            responseList = pd.DataFrame(responseList)

            screenshotList = lister(service, driveID, screenshotFolderID, 1)
            screenshotList = pd.DataFrame(screenshotList)

            pcapList = lister(service, driveID, pcapFolderID, 1)
            pcapList = pd.DataFrame(pcapList)

            print("\nTest Name::: [ {0} ]".format(test_name))
            # print("\nRecording List:::", recordingList)
            # print("\nResponse List:::", responseList)
            # print("\nScreenshot List:::", screenshotList)
            # print("\nPCAP List:::", pcapList)

            # urlInfoCSVLink = sharer(urlInfoID)
            # print("urlInfo.csv LINK:::", urlInfoCSVLink)

            print("\n--Gathering urlInfo.csv Information--", end='')
            request = service.files().get_media(fileId=urlInfoID)
            csvFile = request.execute()
            csvFile = csvFile.decode("utf-8")
            csvFile = io.StringIO(csvFile)
            resultDataframe = pd.read_csv(csvFile)
            resultDataframe = resultDataframe.replace(np.nan, '')
            resultDataframe = resultDataframe.astype(str)
            print(">> DONE")

            resultDataframe = resultDataframe.rename(
                columns=lambda x: x.replace('.1', '') if '.1' in x else x)

            print("--PARSING SPREADSHEET INFORMATION--", end='')
            for index, result in resultDataframe.iterrows():
                index = int(index)
                resultId = f"{result['Test ID']}-{result['Payload ID']}"

                # if (len(str((result['Response Code']))) == 0):
                #     resultDataframe.loc[index, 'Response Code'] = 'N/A'

                if (result['File Downloaded'] == "False"):
                    resultDataframe.loc[index, 'Downloaded File Name'] = 'N/A'
                    resultDataframe.loc[index, 'Downloaded File MD5'] = 'N/A'
                    resultDataframe.loc[index,
                                        'Downloaded File Sha256'] = 'N/A'
                    resultDataframe.loc[index, 'Downloaded File Size'] = 'N/A'
                    resultDataframe.loc[index,
                                        'File First Submission Date'] = 'N/A'
                    resultDataframe.loc[index,
                                        'File Last Submission Date'] = 'N/A'
                    resultDataframe.loc[index,
                                        'File Last Analysis Date'] = 'N/A'

                # Parsing Search IDs...
                recordingSearchId = f"{resultId}.mp4"
                responseSearchId = f"{resultId}.json"
                screenshotSearchId = f"{resultId}.jpeg"
                pcapSearchId = f"{resultId}.pcap"

                # Gathering File IDs...
                recordingID = idreturner(
                    recordingSearchId, recordingList, 'name')
                responseID = idreturner(responseSearchId, responseList, 'name')
                screenshotID = idreturner(
                    screenshotSearchId, screenshotList, 'name')
                pcapID = idreturner(pcapSearchId, pcapList, 'name')

                # Generating shareable links for above File IDs...
                if (recordingID != None):
                    resultDataframe.loc[index,
                                        'Video Link'] = rf"https://drive.google.com/file/d/{recordingID}"
                else:
                    resultDataframe.loc[index, 'Video Link'] = 'N/A'

                if (responseID != None):
                    resultDataframe.loc[index,
                                        'Response Link'] = rf"https://drive.google.com/file/d/{responseID}"
                else:
                    resultDataframe.loc[index, 'Response Link'] = 'N/A'

                if (screenshotID != None):
                    resultDataframe.loc[index,
                                        'Screenshot Link'] = rf"https://drive.google.com/file/d/{screenshotID}"
                else:
                    resultDataframe.loc[index, 'Screenshot Link'] = 'N/A'

                if (pcapID != None):
                    resultDataframe.loc[index,
                                        'PCAP Link'] = rf"https://drive.google.com/file/d/{pcapID}"
                else:
                    resultDataframe.loc[index, 'PCAP Link'] = 'N/A'
            print(">> DONE")

            try:
                # GSPREAD Operation to fill SPREADSHEET...
                print("--POPULATING SPREADSHEET--", end='')

                # Selecting & Retrieving Sheet Value...
                testID = ""
                for i in range(18):
                    index = i+1

                    if (index < 10):
                        # Parsing Sheet Name for T01,T02,...,T09
                        index = str(f"T0{index}")
                    else:
                        # Parsing Sheet Name for T10,T11,...,T18
                        index = str(f"T{index}")

                    if (index in test_name):
                        testID = index
                        break

                # Selecting Worksheet Based on sheetName...
                sheetName = ""
                worksheetList = spreadsheetObject.worksheets()
                for index, worksheet in enumerate(worksheetList):
                    if (testID in worksheet.title):
                        sheetName = worksheetList[index].title
                        break

                if (sheetName == ""):
                    sheetName = "ETC"

                worksheetObject = spreadsheetObject.worksheet(f"{sheetName}")
                sheetdata = worksheetObject.get_all_values()
                finalDataframe = None

                if (len(sheetdata) > 0):
                    sheetColumns = sheetdata[0]

                    if (len(sheetColumns) != len(resultDataframe.columns)):
                        print(f"\nColumns Length Mismatch for {test_name}")
                        raise Exception("Columns Length Mismatch")

                    sheetValues = sheetdata[1:]

                    resultDataframe = resultDataframe.to_dict('records')
                    keys_list = list(resultDataframe[0].keys())
                    values_list = [[item[key] for key in keys_list]
                                   for item in resultDataframe]

                    sheetValues = sheetValues + values_list

                    finalDataframe = pd.DataFrame(
                        sheetValues, columns=sheetColumns)
                else:
                    finalDataframe = resultDataframe

                finalDataframe = finalDataframe.astype(str)

                # Adding Result Row for Results...
                resultRow = [""] * len(finalDataframe.columns)
                resultRow[5] = "Result"
                resultRow = pd.Series(resultRow, index=finalDataframe.columns)

                finalDataframe = pd.concat(
                    [finalDataframe, resultRow.to_frame().T], ignore_index=True)

                # Spreadsheet Dump Operation...
                set_with_dataframe(worksheetObject, finalDataframe)
                print(">> DONE")

                # Adding Result Folder ID to generatedReports.txt (if only the reports are Generated)...
                generatedReports.append(test_name)
                file = open(os.path.normpath(
                    '.\dependencies\generatedReports.txt'), 'a')
                file.write(f"{directoryID}\n")
                file.close()

            except Exception as error:
                print(">> FAILED")
                print("\nGSPREAD ERROR [ {0} ]".format(error))
                pass

    except Exception as error:
        print("Error in Driver Code ::: ", error)
        pass

if (reportGenerationCount == 0):
    print("\n--[RESULT]--No New Reports Generated")
elif ((reportGenerationCount > 0) or len(generatedReports)):
    print("\n--[RESULT]--Generated Reports of :: {0}".format(generatedReports))
