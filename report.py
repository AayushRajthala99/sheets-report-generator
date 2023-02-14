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

CLIENT_SECRET_FILE = '.\dependencies\client-secret.json'
SERVICE_ACCOUNT_FILE = '.\dependencies\service-account.json'
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

        response = service.files().list(q=query, fields="files(id, name, mimeType)", orderBy='name', includeItemsFromAllDrives=True,
                                        corpora='drive', driveId=driveId, supportsAllDrives=True).execute()

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

if path.isfile('.\dependencies\generatedReports.txt') is False:
    print("--generatedReports.txt File Created!--")
    File = open('.\dependencies\generatedReports.txt', "x")
    File.close()

reportFile = open('.\dependencies\generatedReports.txt', 'rb')
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
worksheetObject = spreadsheetObject.get_worksheet(0)

for directory in resultDirectories:
    try:
        directoryID = str(directory["id"])
        if (not (directoryID in alreadyGeneratedReportList) and (directory["mimeType"] == 'application/vnd.google-apps.folder')):
            reportGenerationCount += 1
            sheetdata = worksheetObject.get_all_values()
            sheetdata = pd.DataFrame(sheetdata[1:], columns=sheetdata[0])
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
            print(">> DONE")

            print("PARSING SPREADSHEET INFORMATION--", end='')
            for index, result in resultDataframe.iterrows():
                index = int(index)
                resultId = f"{result['Test ID']}-{result['Payload ID']}"

                if (len(str((result['Response Code']))) == 0):
                    resultDataframe.loc[index, 'Response Code'] = 'N/A'

                if (result['File Downloaded'] == False):
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

            # GSPREAD Operation to fill SPREADSHEET...
            print("Filling Spreadsheet for Test ::: [ {0} ]".format(
                test_name), end='')

            try:
                sheetdata = pd.concat(
                    [sheetdata, resultDataframe], axis=0, ignore_index=True)
                set_with_dataframe(worksheetObject, sheetdata)
            except Exception as error:
                print("GSPREAD ERROR [ {0} ]".format(error))
            print(">> DONE")

            # Extracted Headers for Spreadsheet...
            # headers = list(resultDataframe[0].keys())

            # Adding Result Folder ID to generatedReports.txt (if only the reports are Generated)...
            generatedReports.append(test_name)
            file = open('.\dependencies\generatedReports.txt', 'a')
            file.write(f"{directoryID}\n")
            file.close()

    except Exception as error:
        print("Error in Driver Code::: ", error)
        pass

if (reportGenerationCount == 0):
    print("\nNo New Reports Generated")
elif ((reportGenerationCount > 0) or len(generatedReports)):
    print("\nGenerated Reports of :: {0}".format(generatedReports))
