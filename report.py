import io
import os
import csv
import json
import time
import numpy as np
from os import path
import pandas as pd
from sheet_writer import sharer, writer, shareUrlParse
from service_creator import Create_Service
from service_creator import Create_Service2

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

driveID = '0AGRZ_Z8RC-_fUk9PVA'
parent_directory_id = '13iFrOF6KSWRl09OTUJWl_pGMaSu8_Udf'
spreadsheet_id = '1bVmvF3wigQNekqBvHPxR3aDES8o5HYRet6z5CL66kWc'

# ----------------------------------------------------------------------------------------------------------------------------------

# service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)
service = Create_Service2(SERVICE_ACCOUNT_FILE, API_NAME, API_VERSION, SCOPES)


def sortListofDictionaries(dictionaryList):
    prefix, postfix = dictionaryList['name'].split('-P')
    return (prefix, int(postfix.split('.')[0]))


def idreturner(id, array):
    try:
        result = [item['id'] for item in array if f"{id}" in item['name']]
        if result:
            return result[0]
    except Exception as error:
        print("Error in idreturner Function::: ", error)
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
# print(resultDirectories)

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
# print(alreadyGeneratedReportList)

reportGenerationCount = 0
generatedReports = []

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

            # Retrieving List of Files in Respective Directories...
            recordingList = lister(service, driveID, recordingFolderID)
            responseList = lister(service, driveID, responseFolderID)
            screenshotList = lister(service, driveID, screenshotFolderID)
            pcapList = lister(service, driveID, pcapFolderID)

            # Sorted List of Dictionaries...
            recordingList.sort(key=sortListofDictionaries)
            responseList.sort(key=sortListofDictionaries)
            screenshotList.sort(key=sortListofDictionaries)
            pcapList.sort(key=sortListofDictionaries)

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
            resultDataframe = resultDataframe.to_dict('records')
            print(">> DONE")

            # Extracted Headers for Spreadsheet...
            headers = list(resultDataframe[0].keys())

            print(
                "\n--Started Link Generation for Recording, Response, Screenshot & Packet Capture Files--")

            for result in resultDataframe:
                try:
                    resultId = f"{result['Test ID']}-{result['Payload ID']}"

                    print(
                        "\n--Generating Links for [ {0} ]--".format(resultId), end='')

                    # Gathering File IDs...
                    recordingID = idreturner(resultId, recordingList)
                    responseID = idreturner(resultId, responseList)
                    screenshotID = idreturner(resultId, screenshotList)
                    pcapID = idreturner(resultId, pcapList)

                    print("Payload File IDs==", recordingID,
                          responseID, screenshotID, pcapID)

                    # Generating shareable links for above File IDs...
                    if (recordingID):
                        # result['Video Link'] = f"{shareUrlParse(recordingID)}"
                        result['Video Link'] = rf"https://drive.google.com/file/d/{recordingID}"
                    else:
                        result['Video Link'] = 'N/A'

                    if (responseID):
                        # result['Response Link'] = f"{shareUrlParse(responseID)}"
                        result['Response Link'] = rf"https://drive.google.com/file/d/{responseID}"
                    else:
                        result['Response Link'] = 'N/A'

                    if (screenshotID):
                        # result['Screenshot Link'] = f"{shareUrlParse(screenshotID)}"
                        result['Screenshot Link'] = rf"https://drive.google.com/file/d/{screenshotID}"
                    else:
                        result['Screenshot Link'] = 'N/A'

                    if (pcapID):
                        # result['PCAP Link'] = f"{shareUrlParse(pcapID)}"
                        result['PCAP Link'] = rf"https://drive.google.com/file/d/{pcapID}"
                    else:
                        result['PCAP Link'] = 'N/A'

                    print(">> DONE")

                    # Spreadsheet Writer Function Call for Individual Payload
                    writer(spreadsheet_id, resultId, headers, result)

                except Exception as error:
                    resultId = f"{result['Test ID']}-{result['Payload ID']}"

                    print("Error in Generating Link For [ {0} ]::: {1}".format(
                        resultId, error))
                    pass

            print("\n--Link Generation Operation Ended--")
            # print(resultDataframe)

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
