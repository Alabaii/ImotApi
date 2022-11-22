# This module uses Google Sheets API v4 (and Google Drive API v3 for sharing spreadsheets)

from datetime import datetime
from pprint import pprint

import apiclient.discovery
import googleapiclient.errors
import httplib2
from oauth2client.service_account import ServiceAccountCredentials

from ImotApi import ImotApi


def htmlColorToJSON(htmlColor):
    if htmlColor.startswith("#"):
        htmlColor = htmlColor[1:]
    return {"red": int(htmlColor[0:2], 16) / 255.0, "green": int(htmlColor[2:4], 16) / 255.0, "blue": int(htmlColor[4:6], 16) / 255.0}

class SpreadsheetError(Exception):
    pass

class SpreadsheetNotSetError(SpreadsheetError):
    pass

class SheetNotSetError(SpreadsheetError):
    pass

class Spreadsheet:
    def __init__(self, jsonKeyFileName, debugMode = False):
        self.debugMode = debugMode
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(jsonKeyFileName, ['https://www.googleapis.com/auth/spreadsheets',
                                                                                              'https://www.googleapis.com/auth/drive'])
        self.httpAuth = self.credentials.authorize(httplib2.Http())
        self.service = apiclient.discovery.build('sheets', 'v4', http = self.httpAuth)
        self.driveService = None  # Needed only for sharing
        self.spreadsheetId = None
        self.sheetId = None
        self.sheetTitle = None
        self.requests = []
        self.valueRanges = []

    # Creates new spreadsheet
    def create(self, title, sheetTitle, rows = 1000, cols = 26, locale = 'en_US', timeZone = 'Etc/GMT'):
        spreadsheet = self.service.spreadsheets().create(body = {
            'properties': {'title': title, 'locale': locale, 'timeZone': timeZone},
            'sheets': [{'properties': {'sheetType': 'GRID', 'sheetId': 0, 'title': sheetTitle, 'gridProperties': {'rowCount': rows, 'columnCount': cols}}}]
        }).execute()
        if self.debugMode:
            pprint(spreadsheet)
        self.spreadsheetId = spreadsheet['spreadsheetId']
        self.sheetId = spreadsheet['sheets'][0]['properties']['sheetId']
        self.sheetTitle = spreadsheet['sheets'][0]['properties']['title']

    def share(self, shareRequestBody):
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.driveService is None:
            self.driveService = apiclient.discovery.build('drive', 'v3', http = self.httpAuth)
        shareRes = self.driveService.permissions().create(
            fileId = self.spreadsheetId,
            body = shareRequestBody,
            fields = 'id'
        ).execute()
        if self.debugMode:
            pprint(shareRes)

    def shareWithEmailForReading(self, email):
        self.share({'type': 'user', 'role': 'reader', 'emailAddress': email})

    def shareWithEmailForWriting(self, email):
        self.share({'type': 'user', 'role': 'writer', 'emailAddress': email})

    def shareWithAnybodyForReading(self):
        self.share({'type': 'anyone', 'role': 'reader'})

    def shareWithAnybodyForWriting(self):
        self.share({'type': 'anyone', 'role': 'writer'})

    def getSheetURL(self):
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        if self.sheetId is None:
            raise SheetNotSetError()
        return 'https://docs.google.com/spreadsheets/d/' + self.spreadsheetId + '/edit#gid=' + str(self.sheetId)

    # Sets current spreadsheet by id; set current sheet as first sheet of this spreadsheet
    def setSpreadsheetById(self, spreadsheetId):
        spreadsheet = self.service.spreadsheets().get(spreadsheetId = spreadsheetId).execute()
        if self.debugMode:
            pprint(spreadsheet)
        self.spreadsheetId = spreadsheet['spreadsheetId']
        self.sheetId = spreadsheet['sheets'][0]['properties']['sheetId']
        self.sheetTitle = spreadsheet['sheets'][0]['properties']['title']

    # spreadsheets.batchUpdate and spreadsheets.values.batchUpdate
    def runPrepared(self, valueInputOption = "USER_ENTERED"):
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        upd1Res = {'replies': []}
        upd2Res = {'responses': []}
        try:
            if len(self.requests) > 0:
                upd1Res = self.service.spreadsheets().batchUpdate(spreadsheetId = self.spreadsheetId, body = {"requests": self.requests}).execute()
                if self.debugMode:
                    pprint(upd1Res)
            if len(self.valueRanges) > 0:
                upd2Res = self.service.spreadsheets().values().batchUpdate(spreadsheetId = self.spreadsheetId, body = {"valueInputOption": valueInputOption,
                                                                                                                       "data": self.valueRanges}).execute()
                if self.debugMode:
                    pprint(upd2Res)
        finally:
            self.requests = []
            self.valueRanges = []
        return (upd1Res['replies'], upd2Res['responses'])

    def prepare_addSheet(self, sheetTitle, rows = 1000, cols = 26):
        self.requests.append({"addSheet": {"properties": {"title": sheetTitle, 'gridProperties': {'rowCount': rows, 'columnCount': cols}}}})

    # Adds new sheet to current spreadsheet, sets as current sheet and returns it's id
    def addSheet(self, sheetTitle, rows = 1000, cols = 26):
        if self.spreadsheetId is None:
            raise SpreadsheetNotSetError()
        self.prepare_addSheet(sheetTitle, rows, cols)
        addedSheet = self.runPrepared()[0][0]['addSheet']['properties']
        self.sheetId = addedSheet['sheetId']
        self.sheetTitle = addedSheet['title']
        return self.sheetId

    # Converts string range to GridRange of current sheet; examples:
    #   "A3:B4" -> {sheetId: id of current sheet, startRowIndex: 2, endRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
    #   "A5:B"  -> {sheetId: id of current sheet, startRowIndex: 4, startColumnIndex: 0, endColumnIndex: 2}
    def toGridRange(self, cellsRange):
        if self.sheetId is None:
            raise SheetNotSetError()
        if isinstance(cellsRange, str):
            startCell, endCell = cellsRange.split(":")[0:2]
            cellsRange = {}
            rangeAZ = range(ord('A'), ord('Z') + 1)
            if ord(startCell[0]) in rangeAZ:
                cellsRange["startColumnIndex"] = ord(startCell[0]) - ord('A')
                startCell = startCell[1:]
            if ord(endCell[0]) in rangeAZ:
                cellsRange["endColumnIndex"] = ord(endCell[0]) - ord('A') + 1
                endCell = endCell[1:]
            if len(startCell) > 0:
                cellsRange["startRowIndex"] = int(startCell) - 1
            if len(endCell) > 0:
                cellsRange["endRowIndex"] = int(endCell)
        cellsRange["sheetId"] = self.sheetId
        return cellsRange

    def prepare_setDimensionPixelSize(self, dimension, startIndex, endIndex, pixelSize):
        if self.sheetId is None:
            raise SheetNotSetError()
        self.requests.append({"updateDimensionProperties": {
            "range": {"sheetId": self.sheetId,
                      "dimension": dimension,
                      "startIndex": startIndex,
                      "endIndex": endIndex},
            "properties": {"pixelSize": pixelSize},
            "fields": "pixelSize"}})

    def prepare_setColumnsWidth(self, startCol, endCol, width):
        self.prepare_setDimensionPixelSize("COLUMNS", startCol, endCol + 1, width)

    def prepare_setColumnWidth(self, col, width):
        self.prepare_setColumnsWidth(col, col, width)

    def prepare_setRowsHeight(self, startRow, endRow, height):
        self.prepare_setDimensionPixelSize("ROWS", startRow, endRow + 1, height)

    def prepare_setRowHeight(self, row, height):
        self.prepare_setRowsHeight(row, row, height)

    def prepare_setValues(self, cellsRange, values, majorDimension = "ROWS"):
        if self.sheetTitle is None:
            raise SheetNotSetError()
        self.valueRanges.append({"range": self.sheetTitle + "!" + cellsRange, "majorDimension": majorDimension, "values": values})

    def prepare_mergeCells(self, cellsRange, mergeType = "MERGE_ALL"):
        self.requests.append({"mergeCells": {"range": self.toGridRange(cellsRange), "mergeType": mergeType}})

    # formatJSON should be dict with userEnteredFormat to be applied to each cell
    def prepare_setCellsFormat(self, cellsRange, formatJSON, fields = "userEnteredFormat"):
        self.requests.append({"repeatCell": {"range": self.toGridRange(cellsRange), "cell": {"userEnteredFormat": formatJSON}, "fields": fields}})

    # formatsJSON should be list of lists of dicts with userEnteredFormat for each cell in each row
    def prepare_setCellsFormats(self, cellsRange, formatsJSON, fields = "userEnteredFormat"):
        self.requests.append({"updateCells": {"range": self.toGridRange(cellsRange),
                                              "rows": [{"values": [{"userEnteredFormat": cellFormat} for cellFormat in rowFormats]} for rowFormats in formatsJSON],
                                              "fields": fields}})



GOOGLE_CREDENTIALS_FILE = 'imot-368722-164b5444f663.json'


def create_calls_report(data):
    docTitle = "Отчёт о звонках"
    sheetTitle = "Таблица звонков"
    values = [["Дата", "Оператор", "Начало первого звонка", "Начало последнего звонка", "Общая длительность, c","Количество звонков","Рабочий день оператора, c","Среднее время звонка, c"],]  # header row
    data_row=[]
    for item in data:
        data_row=[1,1,1,1,1,1,1,1]
        for key, value in item.items():
            if key =='the_date' : data_row[0]=value
            if key =='name_employee' : data_row[1]=value
            if key =='time_first' : data_row[2]=datetime.utcfromtimestamp(value).strftime('%Y-%m-%d %H:%M:%S')
            if key =='time_last' : data_row[3]=datetime.utcfromtimestamp(value).strftime('%Y-%m-%d %H:%M:%S')
            if key =='duration_call' : data_row[4]=value/1000
            if key =='total' : data_row[5]=value
            if key =='work_time' : data_row[6]=value
            if key =='average_time' :data_row[7]=value/1000
        values.append(data_row)

    rowCount = len(values) - 1



    ss = Spreadsheet(GOOGLE_CREDENTIALS_FILE, debugMode = False)
    ss.create(docTitle, sheetTitle, rows = rowCount + 3, cols = 9, locale = "ru_RU", timeZone = "Europe/Moscow")
    ss.shareWithAnybodyForWriting()

    ss.prepare_setColumnWidth(0, 80)
    ss.prepare_setColumnWidth(1, 200)
    ss.prepare_setColumnsWidth(2, 7, 200)
    ss.prepare_mergeCells("A1:H1")  # Merge A1:E1



    ss.prepare_setCellsFormat("A1:A1", {"textFormat": {"fontSize": 16}, "horizontalAlignment": "CENTER"})  # Font size 14 and center aligment for A1 cell
    ss.prepare_setCellsFormat("A3:H3", {"textFormat": {"bold": True}, "horizontalAlignment": "CENTER"})  # Bold and center aligment for A3:E3 row

    # ss.prepare_setCellsFormat("A4:B%d" % (rowCount + 3), {"numberFormat": {'type': 'TEXT'}}, fields = "userEnteredFormat.numberFormat")  # Text format for A4:B* columns
    # ss.prepare_setCellsFormat("E4:E%d" % (rowCount + 3), {"numberFormat": {'pattern': '[h]:mm:ss', 'type': 'TIME'}}, fields = "userEnteredFormat.numberFormat")  # Duration number format for E4:E* column
    ss.prepare_setCellsFormat("A4:H%d" % (rowCount + 3), {"numberFormat": {'type': 'TEXT'}}, fields = "userEnteredFormat.numberFormat")  # Text format for A4:B* columns
    #ss.prepare_setCellsFormat("A4:E%d" % (rowCount + 3), {"numberFormat": {'pattern': '[h]:mm:ss', 'type': 'TIME'}}, fields = "userEnteredFormat.numberFormat")  # Duration number format for E4:E* column

    # Bottom border for A3:H3 row
    ss.requests.append({"updateBorders": {"range": {"sheetId": ss.sheetId, "startRowIndex": 2, "endRowIndex": 3, "startColumnIndex": 0, "endColumnIndex": 8},
                                          "bottom": {"style": "SOLID", "width": 1, "color": htmlColorToJSON("#000001")}}})

    ss.prepare_setValues("A1:A1", [[sheetTitle]])
    ss.prepare_setValues("A3:H%d" % (rowCount + 3), values)




    ss.runPrepared()
    print(ss.getSheetURL())

if __name__ == "__main__":
    pass
    