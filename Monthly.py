import datetime
from Excelwin32com import ExcelApp
from BasicInfo import *
import os
import shutil

class Monthly(object):
    inPath = DIR_IN + "monthly report.xlsx"
    outPath = DIR_OUT + "monthly report.xlsx"

    def __init__(self, dtime):
        self.dtime = datetime.datetime(int(dtime[:4]), int(dtime[4:6]), int(dtime[6:8]),23,0,0)
        self.sheetName = self.dtime.strftime("%B %y") + " New"
        self.excel: ExcelApp = None
        self.sheet = None
        self.sheetLst = None
        self.sheetLstYr = None

    def createsheets(self):
        if os.path.exists(DIR_OUT) == False:
            os.makedirs(DIR_OUT)
        if os.path.exists(Monthly.outPath):
            os.remove(Monthly.outPath)
        shutil.copyfile(Monthly.inPath, Monthly.outPath)
        self.excel = ExcelApp()
        self.excel.open(Monthly.outPath, SECRET)
        self.sheetLst = self.excel.wBook.Sheets[self.excel.wBook.Sheets.Count - 1]
        self.sheetLstYr = self.excel.wBook.Sheets[self.excel.wBook.Sheets.Count - 12]
        self.sheet = self.excel.createsheets(self.sheetName)
        self.sheetLst.Range("A1", "AD188").Copy(self.sheet.Range("A1", "AD188"))

    def handle(self):
        pass

    def close(self):
        self.sheet = None
        self.sheetLst = None
        self.sheetLstYr = None
        self.excel.close()

    def main(self):
        self.createsheets()
        self.handle()


if __name__ == "__main__":
    month = Monthly("201608")
    month.main()
    month.close()