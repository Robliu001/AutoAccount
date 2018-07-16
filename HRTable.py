from BasicInfo import *
import time
from Excelwin32com import ExcelApp
import shutil
import os
import datetime
import sys


class HRTable(object):
    """datetime:%y-%m"""

    def __init__(self, dtime):
        self.dtime = datetime.datetime(int(dtime[:4]), int(dtime[4:6]), 1,23,0,0)
        self.sheetName = self.dtime.strftime("%B %y")
        self.excel = None
        self.sheet = None

    excelInPath = DIR_IN + "hr table.xlsx"
    excelOutPath = DIR_OUT + "hr table.xlsx"
    hrText = DIR_IN + "hr.txt"
    hrWagesFound = "66020110"
    hrSalesEfficiencyRMB = "6001"
    hrSalesEfficiencyMT = 27

    def createnewsheet(self):
        if os.path.exists(DIR_OUT) == False:
            os.makedirs(DIR_OUT)
        if os.path.exists(HRTable.excelOutPath):
            os.remove(HRTable.excelOutPath)
        shutil.copyfile(HRTable.excelInPath, HRTable.excelOutPath)
        self.hrApp = ExcelApp()
        self.hrApp.open(HRTable.excelOutPath)
        self.hrSheetLst = self.hrApp.wBook.Sheets[self.hrApp.wBook.Sheets.Count - 1]
        self.hrSheet = self.hrApp.createsheets(self.sheetName)



    def createpeoplelist(self):
        f = open(HRTable.hrText)
        l = f.readlines()
        self.pleDic = {}
        for x in l:
            key, value = x.strip('\n').split(':')
            self.pleDic[key] = float(value)


if __name__ == "__main__":
    hr = HRTable("201806")
    print(hr.sheetName)
    # hr.createnewsheet()
    if hr.excel is not None:
        hr.excel.close()

