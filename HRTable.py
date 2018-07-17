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
        self.excel: ExcelApp = None
        self.sheet = None
        self.sheetLst = None
        self.pleDic = {}

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
        self.excel = ExcelApp()
        self.excel.open(HRTable.excelOutPath, SECRET)
        self.sheetLst = self.excel.wBook.Sheets[self.excel.wBook.Sheets.Count - 1]
        self.sheet = self.excel.createsheets(self.sheetName)[0]

    def createpeoplelist(self):
        f = open(HRTable.hrText)
        l = f.readlines()

        for x in l:
            key, value = x.strip('\n').split(':')
            self.pleDic[key] = float(value)
        f.close()

    def handle(self):
        self.sheetLst.Range("A1", "J42").Copy(self.sheet.Range("A1", "J42"))

    def close(self):
        self.sheetLst = None
        self.sheet = None
        self.excel.close()
        basicinfoclose()

    def main(self):
        self.createpeoplelist()
        self.createnewsheet()
        self.handle()


if __name__ == "__main__":
    hr = HRTable("201608")
    hr.main()
    hr.close()


