from BasicInfo import *
import time
from Excelwin32com import ExcelApp
import shutil
import os
import datetime
import sys


class HRTable(object):
    excelInPath = DIR_IN + "hr table.xlsx"
    excelOutPath = DIR_OUT + "hr table.xlsx"
    hrText = DIR_IN + "hr.txt"
    hrWagesFound = "66020110"
    hrSalesEfficiencyRMB = "6001"
    hrSalesEfficiencyMT = 27

    """datetime:%y-%m"""
    def __init__(self, dtime, tonns):
        self.dtime = datetime.datetime(int(dtime[:4]), int(dtime[4:6]), int(dtime[6:8]),23,0,0)
        self.sheetName = self.dtime.strftime("%B %y")
        self.excel: ExcelApp = None
        self.sheet = None
        self.sheetLst = None
        self.sheetLstYr = None
        self.pleDic = []
        self.tonns: ExcelApp = tonns

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
            self.pleDic.append(float(value))
        f.close()

    def handle(self):
        self.sheetLst.Range("A1", "J42").Copy(self.sheet.Range("A1", "J42"))
        self.sheet.Cells(1, 2).Value = self.dtime.strftime("%Y-%m-%d")
        self.sheet.Cells(3, 2).Value = "=B4+B5"
        self.sheet.Cells(10, 2).Value = "=SUM(B11:B15)"
        for i in range(6):
            self.sheet.Cells(i + 4, 2).Value = self.pleDic[i]
            self.sheet.Cells(i + 4, 2).NumberFormatLocal = "0"
            self.sheet.Cells(i + 11, 2).Value = self.pleDic[i + 6]
            self.sheet.Cells(i + 11, 2).NumberFormatLocal = "0"

        b16 = self.pleDic[-1]
        self.sheet.Cells(17, 2).Value = TrialBalanceExcel.excel.sumofspecialname(HRTable.hrWagesFound,TrialBalanceExcel.billOfAccountNum,TrialBalanceExcel.currentBorrowNum, TrialBalanceExcel.rowOriginalPosition,TrialBalanceExcel.mSheet)
        self.sheet.Cells(18, 2).Value = "=B17/B16"
        self.sheet.Cells(20, 2).Value = TrialBalanceExcel.excel.sumofspecialname(HRTable.hrSalesEfficiencyRMB,TrialBalanceExcel.billOfAccountNum,TrialBalanceExcel.currentLoadNum, TrialBalanceExcel.rowOriginalPosition,TrialBalanceExcel.mSheet) / b16
        tonnsSheet = self.tonns.getsheetfromname(TonnsExcel.sheetNameQuantity)
        hrSalesEfficiencyMTValue = str(tonnsSheet.Cells(HRTable.hrSalesEfficiencyMT, TonnsExcel.columnMonthOffset + self.dtime.month).Value)
        self.sheet.Cells(19, 2).Value = float(hrSalesEfficiencyMTValue)
        if self.excel.wBook.Sheets.Count >= 13:
            self.sheetLstYr = self.excel.wBook.Sheets[self.excel.wBook.Sheets.Count - 13]
        lastObject = self.excel.copycolumnfromsheet(1, 2, 20, self.sheetLst)
        for i in range(lastObject.__len__()):
            self.sheet.Cells(i + 1, 3).Value = lastObject[i]
        lastObject = self.excel.copycolumnfromsheet(1, 2, 20, self.sheetLstYr)
        for i in range(lastObject.__len__()):
            self.sheet.Cells(i + 1, 5).Value = lastObject[i]
        self.sheet.Range("A1", "F20").EntireColumn.AutoFit()

    def close(self):
        self.sheetLst = None
        self.sheet = None
        self.sheetLstYr = None
        self.excel.close()
        basicinfoclose()

    def main(self):
        print("HR 表格开始生成")
        print("生成人员列表，根据hr.txt.")
        self.createpeoplelist()
        print("生成新的sheet")
        self.createnewsheet()
        print("处理数据")
        self.handle()
        print("HR 表格完成")

if __name__ == "__main__":
    tonns = ExcelApp()
    tonns.open(TonnsExcel.outPath, SECRET)
    hr = HRTable("20160831", tonns)
    hr.main()
    hr.close()


