import BasicInfo
import time


class HRTable(object):
    def __init__(self, datetime):
        self.datetime = datetime

    excelInPath = BasicInfo.DIR_IN + "hr table.xlsx"
    excelOutPath = BasicInfo.DIR_OUT + "hr table.xlsx"
    hrWagesFound = "66020110";
    hrSalesEfficiencyRMB = "6001";
    hrSalesEfficiencyMT = 27;
    monthDic

    def getsheetname(self):
        monthV = self.datetime[-2:]
        yearV = self.datetime[-4:-2]


