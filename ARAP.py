import datetime
from Excelwin32com import ExcelApp
from BasicInfo import *
import os
import Monthly


class FirstOf10Custom:
    def __init__(self, num, name, balance, over3month, overdue):
        self.name: str = name
        self.num: str = num
        self.balance: float = balance
        self.over3Month: float = over3month
        self.overDue: float = overdue


class ARAP(object):
    findName = "A/R A/P"
    findColumn = 1
    rowOffsetStatic_AR = 5
    rowOffsetStatic_AP = 19
    colCustomerName = 2
    colCustomerTotal = 3
    colExceed3Month = 4
    colOverDue = 5
    colLstMonthOffset = 6
    colChangeOffset = 10
    colLstYearOffset = 13
    colDeltaOffset = 15

    def __init__(self, dtime):
        self.dtime: datetime.datetime = dtime

    def sumandnamelistspecial(self, colname, colvalue, startrow, sheet, isoverdate=False):
        result = []
        for i in range(startrow,sheet.UsedRange.Rows.Count):
            current = []
            dt = str(sheet.Cells(i, colname).Value)
            isHave = self.datetimeisHave(dt, isoverdate)
            if isHave:
                name = str(sheet.Cells(i, 1).Value)
                for x in result:
                    if name == x[0]:
                        x[1] += ExcelApp.converttofloat(sheet.Cells(i, colvalue).Value)
                    else:
                        current.append(name)
                        current.append(ExcelApp.converttofloat(sheet.Cells(i, colvalue).Value))
                        result.append(current)
        return result

    def datetimeisHave(self, dtime, isoverdate=False) -> bool:
        result = False
        if isoverdate:
            if dtime.__len__() > 0:
                if int(dtime) > 0:
                    result = True
        else:
            if dtime.__len__() > 8:
                dTReal = datetime.datetime(int(dtime[:4]), int(dtime[5:7]), int(dtime[8:10]), 23, 0, 0) + datetime.timedelta(
                    days=90)
                if dTReal.year < self.dtime.year:
                    result = True
                elif dTReal.timetuple()[7] < self.dtime.timetuple()[7]:
                    result = True
        return result

    def EbullitionSorter(self, elist, position):
        done = False
        temp1, temp2 = 0
        for x in elist:
            if done == True:
                break
            done = True
            for i in range(elist.__len__()):
                if i == elist.__len__ - 2:
                    break
                temp1 = x[i][position]
                temp2 = x[i + 1][position]
                if temp1 < temp2:
                    done = False
                    curr = x[i].copy()
                    x[i] = x[i + 1].copy()
                    x[i + 1] = curr

    def main(self, month: Monthly.Monthly):
        rcAccBalance = ReceiveAccount.excel.sumandnamelist(ReceiveAccount.balanceNum, nameindex=ReceiveAccount.customSerialNum, startrow=ReceiveAccount.rowOriginalPosition)
        over3L = self.sumandnamelistspecial(ReceiveAccount.billDateTimeNum, ReceiveAccount.balanceNum, ReceiveAccount.rowOriginalPosition, ReceiveAccount.sheet)
        overDL = self.sumandnamelistspecial(ReceiveAccount.dayNum, ReceiveAccount.balanceNum, ReceiveAccount.rowOriginalPosition, ReceiveAccount.sheet)
        self.EbullitionSorter(rcAccBalance, 1)
        f10lst = []
        len = rcAccBalance.__len__()
        if len > 10:
            len = 10
        for i in range(len):
            first = FirstOf10Custom(str(rcAccBalance[i][0]), "", rcAccBalance[i][1] / 100, 0, 0)
            f10lst.append(first)
            rcAccBalance.remove(rcAccBalance[i])
        rIndexCurr = month.excel.getrownumfromsheet(ARAP.findName, ARAP.findColumn, month.sheet)
        rIndexCurr += ARAP.rowOffsetStatic_AR
        for i in range(rIndexCurr, f10lst.__len__() + rIndexCurr):

