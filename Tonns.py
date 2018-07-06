import datetime
from Excelwin32com import ExcelApp
from BasicInfo import *
import os
import shutil

class Tonns(object):

    def __init__(self, dtime):
        self.tonns : ExcelApp = None
        self.transit : ExcelApp = None
        self.trialB : ExcelApp = None # trialBalance
        self.outB : ExcelApp = None # outBound
        self.dtime = datetime.date(int(dtime[:4]), int(dtime[4:6]), 1)
        self.qSheet = None
        self.aSheet = None
        self.trialSheet = None
        self.transSheet = None
        self.outSheet = None

    def createexcel(self):
        if os.path.exists(DIR_OUT) == False:
            os.makedirs(DIR_OUT)
        if os.path.exists(TonnsExcel.outPath):
            os.remove(TonnsExcel.outPath)
        shutil.copyfile(TonnsExcel.inPath, TonnsExcel.outPath)
        self.tonns = ExcelApp()
        self.tonns.open(TonnsExcel.outPath, SECRET)
        self.transit = ExcelApp()
        self.transit.open(TransitExcel.inPath, SECRET)
        self.trialB = ExcelApp()
        self.trialB.open(TrialBalanceExcel.inPath)
        self.outB = ExcelApp()
        self.outB.open(OutBoundExcel.inPath)
        self.qSheet = self.tonns.getsheetfromname(TonnsExcel.sheetNameQuantity)
        self.aSheet = self.tonns.getsheetfromname(TonnsExcel.sheetNameAmount)
        self.trialSheet = self.trialB.getsheetfromname(TrialBalanceExcel.machineSheet)
        self.tranSheet = self.transit.getsheetfromname(TransitExcel.sheetName)
        self.outSheet = self.outB.getsheetfromname(OutBoundExcel.sheetName)
    def close(self):
        self.tonns.close()
        self.transit.close()
        self.trialB.close()
        self.outB.close()

    def computetonns(self):
        try:
            colPtion = self.dtime.month + TonnsExcel.columnMonthOffset
            colPtionName = ExcelApp.numbertoletter(colPtion)
            colPtionNameLastM = ExcelApp.numbertoletter((colPtion - 1))

            self.qSheet.Cells(TonnsExcel.rowMonthPosition, colPtion).Value = self.dtime
            self.aSheet.Cells(TonnsExcel.rowMonthPosition, colPtion).Value = self.dtime

            tranColumn = (TransitExcel.currentPurchaseQuatityNum, TransitExcel.currentCloseAccountAmountNum)
            tranQuaAmountList = self.transit.sumandnamelist(*tranColumn,TransitExcel.inventorynum,TransitExcel.rowOriginalPosition,self.tranSheet)
            outColumn = (OutBoundExcel.quatityNum, OutBoundExcel.amountNum)
            outBoundList = self.outB.sumandnamelist(outColumn,OutBoundExcel.inventoryNum,OutBoundExcel.rowOriginalPosition,self.outSheet)

            pAmountSum: float = 0
            sAmountSum: float = 0
            sQuatitySum: float = 0

            for x in ProductList.productlist:
                isHave: bool = False
                # tansit Quatity Amount List
                for y in tranQuaAmountList:
                    if y[0] != x.chmc:
                        continue
                    if x.isKg:
                        self.qSheet.Cells(x.tonnsPurchaseRowNum, colPtion).Value = y[1] / 1000
                    else:
                        self.qSheet.Cells(x.tonnsPurchaseRowNum, colPtion).Value = y[1]
                    self.aSheet.Cells(x.tonnsPurchaseRowNum, colPtion).Value = y[2] + self.trialB.sumofspecialname(x.tariff, TrialBalanceExcel.billOfAccountNum, TrialBalanceExcel.currentBorrowNum, TrialBalanceExcel.rowOriginalPosition, self.trialSheet)
                    isHave = True
                    pAmountSum +=y[2]
                    tranQuaAmountList.remove(y)
                    break
                if isHave == False:
                    self.qSheet.Cells(x.tonnsPurchaseRowNum, colPtion).Value = 0
                    self.aSheet.Cells(x.tonnsPurchaseRowNum, colPtion).Value = 0
                else:
                    isHave = False
                # Out Bound List
                for y in outBoundList:
                    if y[0] != x.chmc:
                        continue
                    if x.isKg:
                        self.qSheet.Cells(x.tonnsSalesRowNum, colPtion).Value = y[1] / 1000
                    else:
                        self.qSheet.Cells(x.tonnsSalesRowNum, colPtion).Value = y[1]
                    self.aSheet.Cells(x.tonnsSalesRowNum, colPtion).Value = y[2]
                    isHave = True
                    sQuatitySum += y[1]
                    sAmountSum += y[2]
                    outBoundList.remove(y)
                    break
                if isHave == False:
                    self.qSheet.Cells(x.tonnsSalesRowNum, colPtion).Value = 0
                    self.aSheet.Cells(x.tonnsSalesRowNum, colPtion).Value = 0
                else:
                    isHave = False
                # Inventories
                self.qSheet.Cells(x.tonnsInventoriesRowNum, colPtion).Value = "=" + colPtionNameLastM + str(x.tonnsInventoriesRowNum) + "+" + colPtionName + str(x.tonnsPurchaseRowNum) + "-" + colPtionName + str(x.tonnsSalesRowNum)
                self.aSheet.Cells(x.tonnsInventoriesRowNum, colPtion).Value = "=" + colPtionNameLastM + str(x.tonnsInventoriesRowNum) + "+" + colPtionName + str(x.tonnsPurchaseRowNum) + "-" + colPtionName + str(x.tonnsSalesRowNum)
            # others
            self.qSheet.Cells(ProductList.others.tonnsPurchaseRowNum, colPtion).Value = 0
            self.aSheet.Cells(ProductList.others.tonnsPurchaseRowNum, colPtion).Value = tranQuaAmountList[-1][2] - pAmountSum + myExcel.ReturnSumFromSheet_Microsoft(NameValue.others.accountingsubject_tariff, NameValue.trialBalanceExcel.billOfAccountNum, NameValue.trialBalanceExcel.currentLoanNum, NameValue.trialBalanceExcel.rowOriginalPosition, trialBalanceMachineSheet);

        except Exception as e:
            print(e)
        else:
            pass
        finally:
            pass

    def main(self):
        self.createexcel()
        self.computetonns()
        self.close()

if __name__ == "__main__":
    tonns = Tonns("201806")
    print(tonns.datetime.year)
    print(tonns.datetime.month)