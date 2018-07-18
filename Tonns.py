import datetime
from Excelwin32com import ExcelApp
from BasicInfo import *
import os
import shutil

class Tonns(object):

    def __init__(self, dtime):
        self.tonns : ExcelApp = None
        self.dtime = datetime.datetime(int(dtime[:4]), int(dtime[4:6]), 1,23,0,0)
        print("生成月份：", self.dtime)
        self.qSheet = None
        self.aSheet = None
        self.colPtion = 0

    def createexcel(self):
        if os.path.exists(DIR_OUT) == False:
            os.makedirs(DIR_OUT)
        if os.path.exists(TonnsExcel.outPath):
            os.remove(TonnsExcel.outPath)
        shutil.copyfile(TonnsExcel.inPath, TonnsExcel.outPath)
        self.tonns = ExcelApp()
        self.tonns.open(TonnsExcel.outPath, SECRET)
        self.qSheet = self.tonns.getsheetfromname(TonnsExcel.sheetNameQuantity)
        self.aSheet = self.tonns.getsheetfromname(TonnsExcel.sheetNameAmount)

    def close(self):
        self.qSheet = None
        self.aSheet = None
        self.tonns.close()


    def computetonns(self):
        try:
            print("Tonns Excel start")
            print("first, compute Quatity and Amount Sheet")
            self.colPtion = self.dtime.month + TonnsExcel.columnMonthOffset
            colPtionName = ExcelApp.numbertoletter(self.colPtion)
            colPtionNameLastM = ExcelApp.numbertoletter((self.colPtion - 1))

            self.qSheet.Cells(TonnsExcel.rowMonthPosition, self.colPtion).Value = self.dtime
            self.aSheet.Cells(TonnsExcel.rowMonthPosition, self.colPtion).Value = self.dtime
            print(str(self.qSheet.Cells(TonnsExcel.rowMonthPosition, self.colPtion).Value))
            print(str(self.aSheet.Cells(TonnsExcel.rowMonthPosition, self.colPtion).Value))

            tranColumn = (TransitExcel.currentPurchaseQuatityNum, TransitExcel.currentCloseAccountAmountNum)
            tranQuaAmountList = TransitExcel.excel.sumandnamelist(*tranColumn,nameindex=TransitExcel.inventorynum,startrow=TransitExcel.rowOriginalPosition,sheet=TransitExcel.sheet)
            outColumn = (OutBoundExcel.quatityNum, OutBoundExcel.amountNum)
            outBoundList = OutBoundExcel.excel.sumandnamelist(*outColumn,nameindex=OutBoundExcel.inventoryNum,startrow=OutBoundExcel.rowOriginalPosition,sheet=OutBoundExcel.sheet)
            pAmountSum: float = 0
            sAmountSum: float = 0
            sQuatitySum: float = 0
            print(1)
            for x in ProductList.productlist:
                if x is ProductList.others:
                    continue
                isHave: bool = False
                # tansit Quatity Amount List
                for y in tranQuaAmountList:
                    if y[0] != x.chmc:
                        continue
                    if x.isKg:
                        self.qSheet.Cells(x.tonnsPurchaseRowNum, self.colPtion).Value = y[1] / 1000
                    else:
                        self.qSheet.Cells(x.tonnsPurchaseRowNum, self.colPtion).Value = y[1]
                    self.aSheet.Cells(x.tonnsPurchaseRowNum, self.colPtion).Value = y[2] + TransitExcel.excel.sumofspecialname(x.tariff, TrialBalanceExcel.billOfAccountNum, TrialBalanceExcel.currentBorrowNum, TrialBalanceExcel.rowOriginalPosition, TrialBalanceExcel.mSheet)
                    isHave = True
                    pAmountSum +=y[2]
                    tranQuaAmountList.remove(y)
                    break
                if isHave == False:
                    self.qSheet.Cells(x.tonnsPurchaseRowNum, self.colPtion).Value = 0
                    self.aSheet.Cells(x.tonnsPurchaseRowNum, self.colPtion).Value = 0
                else:
                    isHave = False
                # Out Bound List
                for y in outBoundList:
                    if y[0] != x.chmc:
                        continue
                    if x.isKg:
                        self.qSheet.Cells(x.tonnsSalesRowNum, self.colPtion).Value = y[1] / 1000
                    else:
                        self.qSheet.Cells(x.tonnsSalesRowNum, self.colPtion).Value = y[1]
                    self.aSheet.Cells(x.tonnsSalesRowNum, self.colPtion).Value = y[2]
                    isHave = True
                    sQuatitySum += y[1]
                    sAmountSum += y[2]
                    outBoundList.remove(y)
                    break
                if isHave == False:
                    self.qSheet.Cells(x.tonnsSalesRowNum, self.colPtion).Value = 0
                    self.aSheet.Cells(x.tonnsSalesRowNum, self.colPtion).Value = 0
                else:
                    isHave = False
                # Inventories
                self.qSheet.Cells(x.tonnsInventoriesRowNum, self.colPtion).Value = "=" + colPtionNameLastM + str(x.tonnsInventoriesRowNum) + "+" + colPtionName + str(x.tonnsPurchaseRowNum) + "-" + colPtionName + str(x.tonnsSalesRowNum)
                self.aSheet.Cells(x.tonnsInventoriesRowNum, self.colPtion).Value = "=" + colPtionNameLastM + str(x.tonnsInventoriesRowNum) + "+" + colPtionName + str(x.tonnsPurchaseRowNum) + "-" + colPtionName + str(x.tonnsSalesRowNum)
            # others
            self.qSheet.Cells(ProductList.others.tonnsPurchaseRowNum, self.colPtion).Value = 0
            self.qSheet.Cells(ProductList.others.tonnsSalesRowNum, self.colPtion).Value = 0
            self.qSheet.Cells(ProductList.others.tonnsInventoriesRowNum, self.colPtion).Value = "=" + colPtionNameLastM + str(ProductList.others.tonnsInventoriesRowNum) + "+" + colPtionName + str(ProductList.others.tonnsPurchaseRowNum) + "-" + colPtionName + str(ProductList.others.tonnsSalesRowNum)
            self.aSheet.Cells(ProductList.others.tonnsPurchaseRowNum, self.colPtion).Value = tranQuaAmountList[-1][2] - pAmountSum + TrialBalanceExcel.excel.sumofspecialname(ProductList.others.tariff, TrialBalanceExcel.billOfAccountNum, TrialBalanceExcel.currentLoadNum, TrialBalanceExcel.rowOriginalPosition, TrialBalanceExcel.mSheet)
            self.aSheet.Cells(ProductList.others.tonnsSalesRowNum, self.colPtion).Value = outBoundList[-1][2] - sAmountSum
            self.aSheet.Cells(ProductList.others.tonnsInventoriesRowNum, self.colPtion).Value = "=" + colPtionNameLastM + str(ProductList.others.tonnsInventoriesRowNum) + "+" + colPtionName + str(ProductList.others.tonnsPurchaseRowNum) + "-" + colPtionName + str(ProductList.others.tonnsSalesRowNum)
            print("Quatity and Amount Sheet has done")
            print("Tonns Excel end")
        except Exception as e:
            print(e)
        else:
            pass
        finally:
            pass

    def computeInventory(self):
        print("Inventory Excel start")
        fpath = DIR_OUT + "Inventory.xlsx"
        if os.path.exists(fpath):
            os.remove(fpath)
        trialColName = ("Name", "Tonns", "TrialBalance")
        iSheetName = "result"
        inventory = ExcelApp()
        inventory.create(fpath)
        inventory.open(fpath)
        trialBINVSheet = inventory.createsheets(iSheetName)[0]
        inventory.inserttitletosheet(trialColName, sheet=trialBINVSheet)
        for i in range(ProductList.productlist.__len__()):
            trialBINVSheet.Cells(2 + i, 1).Value = ProductList.productlist[i].tonns
            trialBINVSheet.Cells(2 + i, 2).Value = self.aSheet.Cells(ProductList.productlist[i].tonnsInventoriesRowNum, self.colPtion)
            sum1 = TrialBalanceExcel.excel.sumofspecialname(ProductList.productlist[i].transit_acc, TrialBalanceExcel.billOfAccountNum, TrialBalanceExcel.endNum, TrialBalanceExcel.rowOriginalPosition, TrialBalanceExcel.mSheet)
            sum2 = TrialBalanceExcel.excel.sumofspecialname(ProductList.productlist[i].inventory, TrialBalanceExcel.billOfAccountNum, TrialBalanceExcel.endNum, TrialBalanceExcel.rowOriginalPosition, TrialBalanceExcel.mSheet)
            trialBINVSheet.Cells(2 + i, 3).Value = sum1 + sum2
            trialBINVSheet.Cells(2 + i, 4).Value = '=c' + str(2 + i) + "-B" + str(2 + i)
            trialBINVSheet.Cells(2 + i, 1).NumberFormatLocal = "0.00"
            trialBINVSheet.Cells(2 + i, 2).NumberFormatLocal = "0.00"
            trialBINVSheet.Cells(2 + i, 3).NumberFormatLocal = "0.00"
            trialBINVSheet.Cells(2 + i, 4).NumberFormatLocal = "0.00"
        trialBINVSheet.Cells(2 + ProductList.productlist.__len__(), 4).Value = "=SUM(D2:D" + str(2 + ProductList.productlist.__len__() - 1)
        trialBINVSheet.Cells(2 + ProductList.productlist.__len__(), 4).NumberFormatLocal = "0"
        inventory.close()
        print("Inventory Excel done")

    def main(self):
        self.createexcel()
        self.computetonns()
        self.computeInventory()



if __name__ == "__main__":
    tonns = Tonns("201608")
    tonns.main()
    tonns.close()
