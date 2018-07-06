from typing import List, Iterator
import os

DIR_IN = os.getcwd() + "\input\\"
DIR_OUT =os.getcwd() + "\output\\"
SECRET = "753"

class Product(object):
    """每个产品的基本数据信息"""

    def __init__(self, tonns, transit, chmc, rowNum, isKg):
        self.tonns = tonns
        self.transit = transit
        self.chmc = chmc
        self.tonnsPurchaseRowNum = rowNum
        self.tonnsSalesRowNum = rowNum + 25
        self.tonnsInventoriesRowNum = rowNum + 50
        self.isKg = isKg

    def accountingset(self, tariff, transit, product, inventory):
        self.tariff = tariff
        self.transit_acc = transit
        self.product = product
        self.inventory = inventory
        return self


class ProductList:
    """初始化所有产品数据，创建productlist存储产品信息"""
    productlist: List[Product] = []

    l1: List[tuple] = [("25", "Volgamid 25", "1001", 5, False),
                       ("27", "Volgamid 27", "1002", 6, False),
                       ("34", "Volgamid 34 (F)", "1004", 8, False),
                       ("34F", "Volgamid 34 (F)", "1006", 13, False),
                       ("24 SD", "Volgamid 24 SD", "1007", 11, False),
                       ("24", "Volgamid 24", "1008", 12, False),
                       ("32", "Other grades", "1009", 7, False),
                       ("PA-6 recycled", "5101", "5101", 15, False),
                       ("36", "Other grades", "0000", 9, False),
                       ("40", "Other grades", "0000", 10, False),
                       ("Caprolactam", "Caprolactam", "1101", 16, False),
                       ("Polyamide fiber 0.33 tex", "yarns and fibre", "2101", 24, True),
                       ("Polyamide fiber 0.48 tex", "yarns and fibre", "2102", 22, True),
                       ("Polyamide fiber 0.68 tex", "yarns and fibre", "2103", 23, True),
                       ("Polyamide fiber 1 tex", "yarns and fibre", "2104", 21, True),
                       ("Polyamide light-stabilized yarn 93.5 tex", "yarns and fibre", "3101", 17, True),
                       ("Polyamide not the thermostabilized yarn 187tex", "yarns and fibre", "3102", 18, True),
                       ("yarn 144tex", "yarns and fibre", "3103", 19, True),
                       ("yarn 94tex", "yarns and fibre", "3104", 20, True),
                       ("fish net", "4101", "4101", 25, True),
                       ("        RST", "0000", "0000", 14, False),
                       ("others", "0000", "0000", 26, False)]
    for x in range(len(l1)):
        pr = Product(l1[x][0],l1[x][1],l1[x][2],l1[x][3],l1[x][4])
        productlist.append(pr)
    l2: List[tuple] = [("1402030102", "14020301", "1402030101", "14050201"),
                       ("1402030202", "14020302", "1402030201", "14050202"),
                       ("1402030302", "14020303", "1402030301", "14050203"),
                       ("1402030402", "14020304", "1402030401", "14050204"),
                       ("1402030502", "14020305", "1402030501", "14050205"),
                       ("1402030602", "14020306", "1402030601", "14050206"),
                       ("1402030702", "14020307", "1402030701", "14050207"),
                       ("1402802", "140208", "1402801", "140507"),
                       ("null", "null","null", "null"),
                       ("null", "null", "null", "null"),
                       ("14020402", "140204", "14020401", "140503"),
                       ("1402050102", "14020501", "1402050101", "14050401"),
                       ("1402050202", "14020502", "1402050201", "14050402"),
                       ("1402050302", "14020503", "1402050301", "14050403"),
                       ("1402050402", "14020504", "1402050401", "14050404"),
                       ("1402060102", "14020601", "1402060101", "14050501"),
                       ("1402060202", "14020602", "1402060201", "14050502"),
                       ("1402060302", "14020603", "1402060301", "14050503"),
                       ("1402060402", "14020604", "1402060401", "14050504"),
                       ("14020702", "140207", "14020701", "140506"),
                       ("null", "000000", "null", "null"),
                       ("220203", "140299", "null", "140511")]
    for x in range(len(productlist)):
        productlist[x].accountingset(l2[x][0],l2[x][1],l2[x][2],l2[x][3])

    pa6_25 = productlist[0]
    pa6_27 = productlist[1]
    pa6_34 = productlist[2]
    pa6_34F = productlist[3]
    pa6_24SD = productlist[4]
    pa6_24 = productlist[5]
    pa6_32 = productlist[6]
    pa6_recycled = productlist[7]
    pa6_36 = productlist[8]
    pa6_40 = productlist[9]
    caprolactam = productlist[10]
    fiber033 = productlist[11]
    fiber048 = productlist[12]
    fiber068 = productlist[13]
    fiber100 = productlist[14]
    yarn935 = productlist[15]
    yarn187 = productlist[16]
    yarn144 = productlist[17]
    yarn94 = productlist[18]
    fishingNet = productlist[19]
    rst_PA = productlist[20]
    others = productlist[21]


class TransitExcel:
    inventorynum = 1
    supplierName = "C";
    supplierNum = 3;
    lastQuatityName = "F";
    lastQuatityNum = 6;
    lastAmountName = "G";
    lastAmountNum = 7;
    currentPurchaseQuatityName = "H";
    currentPurchaseQuatityNum = 8;
    currentPurchaseAmountName = "I";
    currentPurchaseAmountNum = 9;
    currentCloseAccountQuatityName = "L";
    currentCloseAccountQuatityNum = 12;
    currentCloseAccountAmountName = "M";
    currentCloseAccountAmountNum = 13;
    currentSurplusQuatityName = "O";
    currentSurplusQuatityNum = 15;
    currentSurplusAmountName = "P";
    currentSurplusAmountNum = 16;
    rowOriginalPosition = 5;
    sheetName = "sheet1";
    excelName = "\\在途货物余额表.xls";
    inPath = DIR_IN + excelName


class TrialBalanceExcel:
    billOfAccountName = "C"
    billOfAccountNum = 3
    startName = "H"
    startNum = 8
    currentBorrowName = "J"
    currentBorrowNum = 10
    currentLoadName = "L"
    currentLoanNum = 12
    endName = "0"
    endNum = 15
    rowOriginalPosition = 2
    machineSheet = "sheet1"
    humanSheet = "sheet2"
    excelName = "\\发生额及余额表.xls"
    inPath = DIR_IN + excelName


class OutBoundExcel:
    inventoryName = "G"
    inventoryNum = 7
    quatityName = "K"
    quatityNum = 11
    amountName = "M"
    amountNum = 13
    rowOriginalPosition = 2
    sheetName = "sheet1"
    excelName = "\\出库汇总表.XLS"
    inPath = DIR_IN + excelName


class ReceiveAccount:
    balanceName = "R";
    balanceNum = 18;
    customSerialName = "A";
    customSerialNum = 1;
    dayName = "U";
    dayNum = 21;
    sheetName = "sheet1";
    billDateTimeName = "J";
    billDateTimeNum = 10;
    rowOriginalPosition = 3;
    excelName = "\\应收账龄分析.xls";
    dayOffset = 90;


class PayAccount:
    billDateTimeName = "J";
    billDateTimeNum = 10;
    dayOffset = 90;
    dayNum = 20;
    sheetName = "sheet1";
    subjectName = "F";
    subjectNum = 6;
    supplySerialName = "A";
    supplySerialNum = 1;
    surplusName = "Q";
    surplusNum = 17;
    subjectDetailArray = ["220201", "220202"]
    rowOriginalPosition = 3;
    excelName = "\\应付账龄分析.xls";


class Supply(object):
    def __init__(self, serialnum, name):
        self.serialnum = serialnum
        self.name = name


class SupplyList:
    kuibyshevazotHK = Supply("Kuibyshevazot HK","V002")
    shekinoRST = Supply("Shekino (RST)","V003")
    kurskhimvolokno = Supply("Kurskhimvolokno", "V004")
    ojsckuibyshevazot = Supply("OJSC Kuibyshevazot", "V001")
    # special one find in trialBalance table.
    chinaSomething = Supply("warehouse fee and transportation fee include the purchase in china", "220203")


class TonnsExcel:
    originalName = "\\tonns of good";
    excelEndName = ".xlsx";
    purchaseName = "Purchase (RMB)";
    salesName = "Sales (tn)";
    inventoriesName = "Inventories (tn)";
    sheetNameQuantity = "quantity";
    sheetNamePrice = "price";
    sheetNameAmount = "amount";
    columnMonthOffset = 3;
    columnMonthOffsetName = "C";
    rowMonthPosition = 1;
    secret = "753";
    inPath = DIR_IN + originalName + excelEndName
    outPath = DIR_OUT + originalName + excelEndName


class ContractedFree:
    paSheetName = "PA";
    yarnSheetName = "YARN";
    planSheetName = "PLAN";
    columnPA_Contracted = 8;
    columnPA_Free = 9;
    columnYarn_Contracted = 6;
    columnYarn_Free = 7;
    columnPlan_Total = 7;
    rowOriginalPosition = 3;
    productNameNum = 1;
    rowOfYarnFiber = 12;






