import os

import win32com.client


class ExcelApp(object):
    def __init__(self):
        self.app = None
        self.wBook = None
        self.path = None
        self.secret = None

    def open(self, path, secret=None):
        try:
            self.app = win32com.client.Dispatch("Excel.Application")
            self.app.Visible = 0
            self.app.DisplayAlerts = False
            self.wBook = self.app.Workbooks.Open(path, False, None, None, secret, secret)
            self.secret = secret
            self.filename = path
        except Exception as e:
            print(e)
            del self.app
            result = False
        else:
            result = True
        finally:
            return result

    def create(self, path, secret=None, isclose=True):
        try:
            self.app = win32com.client.Dispatch("Excel.Application")
            self.app.Visible = 0
            self.app.DisplayAlerts = False
            self.app.UserControl = False
            self.wBook = self.app.Workbooks.Add(True)
            if not os.path.exists(path):
                self.save(path, secret)
            else:
                result = False
        except Exception as e:
            print(e)
            del self.app
            result = False
        else:
            result = True
        finally:
            if isclose:
                if self.wBook is not None:
                    self.wBook.Close()
                if self.app is not None:
                    del self.app
            return result

    def save(self, path=None, secret=None):
        """保存表格"""
        if path:
            self.filename = path
        else:
            pass
        if hasattr(self, "wBook") and self.wBook is not None:
            self.wBook.SaveAs(self.filename, None, None, secret, secret, None, 1, 2, None, None, None, None)

    def close(self, path=None, secret=None):
        try:
            self.save(path, secret)
            if self.wBook is not None:
                self.wBook.Close()
            if self.app is not None:
                self.app.DisplayAlerts = True
                self.app.Quit()
        except Exception as e:
            print(e)
            result = False
        else:
            result = True
        finally:
            del self.app
            return result

    def createsheets(self, *args):
        if self.wBook is None or args is None:
            return None
        else:
            try:
                sheets = []
                for x in args:
                    sheet = self.wBook.Worksheets.Add()
                    sheet.Name = x
                    sheets.append(sheet)
            except Exception as e:
                print(e)
            else:
                pass
            finally:
                return sheets

    def getsheetfromname(self, name):
        if self.wBook is None or self.wBook.Worksheets.Count < 1:
            return None
        else:
            try:
                sheet = self.eApp.wBook.Worksheets(name)
            except Exception as e:
                print(e)
            else:
                pass
            finally:
                return sheet

    def inserttitletosheet(self, *args, rowindex=1, sheet=None):
        try:
            if sheet is None:
                sheet = self.app.ActiveSheet
            i = 0
            for x in args:
                i += 1
                cell = sheet.Cells(rowindex, i)
                cell.Value = x
                cell.Font.Bold = True
                cell.EntireColumn.AutoFit()
                cell.HorizontalAlignment = -4108
        except Exception as e:
            print(e)
            result = False
        else:
            result = True
        finally:
            return result

    def insertcellstosheet(self, *args, rowindex=1, colindex=1, format=None, sheet=None):
        try:
            if sheet is None:
                sheet = self.app.ActiveSheet
            if format is None:
                format = []
                for i in range(len(args)):
                    format.append(None)
            for x in range(len(args)):
                cell = sheet.Cells(rowindex, colindex)
                if format[x] is not None:
                    cell.NumberFormat = format[x]
                cell.Value = args[x]
                colindex += 1
            cell.EntireColumn.AutoFit()
            cell.HorizontalAlignment = -4108
        except Exception as e:
            print(e)
            result = False
        else:
            result = True
        finally:
            return result

    def getrownumfromsheet(self, value, column, sheet):
        try:
            result = 0
            if sheet is None:
                sheet = self.app.ActiveSheet
            for i in range(1, sheet.UsedRange.Rows.Count + 1):
                sCell = str(sheet.Cells(i, column).Value)
                if sCell == value:
                    result = i
                    break
        except Exception as e:
            print(e)
        else:
            pass
        finally:
            return result

    def copycolumnfromsheet(self, rowindex=1, colindex=1, length=1, sheet=None):
        try:
            result = []
            if sheet is None:
                sheet = self.app.ActiveSheet
            for i in range(length):
                result.append(str(sheet.Cells(rowindex + i, colindex).Value))
        except Exception as e:
            print(e)
            result = None
        else:
            pass
        finally:
            return result

    def sumofspecialname(self, name, nameindex, valueindex, startrow, sheet=None):
        try:
            result = 0
            if sheet is None:
                sheet = self.app.ActiveSheet
            for i in range(startrow, sheet.UsedRange.Rows.Count + 1):
                sname = str(sheet.Cells(i, nameindex))
                if sname == name:
                    result += float(str(sheet.Cells(i, valueindex).Value))
        except Exception as e:
            print(e)
            result = None
        else:
            pass
        finally:
            return result

    def valueofspecialname(self, name, nameindex, valueindex, startrow, sheet=None):
        try:
            result = 0
            if sheet is None:
                sheet = self.app.ActiveSheet
            for i in range(startrow, sheet.UsedRange.Rows.Count + 1):
                sname = str(sheet.Cells(i, nameindex))
                if sname == name:
                    result = float(str(sheet.Cells(i, valueindex).Value))
                    break
        except Exception as e:
            print(e)
            result = None
        else:
            pass
        finally:
            return result

    @staticmethod
    def ishavechinese(string):
        try:
            result = False
            if string is None:
                return result
            for s in string:
                if ord(s) > 128:
                    result = True
                    break
        except Exception as e:
            print(e)
            result = False
        else:
            pass
        finally:
            return result

    @staticmethod
    def numbertoletter(value: int) -> str:
        result: str = ""
        ilist = []
        while value // 26 != 0 or value % 26 != 0:
            ilist.append(value % 26)
            value //= 26

        for i in range(ilist.__len__() - 1):
            print(ilist.__len__())
            if ilist[i] == 0:
                ilist[i + 1] -= 1
                ilist[i] = 26
        if ilist[-1] == 0:
            ilist.remove(ilist[-1])
        ilist.reverse()
        for x in ilist:
            print(x)
            result += str(chr((x + 64)))
        return result

    @staticmethod
    def lettertonumber(value: str) -> int:
        result: int = 0
        powIndex = 0
        l = list(value)
        l.reverse()
        for x in l:
            ir = ord(x) - 64
            result += int(pow(26,powIndex)) * ir
            powIndex += 1
        return result

    def sumandnamelist(self, *args, nameindex, startrow, sheet=None):
        try:
            result = []
            if sheet is None:
                sheet = self.app.ActiveSheet
            for i in range(startrow, sheet.UsedRange.Rows.Count):
                sname = str(sheet.Cells(i, nameindex))
                if ExcelApp.ishavechinese(sname.strip()):
                    continue
                isHave = False
                current = None
                for x in result:
                    if sname == x[0]:
                        isHave = True
                        current = x
                        break
                if isHave:
                    for x in len(args):
                        sheet.Cells(i, args[x]).EntireColumn.AutoFit()
                        current[x + 1] += float(str(sheet.Cells(i, args[x]).Value))
                else:
                    current = []
                    current.append(sname)
                    for x in args:
                        sheet.Cells(i, x).EntireColumn.AutoFit()
                        current.append(float(str(sheet.Cells(i, x).Value)))
                    result.append(current)
            current = []
            current.append("sum")
            for x in args:
                sheet.Cells(sheet.UsedRange.Rows.Count, x).EntireColumn.AutoFit()
                current.append(float(str(sheet.Cells(sheet.UsedRange.Rows.Count, x).Value)))
            result.append(current)
        except Exception as e:
            print(e)
            result = None
        else:
            pass
        finally:
            return result

    typeOfFormat = {"Int": "0", "Double1P": "0.0", "Double2P": "0.00", "Double3P": "0.000", "Percent": "0.00%"}


if __name__ == "__main__":
    eApp = ExcelApp()
    path = "D:\\backup\YingWork\output\\test.xlsx"
    eApp.create(path)
    eApp.open(path)
    eApp.createsheets("s1", "s2", "s3")
    s1 = eApp.getsheetfromname("s1")
    eApp.inserttitletosheet(*("T1", "T2", "T3", "T4", "T5"), rowindex=1, sheet=s1)
    eApp.insertcellstosheet(*(1, 2.1, 3.11, 0.55, "xxx"), rowindex=2, format= \
        [ExcelApp.typeOfFormat["Int"], ExcelApp.typeOfFormat["Double1P"], ExcelApp.typeOfFormat["Double2P"],
         ExcelApp.typeOfFormat["Percent"], None], \
                            sheet=s1)
    line = eApp.getrownumfromsheet("0.55", 4, s1)
    print(line)
    eApp.close()

