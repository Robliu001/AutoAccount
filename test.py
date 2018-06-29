import time
import xlrd
import xlwt
import BasicInfo

class Run(object):

    def main(self):
        for x in BasicInfo.ProductList.productlist:
            print(x.tonns,x.transit,x.chmc,x.tariff,x.transit_acc,x.product,x.inventory,x.tonnsPurchaseRowNum,x.tonnsSalesRowNum,x.tonnsInventoriesRowNum,x.isKg)


if __name__ == "__main__":
    run = Run()
    run.main()







