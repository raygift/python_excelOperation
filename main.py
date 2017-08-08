
class __main__:

    if __name__=="__main__":
        targetFile=pyExcelOperation.easyExcel(r'C:/Users/bigticket/Desktop/python_excel/targetFile.xlsx')
        excelFile=pyExcelOperation.easyExcel(r'C:/Users/bigticket/Desktop/python_excel/origin'++'.xlsx')
        multilist=excelFile.getRangeValue("广告统计信息汇总",4,4,217,7)
        spmlist=targetFile.getRangeValue("PC",6,3,35,3)
        targetlist= [[0 for col in range(1)] for row in range(30)]

        for i in range(len(spmlist)):
            for j in range(len(multilist)):
                if spmlist[i][0]==multilist[j][0]:
                    targetlist[i][0]=multilist[j][3] 

        rangeObj = targetFile.getRange("PC",1,36,5,36)
        rangeObj.EntireColumn.Insert()
        targetFile.setRange("PC",6,35,targetlist)

        excelFile.close()

        targetFile.save()
        targetFile.close()