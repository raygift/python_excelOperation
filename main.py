import dataOperation

class __main__:

    if __name__=="__main__":

        date="0802"
        filePath=r"C:/Users/bigticket/Desktop/python_excel/"
        
        originData=dataOperation.excelOperation.readData(filePath,date,"广告统计信息汇总",4,4,217,7)
        print("originData")

        targetData=dataOperation.excelOperation.readData(filePath,date,"PC",6,3,35,3)
        print("targetData")

        targetList=dataOperation.excelOperation.matchData(originData,targetData)

        dataOperation.excelOperation.addBlankCol(filePath,date,"PC",1,1)

        dataOperation.excelOperation.writeTargetData(filePath,date,"PC",6,35,targetList)