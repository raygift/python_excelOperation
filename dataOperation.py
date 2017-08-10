#-*- encoding: utf8 -*-
import pyExcelOperation

class excelOperation:

    def readData(filepath,date,sheetName,startRow,startCol,endRow,endCol):
        filename="origin"+date+".xlsx"
        file=filepath+filename
        originFile=pyExcelOperation.easyExcel(file)
        multilist=originFile.getRangeValue(sheetName,startRow,startCol,endRow,endCol)
        originFile.close()
        return multilist

    def writeTargetData(filepath,date,sheetName,startRow,startCol,datalist):
        filename="target"+date+".xlsx";
        file=filepath+filename;
        targetFile=pyExcelOperation.easyExcel(file);
        targetFile.setRange(sheetName,startRow,startCol,datalist);
        targetFile.save();
        targetFile.close();

    def matchData(originList,targetList):
        matchlist= [[0 for col in range(1)] for row in range(30)];
        for i in range(len(targetList)):
            for j in range(len(originList)):
                if targetList[i][0]==originList[j][0]:
                    matchlist[i][0]=originList[j][3];   

    def addBlankCol(filepath,date,sheet,row,col):
        count=getColCount(sheet,row,col)
        nextCol=count-3
        file=pyExcelOperation.easyExcel(targetFile);
        rangeObj = file.getRange(sheet,1,nextCol,1,nextCol)
        rangeObj.EntireColumn.Insert()
        file.save()
        file.close()