#-*- encoding: utf8 -*-
import win32com.client

class easyExcel:

    def __init__(self,filename=None):
        self.xlApp=win32com.client.Dispatch("Excel.Application")
        if filename:
            self.filename=filename
            self.xlBook=self.xlApp.Workbooks.Open(filename)
        else:
            self.xlBook=self.xlApp.Workbooks.Add()
            self.filename=''

    def save(self,newfilename=None):#保存文件
        if newfilename:
            self.filename=newfilename
            self.xlBook.SaveAs(newfilename)
        else:
            self.xlBook.Save()

    def close(self):
        self.xlBook.Close(SaveChanges=0)
        del self.xlApp

    def getCell(self,sheet,row,col):
        sht=self.xlBook.Worksheets(sheet)
        return sht.Cells(row,col).Value

    def setCell(self,sheet,row,col,value):
        sht=self.xlBook.Worksheets(sheet)
        sht.Cells(row,col).Value=value
    
    def cpSheet(self,before,newfilename):
        shts=self.xlBook.Worksheets
        newshts=newfilename.xlBook.Worksheets
        shts(1).Copy(None,newshts(1))

    def getRange(self, sheet, row1, col1, row2, col2):
        sht = self.xlBook.Worksheets(sheet)  
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2))

    def getRangeValue(self, sheet, row1, col1, row2, col2):  #获得一块区域的数据，返回为一个二维元组
        "return a 2d array (i.e. tuple of tuples)"  
        sht = self.xlBook.Worksheets(sheet)  
        return sht.Range(sht.Cells(row1, col1), sht.Cells(row2, col2)).Value  
    
    def setRange(self, sheet,topRow, leftCol, data):
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(
            sht.Cells(topRow, leftCol), 
            sht.Cells(bottomRow, rightCol)
            ).Value = data

    # def findColPosition():
    #     "在targetExcel中找到最后一天的记录，并在其右侧插入空白列"