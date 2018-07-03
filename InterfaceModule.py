#接口API
from win32com.client import Dispatch
import win32com.client
class easyExcel:
        def __init__(self,filename=None):
                self.xlApp=win32com.client.Dispatch('Excel.Application')
                if filename:
                        self.xlApp.Visible=True
                        self.filename=filename
                        self.xlBook=self.xlApp.Workbooks.Open(Filename=filename,UpdateLinks=2, ReadOnly=False,Format = None,Password='123',WriteResPassword='123')
                else:
                        self.xlBook=self.xlApp.Workbooks.Add()
                        self.filename=''
        def get_sheet(self, sheetName):

                A=[]
                for i in range(1,1000000):
                        if (str(self.xlBook.Worksheets(sheetName).Cells(i, 1)) == 'None'):
                                break
                        #print(self.xlBook.Worksheets(sheetName).Rows(i))
                        B=[]
                        for j in range(1,1000000):
                                if (str(self.xlBook.Worksheets(sheetName).Cells(i, j)) == 'None') and (str(self.xlBook.Worksheets(sheetName).Cells(i, j+1)) == 'None') and (str(self.xlBook.Worksheets(sheetName).Cells(i, j+2)) == 'None'):
                                        break
                                B.append(str(self.xlBook.Worksheets(sheetName).Cells(i, j)))


                        A.append(B)
                #print(A)
                return A
        def close(self):
                self.xlBook.Close(self.filename)
                del self.xlApp
        def save(self):
                self.xlBook.Save()
        def setSheet(self,sheetName,content):
                sht=self.xlBook.Worksheets(sheetName)
                for i in range(len(content)):
                        for j in range(len(content[0])):
                                sht.Cells(i+1,j+1).Value=content[i][j]
                self.save()
        def createSheet(self,sheetName):

                pass




#测试一下
test=easyExcel("C:\Project\\abc.xlsx")
K=test.get_sheet('Sheet1')
print(K)
