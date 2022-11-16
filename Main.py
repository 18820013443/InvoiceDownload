from Excel import Data, AppExcel
from Invoice import InvoiceOperations
from ClipPic import Pic
import os


class Main:

    def main(self):
        objData = Data()
        objPic = Pic()
        inputFileList = self.GetInputFileList()
        screenshotsList = []
        for strFilePath in inputFileList:
            sheetName = 'sheet1'
            isFirstTime = True
            df = objData.GetInputDf(strFilePath, sheetName)

            for index, row in df.iterrows():
                strInvoiceNum = row['发票号码'].strip()
                isFirstTime = True if index == 0 else False
                dicDateFrom = self.GetDicDate(row[''].strip())
                dicDateTo = self.GetDicDate(row[''].strip())
                strImgInputPath = InvoiceOperations(dicDateFrom, dicDateTo, strInvoiceNum, isFirstTime)
                screenshotsList.append(strImgInputPath)
        
        for strImgInputPath in screenshotsList:
            strOutputImgName = os.path.basename(strOutputImgPath)
            strOutputImgPath = os.path.join(os.getcwd(), strOutputImgName)
            objPic.ClipImg(strInputImgPath, strOutputImgPath)
                

    def GetInputFileList(self):
        outputList = []
        strInputFolder = os.path.join(os.getcwd(), r'InputFile')
        inputFileList = os.listdir(strInputFolder)
        for fileName in inputFileList:
            if not '$' in fileName and 'xlsx' in fileName:
                strFilePath = os.path.join(os.getcwd(), r'InputFile\%s'%fileName)
                outputList.append(strFilePath)
        return outputList

    def GetDicDate(self, date):
        pass
        

if __name__ == '__main__':
    objMain = Main()
    objMain.main()