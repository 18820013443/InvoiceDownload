from Excel import Data, AppExcel
from Invoice import InvoiceOperations
from ClipPic import Pic
import os
import datetime


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
                dicDateFrom = self.GetDicDate(True, row['开票日期'].strip())
                dicDateTo = self.GetDicDate(False)
                strImgInputPath = InvoiceOperations(dicDateFrom, dicDateTo, strInvoiceNum, isFirstTime)
                screenshotsList.append(strImgInputPath)
            
            objExcel = AppExcel()
            objExcel.WriteToExcel(strFilePath, sheetName, df)
        
        for strImgInputPath in screenshotsList:
            strOutputImgName = os.path.basename(strOutputImgPath)
            strOutputImgPath = os.path.join(os.getcwd(), r'ClipImgs\%s'%strOutputImgName)
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

    def GetDicDate(self, isFromDate, date=''):
        if isFromDate:
            strList = date.split('-')
            dicDateFrom = {
                'elXYear': 150,
                'year': '2022',
                'elYYear': 94,
                'elXMonth': 185,
                'month': '10',
                'elYMonth': 94,
                'elXDay': 219,
                'day': '01',
                'exYDay': 94
            } # Input Params
            dicDateFrom['year'] = strList[0]
            dicDateFrom['month'] = strList[1]
            dicDateFrom['day'] = strList[2].split(' ')[0]
            return dicDateFrom
        else:
            date = datetime.datetime.now()
            dicDateTo = {
                'elXYear': 150,
                'year': '2022',
                'elYYear': 156,
                'elXMonth': 185,
                'month': '11',
                'elYMonth': 156,
                'elXDay':219,
                'day': '15',
                'exYDay': 156
            } # Input Params
            dicDateTo['year'] = str(date.year)
            dicDateTo['month'] = str(date.month)
            dicDateTo['day'] = str(date.day)
            return dicDateTo

if __name__ == '__main__':
    objMain = Main()
    objMain.main()