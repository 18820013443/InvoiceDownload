from Excel import Data, AppExcel
from Invoice import InvoiceSoftware
from ClipPic import Pic
from Log import Log
import os
import datetime


class Main:
    def __init__(self) -> None:
        self.objPic = Pic()
        log = Log()
        self.logger = log.GetLog()
        self.objInvoice = InvoiceSoftware()

    def main(self):
        objData = Data()
        
        inputFileList = self.GetInputFileList()
        screenshotsList = []
        for strFilePath in inputFileList:
            sheetName = 'sheet1'
            StrExcelName = os.path.basename(strFilePath)
            isFirstTime, hasError, hasReset = True, False, False
            df = objData.GetInputDf(strFilePath, sheetName)

            for index, row in df.iterrows():
                strInvoiceNum = row['发票号码'].strip()
                isFirstTime = True if index == 0 or hasReset else False
                hasReset, numTryTimes = False, 0
                dicDateFrom = self.GetDicDate(True, row['开票日期'].strip())
                dicDateTo = self.GetDicDate(False)

                # 异常处理，如果出现报错，将df中的数据写入文件，并且将sreenshot处理好，最好写入日志
                while numTryTimes < 3:
                    try:
                        strImgInputPath = self.objInvoice.InvoiceOperations(dicDateFrom, dicDateTo, strInvoiceNum, isFirstTime)
                        screenshotsList.append(strImgInputPath)
                        row['isSuccessful'] = ['Y']
                        break
                    except Exception as e:
                        # hasError = True
                        numTryTimes += 1
                        self.logger.error(e)
                        # objExcel = AppExcel()
                        # objExcel.WriteToExcel(strFilePath, sheetName, df)
                        # self.ClipPicture(screenshotsList)
                        hasReset = True
            
            if df.shape[0] > 0:
                self.ClipPicture(screenshotsList)
                objExcel = AppExcel()
                self.logger.info(f'结果回写到Excel（{StrExcelName}）')
                objExcel.WriteToExcel(strFilePath, sheetName, df)
                self.logger.info(f'回写到Excel（{StrExcelName}）完成')
        
        os.system('"taskkill /f /im skfpShell.exe"')
        os.system('"taskkill /f /im skfp.exe"')
        self.logger.info('所有输入文件处理完成！')

        
       
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

    def ClipPicture(self, screenshotsList):
        self.logger.info('----------------------------开始裁剪发票------------------------------')
        for strInputImgPath in screenshotsList:
            strOutputImgName = os.path.basename(strInputImgPath)
            # strOutputImgName = os.path.basename(strOutputImgPath)
            strOutputImgPath = os.path.join(os.getcwd(), r'ClipImgs\%s'%strOutputImgName)
            try:
                self.objPic.ClipImg(strInputImgPath, strOutputImgPath)
                self.logger.info(f'发票{strOutputImgName}裁剪成功')
            except Exception as e:
                self.logger.error(f'发票{strOutputImgName}裁剪失败--{e}')
                next
        self.logger.info('----------------------------结束裁剪发票------------------------------')

if __name__ == '__main__':
    objMain = Main()
    try:
        objMain.main()
    except Exception as e:
        objMain.logger.error(e)