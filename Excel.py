import os
import pandas as pd
import win32com.client as win32


class AppExcel:
    def __init__(self):
        self.KillExcel()
        self.app = win32.DispatchEx('Excel.Application')
        self.app.Application.DisplayAlerts = False
        self.app.Application.ScreenUpdating = False
        self.app.Application.Visible = False
        self.wk = None
        self.hasColIsSuccessful = False

    def KillExcel(self):
        os.system('"taskkill /f /im excel.exe"')

    def AddColumnIsSuccessful(self, sheetName):
        # hasColIsSuccessful = False
        numCols = self.wk.Sheets(sheetName).UsedRange.Columns.Count
        for i in range(numCols):
            strCellValue = str(self.wk.Sheets(sheetName).Cells(1, i + 1).Value)
            if 'isSuccessful' in strCellValue:
                self.hasColIsSuccessful = True
                break
        if not self.hasColIsSuccessful:
            self.wk.Sheets(sheetName).Cells(1, numCols + 1).Value = 'isSuccessful'
        # self.wk.Save()

    def WriteToExcel(self, strFilePath, sheetName, df):
        self.wk = self.app.Workbooks.Open(strFilePath, False, False)
        if not self.hasColIsSuccessful:
            self.AddColumnIsSuccessful(sheetName)
        numIsSuccessfulIndex = df.columns.get_loc('isSuccessful') + 1
        for index, row in df.iterrows():
            self.wk.Sheets('Sheet1').Cells(index + 2, numIsSuccessfulIndex).Value = row['isSuccessful']
        self.wk.Save()
        self.wk.Close()
        self.app.Quit()


class Data:
    def __init__(self):
        self.KillExcel()

    def KillExcel(self):
        os.system('"taskkill /f /im excel.exe"')

    def GetInputDf(self, strFilePath, sheetName):
        df = pd.read_excel(strFilePath, sheet_name=sheetName, dtype='str')
        # df['isSuccessful'] = '' if 'isSuccessful' in df.columns else None
        if not 'isSuccessful' in df.columns:
            df['isSuccessful'] = ''
        df=df[ ~ df['isSuccessful'].str.contains('Y', na=False)]
        return df





if __name__ == '__main__':
    pass