import os
import uiautomation as uia
import pyautogui
from Tools import BasicUIA
from Log import Log
import subprocess
import pyperclip
import time


class InvoiceSoftware(BasicUIA):
    def __init__(self):
        super().__init__()
        log = Log()
        self.logger = log.GetLog()

    def OpenSoftware(self):
        self.KillSoftware()
        strPath = r'"C:\Program Files (x86)\Aisino\开票服务器开票软件\Bin\skfpShell.exe" "-h"'
        subprocess.Popen(strPath)
        self.logger.info('打开软件')
        # strPath = r"C:\Program Files (x86)\Aisino\开票服务器开票软件\Bin\skfp.exe"
        # os.system(f'"{strPath}" &')
    
    def KillSoftware(self):
        os.system('"taskkill /f /im skfpShell.exe"')
        os.system('"taskkill /f /im skfp.exe"')

    def LocateEl(self, strPicPath):
        path = os.path.join(os.getcwd(), r'%s'%strPicPath)
        print(path)
        x, y = pyautogui.locateCenterOnScreen(path, confidence=0.9)
        return x, y

    # def CheckWinExists(self, isTopWin, ctlType, winName, depth=1):
    #     self.logger.info(f'Check {winName} exists or not')
    #     if isTopWin:
    #         win = self.GetMainWindow(winName)
    #         win.SetActive()
    #         self.logger.info(f'{winName} exists')
    #         return True, win if win else False
    #     else:
    #         win = self.GetMainWindow(winName)
    #         win.SetActive()
    #         winChild = self.FindEl(win, ctlType=ctlType, type='name', param=winName, depath=depth)
    #         return True, winChild if winChild else False


    def ClickLoginBtn(self):
        # strPath = 'ExamplePics\1_login.png'
        winLogin = self.GetMainWindow('登录-增值税发票开票软件金税盘版')

        totalX, loginX = winLogin.BoundingRectangle.width(), 830 # 830 表示元素中心距离窗口左边距的长度
        totalY, loginY = winLogin.BoundingRectangle.height(), 528 # 528 表示元素中心距离窗口top的高度
        ratioX = loginX / totalX
        ratioY = loginY / totalY

        # x, y = LocateEl(strPath)
        x, y = winLogin.GetPosition(ratioX, ratioY)
        pyautogui.click(x, y)
        self.logger.info('点击登陆按钮')

    def ClickInvoiceMgt(self):
        winMain = self.GetMainWindow('增值税发票税控开票软件（金税盘版）')

        totalX, loginX = winMain.BoundingRectangle.width(), 372
        totalY, loginY = winMain.BoundingRectangle.height(), 107
        ratioX = loginX / totalX
        ratioY = loginY / totalY

        # x, y = LocateEl(strPath)
        x, y = winMain.GetPosition(ratioX, ratioY)
        pyautogui.click(x, y)
        self.logger.info('点击发票管理')
    
    def ClickEl(self, winName, elX, elY, el=None, shouldDoubleClick=False):
        
        win = self.GetMainWindow(winName) if not el else el 
        totalX, loginX = win.BoundingRectangle.width(), elX # elX 表示元素中心距离窗口左边距的长度
        totalY, loginY = win.BoundingRectangle.height(), elY # elY 表示元素中心距离窗口top的高度
        ratioX = loginX / totalX
        ratioY = loginY / totalY

        x, y = win.GetPosition(ratioX, ratioY)
        if not shouldDoubleClick:
            pyautogui.click(x, y)
        else:
            pyautogui.doubleClick(x, y)
        pass

    def ClickQueryInvoice(self, winName, elX, elY1, elY2):
        # Click 查询修复
        self.ClickEl(winName, elX, elY1)
        # Click 已开发票查询
        self.ClickEl(winName, elX, elY2)
        self.logger.info('打开已开发票查询窗口')
    
    def InputDate(self, winName, dicDateFrom, dicDateTo):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        # DateFrom 输入年, 月， 日
        self.ClickAndSendKeys(winInvoiceQuery, dicDateFrom['elXYear'], dicDateFrom['elYYear'], dicDateFrom['year'])
        self.ClickAndSendKeys(winInvoiceQuery, dicDateFrom['elXMonth'], dicDateFrom['elYMonth'], dicDateFrom['month'])
        self.ClickAndSendKeys(winInvoiceQuery, dicDateFrom['elXDay'], dicDateFrom['exYDay'], dicDateFrom['day'])

        # DateTo 输入年, 月， 日
        self.ClickAndSendKeys(winInvoiceQuery, dicDateTo['elXYear'], dicDateTo['elYYear'], dicDateTo['year'])
        self.ClickAndSendKeys(winInvoiceQuery, dicDateTo['elXMonth'], dicDateTo['elYMonth'], dicDateTo['month'])
        self.ClickAndSendKeys(winInvoiceQuery, dicDateTo['elXDay'], dicDateTo['exYDay'], dicDateTo['day'])
        self.logger.info('输入日期')
        
    def ClickAndSendKeys(self, el, elX, elY, value):
        self.ClickEl('发票查询', elX, elY, el=el)
        pyautogui.typewrite(str(value), interval=0.05)

    def InputInvoiceNumber(self, winName, elX, elY, value):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        self.ClickAndSendKeys(winInvoiceQuery, elX, elY, value)
        self.logger.info(f'输入Invoice Number-{value}')

    def ClickQueryBtn(self, winName, elX, elY):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        self.ClickEl('发票查询', elX, elY, el=winInvoiceQuery)
        self.logger.info('点击发票查询按钮')

    def DoubleClickQueryResult(self, winName, elX, elY):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        self.ClickEl('发票查询', elX, elY, el=winInvoiceQuery, shouldDoubleClick=True)

    def ShouldDoubleClickQueryResult(self, winName):
         
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)

        monitoringList = ['Y', 'Y']
        endTime = time.time() + 30
        count = 1
        while True and time.time() < endTime:

            # winMain = self.GetMainWindow(winName)
            # # 找到window 发票查询
            # winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)

            self.SetTimeOut(0.5)
            # winInvoiceQuery.SetActive()
            try:
                el = self.FindEl(winInvoiceQuery, ctlType='PaneCtl', type='name', param='过程提示', depth=1)
                if el.Name == '过程提示' and len(monitoringList) == 2:
                    self.logger.info('找到（过程提示）窗口')
                    monitoringList.pop()
            except:
                count += 1
                self.logger.info('没有找到（过程提示）窗口')
                if len(monitoringList) == 1:
                    monitoringList.pop()
                
            # 如果找了5次还没找到弹窗，直接退出将monitoringList设置为空，退出查找
            if count == 5:
                monitoringList = []

            if len(monitoringList) == 0:
                break


        if len(monitoringList) == 0:
            self.SetTimeOut(25)
            return True
        else:
            return True
            # raise Exception('找弹窗错误')

    def ClickPrintButton(self, winName, elX, elY):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        # 找到Window 增值税专用发票查询
        winInvoiceQueryChild = self.FindEl(winInvoiceQuery, ctlType='PaneCtl', type='name', param='布局', depth=1)
        # 点击打印
        self.ClickEl('增值税专用发票查询', elX, elY, el=winInvoiceQueryChild)
        self.logger.info('点击（增值税专用发票查询）窗口上的（打印）按钮')

        # 点击确定
        try:
            self.SetTimeOut(1)
            # winPopConfirm = self.FindEl(winInvoiceQueryChild, ctlType='PaneCtl', type='name', param='发票打印', depth=1)
            winPopConfirm = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票打印', depth=6)
            time.sleep(0.15)
            self.ClickEl('发票打印', 173, 471, el=winPopConfirm)
            self.logger.info('点击（确认对话框中）中（确认）按钮')
        except:
            time.sleep(0.5)
            self.logger.info('点击（确认对话框中）中（确认）按钮-出错，重新查找点击按钮！')
            winInvoiceQueryChild = self.RecurrenceConfrimWindow()
            winPopConfirm = self.FindEl(winInvoiceQueryChild, ctlType='PaneCtl', type='name', param='发票打印', depth=2)
            self.ClickEl('发票打印', 173, 471, el=winPopConfirm)
        self.SetTimeOut(20)

    def ClickBtnInPrintWindow(self, winName, btnName, isClickLogic=True):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        # 找到Window 增值税专用发票查询
        winInvoiceQueryChild = self.FindEl(winInvoiceQuery, ctlType='PaneCtl', type='name', param='布局', depth=1)

        # 找到打印窗口
        try:
            self.SetTimeOut(1.5)
            # winPrint = self.FindEl(winInvoiceQueryChild, ctlType='WinCtl', type='name', param='打印', depth=2)
            winPrint = self.FindEl(winMain, ctlType='WinCtl', type='name', param='打印', depth=5)
            self.logger.info('找到（打印）窗口')
        except: # 打印窗口直接位于 Main Window下面
            time.sleep(1)
            winInvoiceQuery = self.Recurrence(winName, btnName)
            winPrint = self.FindEl(winInvoiceQuery, ctlType='WinCtl', type='name', param='打印', depth=1)
            # winPrint = self.Recurrence(winName, btnName)
        self.SetTimeOut(15)

        # Click 预览
        if isClickLogic:
            btn = self.FindEl(winPrint, ctlType='BtnCtl', type='name', param='预览', depth=1)
            x, y = btn.GetPosition()
            time.sleep(0.2)
            pyautogui.click(x, y)
            self.logger.info('点击（预览）窗口中（打印）按钮')
            winPrintView = self.FindEl(winPrint, ctlType='WinCtl', type='name', param='打印预览', depth=1)
            winPrintView.Maximize(1.5)
            self.logger.info('最大化发票窗口')
            pass
        else:
            # 关闭预览窗口
            winPrintView = self.FindEl(winPrint, ctlType='WinCtl', type='name', param='打印预览', depth=1)
            self.CloseWindow(winPrintView)
            self.logger.info('关闭（预览）窗口')

            # 关闭打印窗口
            btn = self.FindEl(winPrint, ctlType='BtnCtl', type='name', param='不打印', depth=1)
            x, y = btn.GetPosition()
            pyautogui.click(x, y)
            time.sleep(1)
            self.logger.info('关闭（打印）窗口')

            # 关闭 增值税专用发票查询 窗口
            self.ClickEl('增值税专用发票查询', 1146, 20, el=winInvoiceQueryChild)
            self.logger.info('关闭（增值税专用发票查询）窗口')
            # self.CloseWindow(winInvoiceQueryChild.GetChildren()[0])

            # 关闭 发票查询 窗口
            # self.CloseWindow(winInvoiceQuery)
        pass

    def GetScreenshot(self, strImgInputPath):
        pyautogui.screenshot(strImgInputPath)

    def CloseWindow(self, el):
        objWindowPattern = el.GetWindowPattern()
        objWindowPattern.Close(0.2)
        pass

    def Recurrence(self, winName, btnName):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        # 找到Window 增值税专用发票查询
        winInvoiceQueryChild = self.FindEl(winInvoiceQuery, ctlType='PaneCtl', type='name', param='布局', depth=1)

        # 找到打印窗口
        winPrint = self.FindEl(winInvoiceQueryChild, ctlType='WinCtl', type='name', param='打印', depth=1)
        return winPrint
    
    def RecurrenceConfrimWindow(self, winName):
        winMain = self.GetMainWindow(winName)
        # 找到window 发票查询
        winInvoiceQuery = self.FindEl(winMain, ctlType='PaneCtl', type='name', param='发票查询', depth=1)
        winInvoiceQuery.SetActive()

        # 找到Window 增值税专用发票查询
        winInvoiceQueryChild = self.FindEl(winInvoiceQuery, ctlType='PaneCtl', type='name', param='布局', depth=1)
        return winInvoiceQueryChild


    def InvoiceOperations(self, dicDateFrom, dicDateTo, strInvoiceNum, isFirstTime):
        strWinQueryName = '增值税发票税控开票软件（金税盘版）'
        strImgInputPath = os.path.join(os.getcwd(), r'Screenshots\%s.jpg'%strInvoiceNum)
        objInvoice = InvoiceSoftware()
        if isFirstTime:
            objInvoice.OpenSoftware()

            # Click Login Button
            objInvoice.ClickEl('登录-增值税发票开票软件金税盘版', 830, 528)

            # Click Invocie Mgt
            objInvoice.ClickEl(strWinQueryName, 372, 107)

            # Click 发票查询
            objInvoice.ClickQueryInvoice(strWinQueryName, 372, 197, 242)

        # Input 开始日期 & 结束日期
        objInvoice.InputDate(strWinQueryName, dicDateFrom, dicDateTo)

        # Input Invoice#
        objInvoice.InputInvoiceNumber(strWinQueryName, 193, 216, strInvoiceNum)

        # Click 查询
        objInvoice.ClickQueryBtn(strWinQueryName, 386, 216)

        # Check 弹窗- 过程提示 -是否已经消失
        isDisappeared = objInvoice.ShouldDoubleClickQueryResult(strWinQueryName)

        # 双击查询结果
        if isDisappeared:
            objInvoice.DoubleClickQueryResult(strWinQueryName, 386, 314)

        # Click 打印
        objInvoice.ClickPrintButton(strWinQueryName, 1115, 83)

        # 打开发票预览
        objInvoice.ClickBtnInPrintWindow(strWinQueryName, '预览', isClickLogic=True)

        # Print screenshot
        objInvoice.GetScreenshot(strImgInputPath)
        objInvoice.logger.info(f'Invoice-{strInvoiceNum}截图成功')

        # 关闭窗口
        objInvoice.ClickBtnInPrintWindow(strWinQueryName, '预览', isClickLogic=False)

        # Check 弹窗- 过程提示 -是否已经消失
        isDisappeared = objInvoice.ShouldDoubleClickQueryResult(strWinQueryName)

        # 结束日志
        objInvoice.logger.info('----------------------------------------------------------')

        return strImgInputPath




if __name__ == '__main__':
    pass
    # dicDateFrom = {
    #     'elXYear': 150,
    #     'year': '2022',
    #     'elYYear': 94,
    #     'elXMonth': 185,
    #     'month': '10',
    #     'elYMonth': 94,
    #     'elXDay': 219,
    #     'day': '01',
    #     'exYDay': 94
    # } # Input Params

    # dicDateTo = {
    #     'elXYear': 150,
    #     'year': '2022',
    #     'elYYear': 156,
    #     'elXMonth': 185,
    #     'month': '11',
    #     'elYMonth': 156,
    #     'elXDay':219,
    #     'day': '15',
    #     'exYDay': 156
    # } # Input Params

    # strInvoiceNum = '20189850' # Input Params
    # strWinQueryName = '增值税发票税控开票软件（金税盘版）'
    # strImgInputPath = os.path.join(os.getcwd(), r'Screenshots\%s.jpg'%strInvoiceNum)


    # objInvoice = InvoiceSoftware()
    # objInvoice.OpenSoftware()

    # # Test Function
    # # testKK = objInvoice.shouldDoubleClickQueryResult(strWinQueryName)
    # # objInvoice.ClickPrintButton(strWinQueryName, 1115, 83)
    # # objInvoice.ClickBtnInPrintWindow(strWinQueryName, '预览', isClickLogic=False)

    # # Click Login Button
    # objInvoice.ClickEl('登录-增值税发票开票软件金税盘版', 830, 528)

    # # Click Invocie Mgt
    # objInvoice.ClickEl(strWinQueryName, 372, 107)

    # # Click 发票查询
    # objInvoice.ClickQueryInvoice(strWinQueryName, 372, 197, 242)

    # # Input 开始日期 & 结束日期
    # objInvoice.InputDate(strWinQueryName, dicDateFrom, dicDateTo)

    # # Input Invoice#
    # objInvoice.InputInvoiceNumber(strWinQueryName, 193, 216, strInvoiceNum)

    # # Click 查询
    # objInvoice.ClickQueryBtn(strWinQueryName, 386, 216)

    # # Check 弹窗- 过程提示 -是否已经消失
    # isDisappeared = objInvoice.ShouldDoubleClickQueryResult(strWinQueryName)

    # # 双击查询结果
    # if isDisappeared:
    #     objInvoice.DoubleClickQueryResult(strWinQueryName, 386, 314)

    # # Click 打印
    # objInvoice.ClickPrintButton(strWinQueryName, 1115, 83)

    # # 打开发票预览
    # objInvoice.ClickBtnInPrintWindow(strWinQueryName, '预览', isClickLogic=True)

    # # Print screenshot
    # objInvoice.GetScreenshot(strImgInputPath)

    # # 关闭窗口
    # objInvoice.ClickBtnInPrintWindow(strWinQueryName, '预览', isClickLogic=False)

    # # Check 弹窗- 过程提示 -是否已经消失
    # isDisappeared = objInvoice.ShouldDoubleClickQueryResult(strWinQueryName)

    