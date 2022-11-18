import uiautomation as uia


class BasicUIA(object):
    
    def __init__(self):
        pass

    def GetMainWindow(self, name):
        uia.SetGlobalSearchTimeout(20)
        winMain = uia.WindowControl(searchDepth=1, Name=name, foundIndex=1)
        winMain.SetActive()
        # winMain.SetTopmost()
        uia.SetGlobalSearchTimeout(15)
        return winMain
    
    def SetTimeOut(self, seconds):
        uia.SetGlobalSearchTimeout(seconds)

    # def FindEl(self, elParent, ctlType, type, param, depth):
    def FindEl(self,*args,**kwargs):
        return self.Dispatch(*args,**kwargs)

    def Dispatch(self,*args,**kwargs):
        handler = getattr(self, kwargs['ctlType'], None)
        return handler(*args, **kwargs)

    def WinCtl(self, *args, **kwargs):
        if args[0]:
            elParent = args[0]
        depth, type, param = kwargs['depth'], kwargs['type'], kwargs['param']
        if type == 'id':
            el = elParent.WindowControl(searchDepth=depth, AutomationId=param, foundIndex=1)
            el.SetActive()
            return el
        elif type == 'name':
            el = elParent.WindowControl(searchDepth=depth, Name=param, foundIndex=1)
            el.SetActive()
            return el

    def PaneCtl(self, *args, **kwargs):
        if args[0]:
            elParent = args[0]
        depth, type, param = kwargs['depth'], kwargs['type'], kwargs['param']
        if type == 'id':
            return elParent.PaneControl(searchDepth=depth, AutomationId=param, foundIndex=1)
        elif type == 'name':
            return elParent.PaneControl(searchDepth=depth, Name=param, foundIndex=1)

    def BtnCtl(self, *args, **kwargs):
        if args[0]:
            elParent = args[0]
        depth, type, param = kwargs['depth'], kwargs['type'], kwargs['param']
        if type == 'id':
            return elParent.ButtonControl(searchDepth=depth, AutomationId=param, foundIndex=1)
        elif type == 'name':
            return elParent.ButtonControl(searchDepth=depth, Name=param, foundIndex=1)



if __name__ == '__main__':
    pass
    # objBase = BasicUIA()
    # uia.WindowControl()
    # winMain = objBase.GetMainWindow('文件资源管理器')
    # el = objBase.FindEl(winMain, ctlType='PaneCtl', type='name', param='文件资源管理器', depth=1)
    # print(el.Name)
