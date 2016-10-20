# import pywinauto
#
# app=pywinauto.Application();
#
# export_win=pywinauto.findwindows.find_windows(title='净网大师官方PC群1')
# print(export_win)
# if export_win:
#     export_dlg=app.window_(handle=export_win[0])
#     print(export_dlg)
# dlg=pywinauto.WindowSpecification.WrapperObject(export_dlg)
# pywinauto.WindowSpecification.Wait(export_dlg,'exists');
# print(pywinauto.controls.HwndWrapper.GetDialogPropsFromHandle(dlg))
#
# print(pywinauto.controls.win32_controls.HwndWrapper.HwndWrapper(dlg).WindowText())
# pywinauto.controls.win32_controls.HwndWrapper.HwndWrapper(dlg).TypeKeys("tag")
#

# !python3
# -*- coding: utf-8 -*-

import time
import automation


def main():
    qqWindow = automation.WindowControl(searchDepth=1, ClassName='TXGuiFoundation', title='净网大师官方PC群1')
    if not qqWindow.Exists(0):
        automation.Logger.WriteLine('please run QQ first', automation.ConsoleColor.Yellow)
        return
    qqWindow.ShowWindow(automation.ShowWindow.Maximize)
    qqWindow.SetActive()
    time.sleep(1)

    # 群聊窗口输入框
    # edit = automation.EditControl(searchFromControl=qqWindow, Name='输入')
    # if !edit
    #     print('没有找到对应的群输入框')
    #     return
    # edit.Click()
    # automation.SendKeys('有人在吗？ {Enter}')
    # automation.SendKeys("{Ctrl}{Enter}")
    # time.sleep(2)

    listControl=automation.ListControl(searchFromControl=qqWindow,Name='')
    listItem=listControl.GetChildren()
    # 确保群里第一个成员可见在最上面
    left, top, right, bottom = listControl.BoundingRectangle
    while listItem[0].BoundingRectangle[1] < top:
        automation.Win32API.MouseClick(right - 5, top + 5)
    for item in listItem:
        item.RightClick()
        itemName=item.Name
        print(item)
        menu = automation.MenuControl(searchDepth=1, ClassName='TXGuiFoundation')
        menuItems = menu.GetChildren()
        for menuItem in menuItems:
            if menuItem.Name == '发送消息':
                menuItem.Click()
                time.sleep(0.5)
                currentTalkWin=automation.WindowControl(searchDepth=1,ClassName='TXGuiFoundation')
                # currentTalkWin.ShowWindow(automation.ShowWindow.Maximize)
                currentTalkWin.SetActive()
                time.sleep(0.5)
                edit = automation.EditControl(searchFromControl=currentTalkWin, Name='输入')
                edit.Click()
                automation.SendKeys('有人在吗？ {Enter}')
                automation.SendKeys("{Alt}{F4}")
            # automation.SendKey("{Ctrl}{Enter}")
            # automation.SendKey("{Alt}{F4}")
                break
        item.Click()
        time.sleep(0.5)
        automation.SendKeys('{Down}')
        time.sleep(0.5)
    return

    searchEdit = automation.FindControl(qqWindow,
                                        lambda c:
                                        (isinstance(c, automation.EditControl) or isinstance(c,
                                                                                             automation.ComboBoxControl)) and c.Name == 'Enter your search term'
                                        )
    searchEdit.Click()
    automation.SendKeys('Python-UIAutomation-for-Windows site:github.com{Enter}', 0.05)
    link = automation.HyperlinkControl(searchFromControl=qqWindow,
                                       SubName='yinkaisheng/Python-UIAutomation-for-Windows')
    automation.Win32API.PressKey(automation.Keys.VK_CONTROL)
    link.Click()  # press control to open the page in a new tab
    automation.Win32API.ReleaseKey(automation.Keys.VK_CONTROL)
    newTab = automation.TabItemControl(searchFromControl=tab, SubName='yinkaisheng/Python-UIAutomation-for-Windows')
    newTab.Click()
    starButton = automation.ButtonControl(searchFromControl=qqWindow, Name='Star this repository')
    if starButton.Exists(5):
        automation.GetConsoleWindow().SetActive()
        automation.Logger.WriteLine('Star Python-UIAutomation-for-Windows after 2 seconds',
                                    automation.ConsoleColor.Yellow)
        time.sleep(2)
        qqWindow.SetActive()
        time.sleep(1)
        starButton.Click()
        time.sleep(2)

if __name__ == '__main__':
    main()
    automation.GetConsoleWindow().SetActive()
    input('press Enter to exit')
