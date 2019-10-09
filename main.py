# coding:utf-8
import os

import PyHook3
import pyautogui
import pythoncom
import win32api
from pymouse import *
from win32con import SW_SHOWMAXIMIZED

# programsToOpen = ["D:\\bulid\\multi_mask.exe", "D:\\AR恐龙可调上下位置\\AR-Dinosaur-adjust.exe",
#                   "C:\\Users\\daping\\Desktop\\企业宣传片.mp4", "C:\\Users\\yongqiangtao\\Desktop\\屏保.png"]
programsToOpen = ["D:\\bulid\\multi_mask.exe", "D:\\AR恐龙可调上下位置\\AR-Dinosaur-adjust.exe", "C:\\Users\\daping\\Desktop\\企业宣传片.mp4", "C:\\Users\\daping\\Desktop\\屏保.jpg"]
programsToClose = ["multi_mask.exe", "AR-Dinosaur-adjust.exe", "Video.UI.exe", "Microsoft.Photos.exe"]
a = len(programsToClose) - 1


# 会按着index:2, 1, 0, 2, 1...循环打开programs里的程序
def openProgrames(index):
    try:
        # 结尾2的时候是一个特殊情况，需要关闭的不是index = 3，而是首个程序
        if index + 1 < len(programsToClose):
            os.system("taskkill /F /IM %s" % programsToClose[index + 1])
        else:
            os.system("taskkill /F /IM %s" % programsToClose[0])

        win32api.ShellExecute(1, 'open', '%s' % programsToOpen[index], '', '', SW_SHOWMAXIMIZED)
    except:
        pass
    pass


# 激光笔上键逻辑
def dispatchUpKeyEvent():
    global a
    openProgrames(a)
    a -= 1
    if a < 0:
        a = len(programsToClose) - 1
    pass


# 激光笔下键逻辑
def dispatchDownKeyEvent():
    m = PyMouse()
    x_dim, y_dim = m.screen_size()
    pyautogui.click(x_dim // 2, y_dim // 2)
    pass


# 激光笔Tab键逻辑：结束本脚本程序，结束激光笔监听
def dispatchTabKeyEvent():
    # sys.exit()
    pass


def onKeyboardEvent(event):
    # 监听键盘事件
    if event.Key == 'Up':
        dispatchUpKeyEvent()
    if event.Key == 'Down':
        dispatchDownKeyEvent()
    if event.Key == 'Tab':
        dispatchTabKeyEvent()
    return True


def main():
    # 创建一个：钩子“管理对象
    hm = PyHook3.HookManager()
    # 监听所有的键盘事件
    hm.KeyDown = onKeyboardEvent
    # 设置键盘”钩子“
    hm.HookKeyboard()
    # 进入循环侦听，需要手动进行关闭，否则程序将一直处于监听的状态。可以直接设置而空而使用默认值
    pythoncom.PumpMessages()
    # 我也不知道为什么直接放置到main函数中不管用


if __name__ == "__main__":
    main()
