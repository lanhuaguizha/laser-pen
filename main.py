import os
import sys

import pyHook
import pythoncom
import win32api

programs = ['chrome.exe', 'notepad.exe', 'notepad.exe']
a = len(programs) -1

# 会按着index:2, 1, 0, 2, 1...循环打开programs里的程序
def openProgrames(index):
    try:
        os.system("taskkill /F /IM %s" % programs[index])
    finally:
        win32api.ShellExecute(1, 'open', '%s' % programs[index], '', '', 1)
    pass


# 激光笔上键逻辑
def dispatchUpKeyEvent():
    global a
    openProgrames(a)
    a -= 1
    if a < 0:
        a = len(programs) - 1
    pass


# 激光笔下键逻辑
def dispatchDownKeyEvent():
    pass


# 激光笔Tab键逻辑：结束本脚本程序，结束激光笔监听
def dispatchTabKeyEvent():
    sys.exit()
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
    hm = pyHook.HookManager()
    # 监听所有的键盘事件
    hm.KeyDown = onKeyboardEvent
    # 设置键盘”钩子“
    hm.HookKeyboard()
    # 进入循环侦听，需要手动进行关闭，否则程序将一直处于监听的状态。可以直接设置而空而使用默认值
    pythoncom.PumpMessages()
    # 我也不知道为什么直接放置到main函数中不管用


if __name__ == "__main__":
    main()
