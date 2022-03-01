from calendar import c
from wsgiref.simple_server import WSGIRequestHandler
import keyboard
import threading
import yaml,win32api,win32com.client
import time
from icecream import ic
import win32gui as w
import win32con
from pprint import pprint
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
# w=win32gui

"""fd = w.FindWindow(None, w.GetWindowText(w.GetForegroundWindow()))

w.ShowWindow(fd, 1)

print(w.GetWindowText(w.GetForegroundWindow()))"""
# w.GetWindowText(w.GetForegroundWindow())找到最前方窗口名
# w.FindWindow(类名，窗口名)
# 隐藏
# w.ShowWindow(id,0)
# 显示
# w.ShowWindow(id,1)
soundnum=0
ishide=1
hiden=[]
datayaml="""#请勿使用tab, 请使用空格
whitelist: #当隐藏时不会隐藏的窗口
  class: [Shell_TrayWnd] #按照类查找窗口
  name_strict: [] #按照严格的窗口名查找
  name_loose: [] #按照关键字查找窗口名

blacklist: #哪怕没有被记录也会尝试隐藏的窗口
  class: [] #按照类查找窗口
  name_strict: [简单整理] #按照严格的窗口名查找
  name_loose: [] #按照关键字查找窗口名

whenhideopenlist: #当隐藏时会被全屏打开的窗口
  class: [] #按照类查找窗口
  name_strict: [] #按照严格的窗口名查找
  name_loose: [] #按照关键字查找窗口名

whenhidenosound: true #隐藏时是否静音
whenshowopensound: true #当取消隐藏时是否恢复声音
keys: [shift,ctrl,d] #你的切屏快捷键
keychecktime: 0.05 #多长时间检查一下你的快捷键是否按下
windowchecktime: 1 #多长时间检查你有哪些窗口在使用
"""

try:
    f=open("data.yaml","r",encoding="utf-8")
    f.close()
except FileNotFoundError:
    f=open("data.yaml","w",encoding="utf-8")
    f.write(datayaml)
    f.close()
finally:
    f=open("data.yaml","r",encoding="utf-8")
    data=yaml.safe_load(f.read())
    f.close()
    data["usingwindows"]=[]
#pprint(data["blacklist"])

def iswhitelist(id=None,name=None):
    if id != None:
        if id in data["whitelist"]["class"]:
            return True
        else:
            pass
    if name != None:
        if name in data["whitelist"]["class"]:
            return True
        else:
            for name_loose in data["whitelist"]["name_loose"]:
                if name_loose in name:
                    return True
            return False        

def findmorewindows(id=None,name=None):
    allwindow=[]
    windowid=w.FindWindowEx(0, 0,id,name)
    allwindow.append(windowid)
    while 1:
        if allwindow[-1]==0:
            break
        else:
            windowid=w.FindWindowEx(0, allwindow[-1],id,name)
            allwindow.append(windowid)
    return allwindow
    

def hide():
    global ishide,hiden,soundnum
    ishide=1
    hiden=[]
    def foo(hwnd,mouse):
        titles.add(w.GetWindowText(hwnd))
    if data["whenhidenosound"]:
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(
            IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        volume = cast(interface, POINTER(IAudioEndpointVolume))
        soundnum=volume.GetMasterVolumeLevel()
        nosound(0)
    titles = set()
    w.EnumWindows(foo, 0)
    lt = [t for t in titles if t]
    
    for window in data["usingwindows"]:
        if not iswhitelist(name=window):
            windows=findmorewindows(name=window)
            for win in windows:
                if w.IsWindowVisible(win):
                    if w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMAXIMIZED:
                        hiden.append([win,3])
                    elif w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMINIMIZED:
                        hiden.append([win,2])
                    else:
                        hiden.append([win,1])
                    w.ShowWindow(win,0)
    
    for window in lt:
        for name_loose in data["blacklist"]["name_loose"]:
            if name_loose in window and not iswhitelist(name=window):
                windows=findmorewindows(name=window)
                for win in windows:
                    if w.IsWindowVisible(win):
                        if w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMAXIMIZED:
                            hiden.append([win,3])
                        elif w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMINIMIZED:
                            hiden.append([win,2])
                        else:
                            hiden.append([win,1])
                        w.ShowWindow(win,0)
                
    for window in data["blacklist"]["name_strict"]:
        if not iswhitelist(name=window):
            windows=findmorewindows(name=window)
            for win in windows:
                if w.IsWindowVisible(win):
                    if w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMAXIMIZED:
                        hiden.append([win,3])
                    elif w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMINIMIZED:
                        hiden.append([win,2])
                    else:
                        hiden.append([win,1])
                    w.ShowWindow(win,0)
                
    for window in data["blacklist"]["class"]:
        if not iswhitelist(id=window):
            windows=findmorewindows(id=window)
            for win in windows:
                if w.IsWindowVisible(win):
                    if w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMAXIMIZED:
                        hiden.append([win,3])
                    elif w.GetWindowPlacement(win)[1]==win32con.SW_SHOWMINIMIZED:
                        hiden.append([win,2])
                    else:
                        hiden.append([win,1])
                    w.ShowWindow(win,0)
    
    for window in lt:
        for name_loose in data["whenhideopenlist"]["name_loose"]:
            if name_loose in window:
                windows=findmorewindows(name=window)
                for win in windows:
                    if win!=0:
                        w.ShowWindow(win,3)
                        #w.SetWindowPos(win)
                        win32com.client.Dispatch("WScript.Shell").SendKeys("%")
                        w.SetForegroundWindow(win)
                        #win32api.PostMessage(win, win32con.WM_KEYUP, win32con.VK_CONTROL, 0)
    
    for window in data["whenhideopenlist"]["name_strict"]:
        windows=findmorewindows(name=window)
        for win in windows:
            if win!=0:
                w.ShowWindow(win,3)
                w.BringWindowToTop(win)
            
    for window in data["whenhideopenlist"]["class"]:
        windows=findmorewindows(id=window)
        for win in windows:
            if win!=0:
                w.ShowWindow(win,3)
                w.BringWindowToTop(win)
    print("hideing")
    #按照data隐藏

def show():
    global ishide,hiden,soundnum
    ishide=0
    #print(hiden)
    for window in hiden:
        w.ShowWindow(window[0],window[1])
    if data["whenhidenosound"]:
        nosound(soundnum)
    print("showing")
    #按照data显示

def getwindow():
    while 1:
        windowname=w.GetWindowText(w.GetForegroundWindow())
        if windowname not in data["usingwindows"] and not ishide and windowname!="":
            data["usingwindows"].append(windowname)
        if windowname not in data["whitelist"]["name_strict"] and ishide and windowname!="":
            data["whitelist"]["name_strict"].append(windowname)
        time.sleep(data["windowchecktime"])

def checkkey():
    print("简单整理is running...\nhideing")
    while 1:
        for key in data["keys"]:
            if keyboard.is_pressed(key):
                pass
            else:
                time.sleep(data["keychecktime"])
                break
        else:
            if ishide:
                show()
            else:
                hide()
            time.sleep(1)

def nosound(open):
    devices = AudioUtilities.GetSpeakers()
    interface = devices.Activate(
        IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
    volume = cast(interface, POINTER(IAudioEndpointVolume))
    if not open:
        volume.SetMasterVolumeLevel(volume.GetVolumeRange()[0], None)
    else:
        volume.SetMasterVolumeLevel(open, None)
    return 0

threading.Thread(target=getwindow, args=()).start()
checkkey()
