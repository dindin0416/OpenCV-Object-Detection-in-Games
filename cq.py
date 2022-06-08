#打挑戰+gui
import re
import tkinter as tk

import cv2
import numpy as np
import win32com
from pythonwin import win32ui
from win10toast import ToastNotifier
from win32 import win32api, win32gui
from win32.lib import win32con


#搜尋窗口
def FindWindow_bySearch(pattern):
    window_list = []
    win32gui.EnumWindows(lambda hwnd, param: param.append(hwnd), window_list)
    for each in window_list:
        if re.search(pattern, win32gui.GetWindowText(each)) is not None:
            return each
#查詢窗口大小
def getWindow_wh(hwnd):
    return win32gui.GetWindowRect(hwnd)

#實時截圖
def getWindow_img(hwnd):
    w = getWindow_wh(hwnd)[2]
    h = getWindow_wh(hwnd)[3]
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    saveBitMap = win32ui.CreateBitmap()
    saveBitMap.CreateCompatibleBitmap(mfcDC, w, h)
    saveDC.SelectObject(saveBitMap)
    saveDC.BitBlt((0, 0), (w, h), mfcDC, (0, 0), win32con.SRCCOPY)
    signedIntsArray = saveBitMap.GetBitmapBits(True)
    img = np.frombuffer(signedIntsArray, dtype='uint8')
    img.shape = (h,w,4)
    #Free resource
    mfcDC.DeleteDC()
    saveDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    win32gui.DeleteObject(saveBitMap.GetHandle())
    return img

#設定按鈕的函式
def btn_start():   #開始按鈕的函式
    start()
def btn_stop():    #暫停按鈕的函式
    win.destroy()

#開始截圖判斷的部分
def start():
    toaster = ToastNotifier()   #初始化windows通知套件
    hwnd = FindWindow_bySearch("BlueStacks")
    while(True):
        img_rgb = getWindow_img(hwnd)
        img_gray = cv2.cvtColor(img_rgb, cv2.COLOR_BGR2GRAY)
        res_retry = cv2.matchTemplate(img_gray,template_retry,cv2.TM_CCOEFF_NORMED)
        loc_retry = np.where( res_retry >= 0.95)
        res_running_test = cv2.matchTemplate(img_gray,template_running_test,cv2.TM_CCOEFF_NORMED)
        loc_running_test = np.where( res_running_test >= 0.95)
        res_paopao = cv2.matchTemplate(img_gray,template_paopao,cv2.TM_CCOEFF_NORMED)
        loc_paopao = np.where( res_paopao >= 0.95)
        #如果在畫面中找到泡泡圖示，代表遇到泡泡，右下角跳出遇到泡泡的通知
        if loc_paopao[0].size & loc_paopao[1].size != 0:
            print('遇到泡泡了!')
            toaster.show_toast("泡泡!",
                    "遇到泡泡了! 停下來買一下競技場票吧!",
                    icon_path=None,
                    duration=10,
                    threaded=True)
        else:
            #如果沒有找到泡泡圖示則代表沒遇到泡泡，
            #進一步確認畫面中有沒有找到跑步測驗圖示，有則代表遇到跑步測驗，右下角跳出遇到跑步測驗的通知
            if loc_running_test[0].size & loc_running_test[1].size != 0:
                toaster.show_toast("跑步測驗!",
                    "遇到跑步測驗ㄌ 你停下來手動按一下名次吧!",
                    icon_path=None,
                    duration=10,
                    threaded=True)
            else:
                #沒有遇到泡泡也沒有遇到跑步測驗，則在畫面中尋找重新開始的按鈕
                if loc_retry[0].size & loc_retry[1].size !=0 :
                    print('重新開始下一場摟')
                    #找到重新開始按紐，發送虛擬點擊指令給目標視窗a
                    long_position = win32api.MAKELONG(384, 384)
                    win32api.SendMessage(a, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, long_position)
                    win32api.SendMessage(a, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, long_position)
        
        cv2.imshow('dindin\'s CQ----auto_rechallenging', img_rgb)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            cv2.destroyAllWindows()
            break

template_running_test = cv2.imread('./image/running_test.jpg',0)            #匯入跑步測試的目標圖片
w_running_test, h_running_test = template_running_test.shape[::-1]          #跑步測試的目標圖片的大小
template_retry = cv2.imread('./image/ch_retry.png',0)                       #匯入重新挑戰的目標圖片
w_retry, h_retry = template_retry.shape[::-1]                               #重新挑戰的目標圖片的大小
template_paopao = cv2.imread('./image/paopao.png',0)                       #匯入遇到泡泡的目標圖片
w_paopao, h_paopao = template_paopao.shape[::-1]                               #處理遇到泡泡的目標圖片

#視窗的部分
win = tk.Tk()              # 建立主視窗物件
win.geometry('640x80')     # 設定主視窗預設尺寸為640x480
win.resizable(False,False) # 設定主視窗的寬跟高皆不可縮放
win.title('dindin\'s CQ')  # 設定主視窗標題
hwnd = FindWindow_bySearch("BlueStacks")    #找顯示的窗口
handle = win32gui.FindWindow(None, "BlueStacks") 
a = win32gui.FindWindowEx(handle, 0, "WindowsForms10.Window.8.app.0.1ca0192_r6_ad1", None)  #找要接收模擬點擊的窗口
win32gui.MoveWindow(hwnd, 0, 0, 500, 500, True)     #把截圖窗口固定到左上角
shell = win32com.client.Dispatch("WScript.Shell")       
shell.SendKeys('%')























win32gui.SetForegroundWindow(hwnd)      #把主窗口移到最前面

#設定按鈕樣式  
btn_start = tk.Button(text='START', command = btn_start, bg='white', fg='Black')
btn_stop = tk.Button(text='STOP', command = btn_stop, bg='gray', fg='Black')
#設定按鈕位置
btn_start.place(rely=0.5, relx=0.3, anchor='center')
btn_stop.place(rely=0.5, relx=0.7, anchor='center')
win.mainloop()







#     #catch-----------------------------------------------------------
#     res_start = cv2.matchTemplate(img_gray,template_start,cv2.TM_CCOEFF_NORMED)
#     loc_start = np.where( res_start >= 0.95)
#     print(loc_start)
#     for pt in zip(*loc_start[::-1]):
#         cv2.rectangle(img_rgb, pt, (pt[0] + w_start, pt[1] + h_start), (0,255,255), 2)
#     #catch-----------------------------------------------------------
#     cv2.imshow('CQ', img_rgb)
#     if cv2.waitKey(25) & 0xFF == ord('q'):
#         cv2.destroyAllWindows()
#         break
    
#     keyboard.start_recording()
#     time.sleep(3)
#     keyboard_events = keyboard.stop_recording()
#     for i in keyboard_events :
#         print(i)
