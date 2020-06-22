from socket import *
from pyperclip import copy,paste
from ctypes import *
from datetime import datetime,timedelta
import pyautogui, time, os, datetime
from threading import Thread
import inspect
import ctypes


def _async_raise(tid, exctype):
    """raises the exception, performs cleanup if needed"""
    tid = ctypes.c_long(tid)
    if not inspect.isclass(exctype):
        exctype = type(exctype)
    res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
    if res == 0:
        raise ValueError("invalid thread id")
    elif res != 1:
        # """if it returns a number greater than one, you're in trouble,
        # and you should call it again with exc=NULL to revert the effect"""
        ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
        raise SystemError("PyThreadState_SetAsyncExc failed")
def stop_thread(thread):
    _async_raise(thread, SystemExit)

# 自动登陆操作
def auto_Gui():
    """
    auto login name

    """
    pyautogui.click(x=633, y=449, clicks=1, button='left')
    time.sleep(1)
    pyautogui.click(x=1000, y=687, clicks=1, button='left')
    judge = color(990,330,[255,255,255])
    # print(a)
    if judge == None:
        restart()
        time.sleep(2)
        auto_Gui()
        return
    pyautogui.hotkey('ctrl','v')
    signIn()

# 签到操作
def signIn():
    """
    Check whether the first name has been entered
    And click sign in

    """
    time.sleep(0.5)
    pyautogui.moveTo(898, 460,duration=0.25)
    pyautogui.click()
    time.sleep(0.8)
    pyautogui.moveTo(386, 515,duration=0.25)
    pyautogui.click()
    judge = color(1355,116,[255,0,0])
    if judge == None:
        signIn()

# 检测是否进入某一界面是返回1，否返回None
def color(x,y,col):
    """
    :param x: 检查点的x坐标
    :param y:检查点的y坐标
    :param col:检查点的颜色RGB
    :return:返回是否进入到判断的界面
    """
    # print(x,y,col)
    time.sleep(0.5)
    startTime = datetime.datetime.now()
    gdi32 = windll.gdi32
    user32 = windll.user32
    hdc = user32.GetDC(None)
    while True:
        time.sleep(0.5)
        pixel = gdi32.GetPixel(hdc, x, y)
        r = pixel & 0x0000ff
        g = (pixel & 0x00ff00) >> 8
        b = pixel >> 16
        # print(r,g,b)
        if r==col[0] and g==col[1] and b==col[2]:
            # print('进入界面',r,g,b)
            return 1
        if (startTime + datetime.timedelta(seconds=8)) < datetime.datetime.now():
            return False

# 重启客户端
def restart():
    # print('restart client')
    os.system('taskkill /F /IM WzhyClient.exe')
    time.sleep(1)
    os.chdir(r'C:\Users\roinfo\Desktop\work\WzhyClient')
    os.system('start WzhyClient.exe')

# 选择服务器
def choice_server(server):
    if server == 'server1':
        pyautogui.click(x=437, y=405, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=633, y=449, clicks=1, button='left')
    if server == 'server2':
        pyautogui.click(x=665, y=405, clicks=1, button='left')
        time.sleep(1)
        pyautogui.click(x=633, y=449, clicks=1, button='left')
    # 返回水牌名字
    if server == 'business':
        pyautogui.click(x=200, y=90, clicks=1, button='left')
    if server == 'exit':
        os._exit(0)

# 刷新会议资料和恢复水牌
def up_data():
    pyautogui.click(x=177, y=529, clicks=1, button='left')
    time.sleep(0.5)
    pyautogui.click(x=1211, y=138, clicks=1, button='left')

def _copy(message, cp_now):

    if cp_now == message.decode():
        judge = color(1355,116,[255,0,0])
        if judge == 1:
            print('已经登陆成功')
            pass
        else:
            print('还没登陆成功')
            pass
        return cp_now
    else:
        copy(message.decode())
        auto_Gui()
        return message.decode()

# 接收服务器的操作指令
def socket_client():
    # 创建套接字UDP
    s = socket(AF_INET,SOCK_DGRAM)
    # 设置套接字可以发送接收广播
    s.setsockopt(SOL_SOCKET,SO_BROADCAST,1)
    # 固定接收端口
    s.bind(('0.0.0.0',9999))
    # 循环接收服务端的签到信息
    a = ' '
    while True:
        try:
            msg,addr = s.recvfrom(64)
            if msg.decode() in ['server1','server2','business','exit']:
                choice_server(msg.decode())
                continue
            if msg.decode() == 'refresh':
                up_data()
                continue
            elif msg.decode() != None:
                a = _copy(msg,a)
            print("从{}获取信息:{}".format(addr,msg.decode()))
        except (KeyboardInterrupt,SyntaxError):
            raise
        except Exception as e:
            print(e)
    s.close()

if '__main__' == __name__:
    if (datetime.datetime.now() > datetime.datetime(2020,9,1,00,00)):
        os._exit(0)
    socket_client()