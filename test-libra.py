# -*- coding: utf-8 -*-
# =====================
# 作者：Rivendell2898
# 2022年4月9日
# github项目：MO_loading
# url：https://github.com/Rivendell2898/MO_loading
# 此代码不可商用，如有修改请开源
# =====================
import os
# 设置VLC库路径，需在import vlc之前
import win32gui
import win32con

os.environ['PYTHON_VLC_MODULE_PATH'] = "./vlc-3.0.16"
import vlc
import subprocess
import tkinter
import tkinter.messagebox
import shutil


import time

import win32file
from win32file import *

from PyQt5 import QtMultimedia
from PyQt5.QtCore import QUrl

# 载入播放音频
file = QUrl.fromLocalFile('libra\BullyKit.wav')  # 音频文件路径
content = QtMultimedia.QMediaContent(file)
wavplayer = QtMultimedia.QMediaPlayer()
wavplayer.setMedia(content)
wavplayer.setVolume(100)

# 载入完成音频
file1 = QUrl.fromLocalFile('gamecreated.wav')  # 音频文件路径
content1 = QtMultimedia.QMediaContent(file1)
wavplayerfin = QtMultimedia.QMediaPlayer()
wavplayerfin.setMedia(content1)
wavplayerfin.setVolume(100)


def is_open(filename):
    try:
        # 首先获得句柄
        vHandle = win32file.CreateFile(filename, GENERIC_READ, 0, None, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, None)
        # 判断句柄是否等于INVALID_HANDLE_VALUE
        if int(vHandle) == INVALID_HANDLE_VALUE:
            # print("# file is already open")
            return True  # file is already open
        else:
            # print("file is not open")
            return False
        win32file.CloseHandle(vHandle)

    except Exception as e:
        # print(e)
        return True

# 视频播放器
class Player:
    '''
        args:设置 options
    '''
    def __init__(self, *args):
        if args:
            instance = vlc.Instance(*args)
            self.media = instance.media_player_new()
        else:
            self.media = vlc.MediaPlayer()

    # 设置待播放的url地址或本地文件路径，每次调用都会重新加载资源
    def set_uri(self, uri):
        self.media.set_mrl(uri)

    # 播放 成功返回0，失败返回-1
    def play(self, path=None):
        if path:
            self.set_uri(path)
            return self.media.play()
        else:
            return self.media.play()

    def stop(self):
        return self.media.stop()

    # 释放资源
    def release(self):
        return self.media.release()

    # 是否正在播放
    def is_playing(self):
        return self.media.is_playing()

    # 已播放时间，返回毫秒值
    def get_time(self):
        return self.media.get_time()

    # 音视频总长度，返回毫秒值
    def get_length(self):
        return self.media.get_length()

    # 当前播放进度情况。返回0.0~1.0之间的浮点数
    def get_position(self):
        return self.media.get_position()

    def set_time(self, ms):
        return self.media.set_time(ms)

    # 获取当前文件播放速率
    def get_rate(self):
        return self.media.get_rate()

    def set_fullscreen(self):
        return self.media.set_fullscreen(1)

    def video_set_logo_string(self,str):
        return self.media.video_set_logo_string(str)

    def get_hwnd(self):
        return self.media.get_hwnd()

    # 设置宽高比率（如"16:9","4:3"）
    def set_ratio(self, ratio):
        self.media.video_set_scale(0)  # 必须设置为0，否则无法修改屏幕宽高
        self.media.video_set_aspect_ratio(ratio)

def show_err():
    root = tkinter.Tk()
    root.withdraw()
    tkinter.messagebox.showerror('这就是传说中的弹窗', '打开MO Client出错，请尝试用管理员权限打开。\n如果依然报错，那么可能是客户端文件被占用，或安装路径错误，也有可能是程序bug...\n请尝试手动打开MO Client')

def show_info():
    root = tkinter.Tk()
    root.withdraw()
    flag = tkinter.messagebox.askyesno('你真的要运行MO吗', '检测到MO正在运行，建议关闭前台或后台进程后再次打开。真的要强行打开MO客户端吗？')
    if flag == 0:
        exit(0)


if "__main__" == __name__:
    # 初始化时间和音频监听
    start_time = time.time()

    # 初始化
    player = Player()
    player.set_ratio("16:9")
    player.set_fullscreen()

    # 防止MO CLIENT 后台运行
    try:
        os.system('taskkill /f /im %s' % 'clientdx.exe')
    except:
        show_info()
        pass

    # 替换风格bgm
    try:
        os.unlink("../Resources/chaoticimpulse.wma")
        # print("delete wma complete")
    except:
        pass
    shutil.copy('libra\chaoticimpulse.wma', r'../Resources/chaoticimpulse.wma')
    # print("copy wma complete")

    time.sleep(0.1)

    # 打开MO
    try:
        file = subprocess.Popen("../Resources/clientdx.exe")
        pass
    except Exception as e:
        show_err()
        file.kill()

    # 定义窗口
    FrameClass = "WindowsForms10.Window.8.app.0.1ca0192_r6_ad1"
    FrameTitle = "MO Client"
    # hwnd = win32gui.FindWindow(FrameClass, FrameTitle)
    hwnd = win32gui.FindWindow(None, FrameTitle)

    playerClass = "IME"
    playerTile = "Default IME"
    playerhwnd = win32gui.FindWindow(None, playerTile)

    i = 0
    j = 1
    # 最小化窗口并播放视频
    while i < 1000:
        if hwnd:
            if j == 1:
                time.sleep(0.01) #防止窗口未展开就被最小化
                player.play("libra\libra.mp4")
            flag = win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
            # flag = 1
            if flag:
                # print("MO_flag:", flag)
                # time.sleep(0.01)
                break
            # print(hwnd)
            time.sleep(0.01)
            # print("found")
            i = i + 1
            j = 0
            # break
        else:
            hwnd = win32gui.FindWindow(FrameClass, FrameTitle)
            time.sleep(0.01)
            # print("wait")
            i = i + 1

    # 如果超时，则报错
    if i >= 1000:
        show_err()
        exit(0)

    j = 1
    flag = 0
    cnt = 0
    wavplayer.play()

    # 防止player进程退出
    while True:
        time.sleep(0.1)  # 放在此处，防止while循环过快而死机
        end_time = time.time()

        # 如果player的窗口未找到
        if playerhwnd == 0:
            playerhwnd = win32gui.FindWindow(None, playerTile)
        # 如果窗口找到
        # elif j:
        else:
            # time.sleep(0.1)
            # print("playerhwnd:", playerhwnd)
            try:
                win32gui.BringWindowToTop(playerhwnd)
                win32gui.SetForegroundWindow(playerhwnd)
                j = 0
            except Exception as e:
                pass

        # 播放视频用
        if player.get_position() >= 0.95:
            # print("shut down-out of time")
            wavplayer.stop()
            time.sleep(1)
            win32gui.CloseWindow(playerhwnd)
            break

        if (end_time - start_time) >= 10:
            # print("shut down-out of time")
            wavplayer.stop()
            time.sleep(1)
            win32gui.CloseWindow(playerhwnd)
            break

        if cnt % 20 == 0 and cnt > 1:
            if is_open("../Resources/OptionsWindow.ini") != True :
                wavplayer.stop()
                wavplayerfin.play()
                player.set_time(13000) #关闭窗口时视频必须仍在播放，否则会死机
                time.sleep(1)
                wavplayerfin.stop()
                time.sleep(1)
                # print("shut down-ready")
                win32gui.CloseWindow(playerhwnd)
                break
        cnt = cnt + 1


        # 这里将player强制放到 顶层 尝试效果
        # BringWindowToTop
        # CloseWindow
        # DestroyWindow
        # SetForegroundWindow

    # print("run time = ", (end_time - start_time))
    # print("cnt:", cnt)
    # print("show window")
    # win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)   # SW_SHOWMAXIMIZED
    time.sleep(1)
    win32gui.ShowWindow(hwnd, win32con.SW_NORMAL)
    # print("finish")

    # 防止死机
    try:
        win32gui.DestroyWindow(playerhwnd)
    except:
        pass
