import sqlite3
import psutil
import time
from tkinter import *
import win32api
import win32con
from win32com.client import Dispatch
from tkinter import messagebox

con = sqlite3.connect('D:/Application.db')
cur = con.cursor()
sql = '''create table if not exists test1(
        id INTEGER PRIMARY KEY NOT NULL,
        name TEXT NOT NULL,
        version TEXT NOT NULL
        )'''
cur.execute(sql)
print("数据库初始化成功……")
con.commit()
con.execute("INSERT OR IGNORE INTO test1(id,name,version)\
		VALUES(1,'chrome.exe','92.0.4515.159')")
con.execute("INSERT OR IGNORE INTO test1(id,name,version)\
		VALUES(2,'WeChat.exe','3.3.0.0')")
con.execute("INSERT OR IGNORE INTO test1(id,name,version)\
		VALUES(3,'RPAStudio.exe','2021.1.1.98')")
con.commit()
print('数据写入成功')
con.close()

num = 0  # num用于存放循环检测次数


# 函数定义
def sleeptime(hour, min, sec):
    return hour * 3600 + min * 60 + sec


# 设置自动执行间隔时间
second = sleeptime(0, 0, 5)
# 死循环
while 1 == 1:
    # 延时
    time.sleep(second)

    num=num+1
    print('第  '+str(num)+'  次检测结果：')
    # 获取进程路径
    def get_exe(pname):
        for proc in psutil.process_iter():
            if proc.name() == pname:
                return proc.exe()


    # 获取程序版本号
    def get_version_via_com(filename):
        parser = Dispatch("Scripting.FileSystemObject")
        version = parser.GetFileVersion(filename)
        return version


    # 获取数据库应用列表数据
    con = sqlite3.connect('D:/Application.db')
    cur = con.cursor()
    results = cur.execute("SELECT * from test1")
    name = []  # 存放从数据库中导出的应用程序名
    application = []  # 二维数组存放从数据库表中读取的每行数据
    for app in results.fetchall():
        # print(app)
        app = list(app)
        application.append(app)
        name.append(app[1])
    con.commit()
    con.close()
    '''
    # 获取程序版本号并与数据库内版本号对比
    for n in application:
        path = get_exe(n[1])
        i = get_version_via_com(path)
        # 判断版本号是否一致
        if i == n[2]:
            print(n[1] + '   版本号无异常')
        else:
            win32api.MessageBox(0, '检测到 ' + n[1] + " 版本号发生改变", "提醒", win32con.MB_ICONWARNING)
    '''
    # 检测本机正在运行的程序列表
    pids = psutil.pids()
    process_now = []
    for pid in pids:
        p = psutil.Process(pid)
        process_name = p.name()
        process_now.append(process_name)

    # 将本机运行的程序列表和数据库中的程序列表对比，获取没有运行的程序
    process_down = set(name).difference(set(process_now))
    if len(process_down) == 3:
        win32api.MessageBox(0, "未检测到正在运行的程序", "提醒", win32con.MB_ICONWARNING)
    process_up = [val for val in name if val in process_now]
    print('--------------检测到以下程序正在运行--------------')
    for run in process_up:
        print(run)
    print('---------------程序版本号检测情况--------------')
    # 获取程序版本号并与数据库内版本号对比

    for n in application:
        for r in process_up:
            if n[1] == r:
                path = get_exe(n[1])
                i = get_version_via_com(path)
                # 判断版本号是否一致
                if i == n[2]:
                    print(n[1] + '   版本号无异常   '+i)
                else:
                    win32api.MessageBox(0, '检测到 ' + n[1] + " 版本号发生改变", "提醒", win32con.MB_ICONWARNING)



    print('--------------以下程序未检测到运行状态--------------')
    for d in process_down:
        print(d)
    print('-----------------------------------------------------------------------------------------------------------')
