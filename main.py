import tkinter as tk
from tkinter import messagebox
import http.client
import pymysql
import datetime
import time
import os
import xlwt
import winreg

'''-----------------------窗口------------------------'''
window = tk.Tk()
window.title('C103考勤系统')
# 获取屏幕 宽、高
ws = window.winfo_screenwidth()
hs = window.winfo_screenheight()
# 计算 x, y 位置
x = (ws/2) - (300/2)
y = (hs/2) - (280/2)
window.geometry('%dx%d+%d+%d' % (300, 280, x, y))
window.resizable(False, False)          #禁用最大化和调整窗口
#window.iconbitmap(r'F:\图标\河马图标\河马-32.ico')

'''---------------------------------------------------'''


'''-------------------获取网络时间---------------------'''
def date_time():
    try:
        conn = http.client.HTTPConnection('192.168.11.75')
        conn.request("GET", "/")
        r = conn.getresponse()
        #r.getheaders() #获取所有的http头
        ts = r.getheader('date')  # 获取http头date部分
        #将GMT时间转换成北京时间
        ltime = time.strptime(ts[5:25], "%d %b %Y %H:%M:%S")
        ttime = time.localtime(time.mktime(ltime)+8*60*60)
        dat = "%u-%02u-%02u" % (ttime.tm_year, ttime.tm_mon, ttime.tm_mday)
        tm = " %02u:%02u:%02u" % (ttime.tm_hour, ttime.tm_min, ttime.tm_sec)
        dt = dat + tm
        #print(dt)
        return dt
    except:
        return '无法连接服务器'
'''---------------------------------------------------'''


'''----------------------时间标签----------------------'''
def update_time():  #刷新时间
    currentTime = date_time()
    label.config(text=currentTime)
    window.update()
    label.after(1000, update_time)

if date_time() == '无法连接服务器':
    print('888')
    label = tk.Label(text='无法连接服务器',font=('黑体', 12))
    label.place(x=150, y=20, anchor='center')
else:
    print('000')
    label = tk.Label(text=date_time(), font=('黑体', 12))
    label.place(x=150, y=20, anchor='center')
    label.after(1000, update_time)


'''----------------------------------------------------'''


'''-----------------------输入框------------------------'''
label2 = tk.Label(text='姓名', font=('黑体', 20))
label2.place(x=45, y=70, anchor='nw')

e = tk.Entry(window, width=8, font=('黑体', 22))
e.place(x=255, y=70, anchor='ne')
'''----------------------------------------------------'''



'''-----------------------按钮功能-------------------------'''
def come():     #上班
    name = e.get()
    if name == '':
        tk.messagebox.showwarning(title='警告', message='请输入姓名')
    else:
        try:
            if user_name(name):
                tk.messagebox.showwarning(title='请注册', message='您尚未注册')
            else:
                if not_again(name, 'clock_in'):
                    da = '"'+name+'"' + ','+'"'+date_time()+'"'
                    write_sql('clock_in', da)
                    mes = name + '签到时间为：' + date_time()
                    tk.messagebox.showinfo(title='上班', message=mes)
                else:
                    tk.messagebox.showwarning(title='警告', message='您已经签到')
        except:
            tk.messagebox.showerror(title='错误', message='无法连接数据库')

def leave():        #下班
    name = e.get()
    if name =='':
        tk.messagebox.showwarning(title='警告', message='请输入姓名')
    else:
        try:
            if user_name(name):
                tk.messagebox.showwarning(title='请注册', message='您尚未注册')
            else:
                if not_again(name,'clock_out'):
                    da = '"'+name+'"' + ','+'"'+date_time()+'"'      
                    write_sql('clock_out',da)
                    mes = name + '签到时间为：' + date_time()
                    tk.messagebox.showinfo(title='下班', message=mes)
                else:
                    tk.messagebox.showwarning(title='警告', message='您已经签到')
        except:
            tk.messagebox.showerror(title='错误', message='无法连接数据库')


def reg():          #注册
    name = e.get() 
    if name == '':
        tk.messagebox.showwarning(title='警告', message='请输入姓名')  
    else:
        if all('\u4e00' <= char <= '\u9fff' for char in name):          #是中文返回Ture
            try:              
                if reg_sql(name, date_time()):
                    tk.messagebox.showwarning(title='警告', message='您已经注册')
                else:
                    tk.messagebox.showinfo(title='注册', message='您已成功注册')
            except:
                tk.messagebox.showerror(title='错误', message='无法连接数据库') 

            
        else:
            tk.messagebox.showwarning(title='警告', message='请输入正确的姓名')


def view(): #生成表格
    try:
        filename = 'c103签到表' + '.xls'
        wbk = xlwt.Workbook()

        sheet1 = wbk.add_sheet('上班', cell_overwrite_ok=True)
        sheet2 = wbk.add_sheet('下班', cell_overwrite_ok=True)
        fileds = ['姓名', '日期']

        results = sql_viewsql('select name,date from clock_in')
        results2 = sql_viewsql('select name,date from clock_out')

        for i in range(0, len(fileds)):
            sheet1.write(0, i, fileds[i])

        for row in range(1, len(results)+1):
            for col in range(0, len(fileds)):
                sheet1.write(row, col, results[row-1][col])

        for i in range(0, len(fileds)):
            sheet2.write(0, i, fileds[i])

        for row in range(1, len(results2)+1):
            for col in range(0, len(fileds)):
                sheet2.write(row, col, results2[row-1][col])

        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders')
        desk = winreg.QueryValueEx(key, "Desktop")[0]
        f = desk + '\\'+filename

        wbk.save(f)

        tk.messagebox.showinfo(title='生成表格', message='表格已保存到桌面')
    except:
        tk.messagebox.showerror(title='生成表格', message='表格生成失败')

'''------------------------------------------------------'''


'''-------------------------按钮--------------------------'''
b_come = tk.Button(window, text='上班', width=15, heigh=2, command=come)
b_come.place(x=15, y=150, anchor='nw')

b_leave = tk.Button(window, text='下班', width=15, heigh=2, command=leave)
b_leave.place(x=285, y=150, anchor='ne')

b_reg = tk.Button(window, text='注册', width=15, heigh=2, command=reg)
b_reg.place(x=15, y=200, anchor='nw')

b_view = tk.Button(window, text='记录', width=15, heigh=2, command=view)
b_view.place(x=285, y=200, anchor='ne')
'''----------------------------------------------------'''

'''-------------------------MySql-------------------------'''


def write_sql(time,da):     #写入签到数据
    db = pymysql.connect(host='192.168.11.75', port=3306,user='root', passwd='root', db='c103')
    cursor = db.cursor()
    sql = 'insert into '+time+'(name,date) values(' + da + ')'
    cursor.execute(sql)
    db.commit()
    cursor.close()
    db.close()

def reg_sql(na,da):         #注册用户
    db = pymysql.connect(host='192.168.11.75', port=3306,user='root', passwd='root', db='c103')
    if user_name(na):
        cursor = db.cursor()
        sql = 'insert into user(name,date) values('+'"'+na+'"'+','+'"'+da+'"'+')'
        cursor.execute(sql)
        db.commit()
        cursor.close()
        db.close()
    else:
        return True


def user_name(name):        #检查是否存在此用户
    db = pymysql.connect(host='192.168.11.75', port=3306,user='root', passwd='root', db='c103')
    cursor = db.cursor()
    sql = 'SELECT * FROM user WHERE name = '+'"'+name+'"'
    cursor.execute(sql)
    results = cursor.fetchone()
    cursor.close()
    db.close()
    if results is None:
        return True
    else:
        return False


def not_again(name,clock):      #防止一个人重复签到
    db = pymysql.connect(host='192.168.11.75', port=3306,user='root', passwd='root', db='c103')
    cur = db.cursor()  # 获取一个游标
    sql = "SELECT * FROM "+clock+" WHERE name='" +name+"' ORDER BY "+clock+".date DESC"  # 定义查询
    cur.execute(sql)  # 执行查询
    date = cur.fetchone()  # 获取查询到的时间
    db.commit()  # 提交事务
    cur.close()  # 关闭游标
    db.close()  # 释放数据库资源在这里插入代码片
    if date is None:            #第一次签到时查询不到上次记录，直接返回进行签到
        return True
    history_time = datetime.datetime.strptime(date[1], '%Y-%m-%d %H:%M:%S')  # 格式化时间
    now_time = datetime.datetime.strptime(date_time(), '%Y-%m-%d %H:%M:%S')
    day = (now_time - history_time).days  # 和上一次签到的天数差
    sec = (now_time - history_time).seconds  # 和上一次签到的秒数差
    if day == 0:
        if sec > 3600:
            return True
        else:
            return False
    else:
        return True


def sql_viewsql(sql):
    conn = pymysql.connect(host='192.168.11.75', port=3306,user='root', passwd='root', db='c103', charset='utf8')
    cursor = conn.cursor()
    cursor.execute(sql)
    results = cursor.fetchall()
    cursor.close()
    conn.close()
    return results
    

'''------------------------------------------------------'''




window.mainloop()
