# 坏枪版权所有。

#初始启动。
version = "坏枪点名器 3.5"
import tkinter as tk
import tkinter.messagebox
import os
import threading
import sys
import win32con
import win32gui
import win32api
hwnd = win32gui.FindWindow(None,version)
try:
    win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
    win32gui.SetForegroundWindow(hwnd)
    sys.exit()
except:
    pass
import random
import time
import ctypes
import xlrd
import pygame
import datetime
time1 = datetime.datetime.now().strftime('%Y-%m-%d %H-%M-%S')
#检测字体。
root = tk.Tk()
import tkinter.font as tkFont
fontlist = tkFont.families()
if '思源黑体 CN Heavy' in fontlist:
    font = '思源黑体 CN Heavy'
else:
    font = '微软雅黑 Bold'
#写入文本。
full_path = time1
def text_create(msg):
    desktop_path = "config/"
    file = open(desktop_path+full_path+'.txt', 'a')
    file.writelines(msg+'\n')
if not os.path.exists("config"):
    os.mkdir("config")
try:
    text_create('以下是本次抽取名单。')
except:
    tkinter.messagebox.showerror(title='错误', message='日志文件创建失败！\n请尝试以管理员身份打开本软件！')
    sys.exit()
#获取屏幕分辨率。
user32 = ctypes.windll.user32
scryfont = win32api.GetSystemMetrics(1)
user32.SetProcessDPIAware()
scry = user32.GetSystemMetrics(1)
print(scry)
print(scryfont)
k2 = scryfont/1440*0.7
k = scry/1440*0.7
print(k)
pygame.mixer.init()
soundfile="sound/chose.wav"
soundfile2="sound/chose2.wav"
soundfile3="sound/passing.wav"
track = pygame.mixer.Sound(soundfile)
track2 = pygame.mixer.Sound(soundfile2)
track2.set_volume(0.6)
track3 = pygame.mixer.Sound(soundfile3)
from PIL import Image
from PIL import ImageTk
try:
    data = xlrd.open_workbook("list/name.xls")
    table = data.sheets()[0]
    namelist=table.col_values(0)
except:
    tkinter.messagebox.showerror(title='错误', message='表格读取错误！\n请检查list/name.xls是否存在并且第一列有内容！')
    sys.exit()
firstlist=[]
secondlist=[]
thirdlist=[]
fourth = 0

#列出单字检测列表。
for tempname in namelist:
    list(tempname)
    firsttemp = tempname[0]
    try:
        secondtemp = tempname[0]+tempname[1]
        secondlist.append(secondtemp)
    except:
        secondtemp = ""
    try:
        thirdtemp = tempname[0]+tempname[1]+tempname[2]
        thirdlist.append(thirdtemp)
    except:
        thirdtemp = ""
    try:
        fourthtemp = tempname[3]
        fourth = 1
    except:
        pass
    firstlist.append(firsttemp)
print(secondlist)
print(thirdlist)

#窗口。
root.title(version)
root.resizable(width=False, height=False)
try:
    root.iconbitmap('logo.ico')
except:
    tkinter.messagebox.showerror(title='错误', message='图标文件丢失！\n请检查logo.ico是否存在！')
    sys.exit()
if fourth == 1:
    canva = tk.Canvas(root,bd=0,bg="MediumBlue",height=1800*k,width=1265*k)
else:
    canva = tk.Canvas(root,bd=0,bg="MediumBlue",height=1800*k,width=1000*k)
canva.config(highlightthickness=0)

#背景处理。
filenamebst = Image.open("image/background.png")
filenamebst = filenamebst.resize((int(1400*k),int(2000*k)),Image.ANTIALIAS)
filenameb = ImageTk.PhotoImage(filenamebst)
imageb = canva.create_image(0,0,anchor='nw',image=filenameb)
filenamest = Image.open("image/logo.png")
filenamest = filenamest.resize((int(500*k),int(200*k)),Image.ANTIALIAS)
filename = ImageTk.PhotoImage(filenamest)
if fourth == 1:
    image = canva.create_image(632.5*k,100*k,image=filename)
else:
    image = canva.create_image(500*k,100*k,image=filename)
fileblockst = Image.open("image/block.png")
fileblockst = fileblockst.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
fileblock = ImageTk.PhotoImage(fileblockst)
block = canva.create_image(75*k,200*k,anchor='nw',image=fileblock)
block2 = canva.create_image(500*k,200*k,anchor='n',image=fileblock)
block3 = canva.create_image(925*k,200*k,anchor='ne',image=fileblock)
if fourth == 1:
    block4 = canva.create_image(1210*k,200*k,anchor='ne',image=fileblock)
filestartst = Image.open("image/start.png")
filestartst = filestartst.resize((int(500*k),int(250*k)),Image.ANTIALIAS)
filestart = ImageTk.PhotoImage(filestartst)

#主程序。
click = 0
loopt = 0
fram = 0
loopt = 0
loop=1
loop2=1
loop3=1
columnFont = (font,int(150*k2))
columnFont2 = (font,int(60*k2))
if fourth == 1:
    fstnames = canva.create_text(int(105*k),int(175*k),font = columnFont,anchor='nw',text="请",tags="fst")
    secondnames = canva.create_text(int(490*k),int(175*k),font = columnFont,anchor='n',text="您",tags="sec")
    thirdnames = canva.create_text(int(880*k),int(175*k),font = columnFont,anchor='ne',text="抽",tags="thi")
    fourthnames = canva.create_text(int(1170*k),int(175*k),font = columnFont,anchor='ne',text="取",tags="for")
else:
    fstnames = canva.create_text(int(105*k),int(175*k),font = columnFont,anchor='nw',text="请",tags="fst")
    secondnames = canva.create_text(int(490*k),int(175*k),font = columnFont,anchor='n',text="抽",tags="sec")
    thirdnames = canva.create_text(int(880*k),int(175*k),font = columnFont,anchor='ne',text="取",tags="thi")
firstname=""
secondname=""
thirdname=""
totalname=""
givenpos = 350*k
times = 0
movespeed = 1.1
checktimes = 0
check1=0
check2=0
check3=0
def chose():
    global fram
    global loop
    global loop2
    global loop3
    global fstnames
    global secondnames
    global thirdnames
    global loopt
    global firstname
    global secondname
    global thirdname
    global totalname
    global givenpos
    global times
    global movespeed
    global checktimes
    global firstlist
    global secondlist
    global thirdlist
    global check1
    global check2
    global check3
    try:
        #第一次检测。
        if loopt == 1:
            track3.play(loops = -1)
            canva.delete("fst")
            canva.delete("sec")
            canva.delete("thi")
            canva.delete("for")
            #刷新列表。
            for i in range(1,10):
                random.shuffle(namelist)
            #抽取名字。
            totalname=str(random.choice(namelist))
            spitname=list(totalname)
            firstname=spitname[0]
            try:
                secondname=spitname[1]
                secondcheck=spitname[0]+spitname[1]
            except:
                secondname=" "
                secondcheck=" "
            try:
                thirdname=spitname[2]
                thirdcheck=spitname[0]+spitname[1]+spitname[2]
            except:
                thirdname=" "
                thirdcheck=" "
            if fourth == 1:
                try:
                    fourthname=spitname[3]
                except:
                    fourthname=" "
        while loopt == 1:
            fram += 1
            if fram < 10:
                loopname = "block_0000"+str(fram)+".png"
                loopname2 = "block2_0000"+str(fram)+".png"
                loopname3 = "block3_0000"+str(fram)+".png"
            else:
                loopname = "block_000"+str(fram)+".png"
                loopname2 = "block2_000"+str(fram)+".png"
                loopname3 = "block3_000"+str(fram)+".png"
            if fram > 67:
                fram = 0
            fileloopst = Image.open("image/loop/"+loopname)
            fileloopst = fileloopst.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
            fileloop = ImageTk.PhotoImage(fileloopst)
            fileloop2st = Image.open("image/loop/"+loopname2)
            fileloop2st = fileloop2st.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
            fileloop2 = ImageTk.PhotoImage(fileloop2st)
            fileloop3st = Image.open("image/loop/"+loopname3)
            fileloop3st = fileloop3st.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
            fileloop3 = ImageTk.PhotoImage(fileloop3st)
            canva.delete("loop")
            loop = canva.create_image(925*k,200*k,anchor='ne',image=fileloop,tags="loop")
            canva.delete("loop2")
            loop2 = canva.create_image(500*k,200*k,anchor='n',image=fileloop2,tags="loop2")
            canva.delete("loop3")
            loop3 = canva.create_image(75*k,200*k,anchor='nw',image=fileloop3,tags="loop3")
            if fourth == 1:
                canva.delete("loop4")
                loop4 = canva.create_image(1210*k,200*k,anchor='ne',image=fileloop3,tags="loop4")
            time.sleep(0.01)
        #第二次循环。
        canva.delete("loop3")
        fstnames = canva.create_text(105*k,175*k,font = columnFont,anchor='nw',text=firstname,tags="fst")
        a=firstlist.count(firstname) > 1
        b=secondname != " "
        check1 = a & b
        if check1:
            track2.play(loops = 0)
            while loopt == 2:
                fram += 1
                if fram < 10:
                    loopname = "block_0000"+str(fram)+".png"
                    loopname2 = "block2_0000"+str(fram)+".png"
                    loopname3 = "block3_0000"+str(fram)+".png"
                else:
                    loopname = "block_000"+str(fram)+".png"
                    loopname2 = "block2_000"+str(fram)+".png"
                    loopname3 = "block3_000"+str(fram)+".png"
                if fram > 67:
                    fram = 0
                fileloopst = Image.open("image/loop/"+loopname)
                fileloopst = fileloopst.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
                fileloop2st = Image.open("image/loop/"+loopname2)
                fileloop2st = fileloop2st.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
                fileloop = ImageTk.PhotoImage(fileloopst)
                fileloop2 = ImageTk.PhotoImage(fileloop2st)
                fileloop3st = Image.open("image/loop/"+loopname3)
                fileloop3st = fileloop3st.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
                fileloop3 = ImageTk.PhotoImage(fileloop3st)
                canva.delete("loop")
                loop = canva.create_image(925*k,200*k,anchor='ne',image=fileloop,tags="loop")
                canva.delete("loop2")
                loop2 = canva.create_image(500*k,200*k,anchor='n',image=fileloop2,tags="loop2")
                if fourth == 1:
                    canva.delete("loop4")
                    loop4 = canva.create_image(1210*k,200*k,anchor='ne',image=fileloop3,tags="loop4")
                time.sleep(0.01)
        canva.delete("loop2")
        secondnames = canva.create_text(490*k,175*k,font = columnFont,anchor='n',text=secondname,tags="sec")
        c=secondlist.count(secondcheck) > 1
        d=thirdname != " "
        f = c & d
        check2 = check1 & f
        if check2:
            track2.play(loops = 0)
            while loopt == 3:
                fram += 1
                if fram < 10:
                    loopname = "block_0000"+str(fram)+".png"
                    loopname3 = "block3_0000"+str(fram)+".png"
                else:
                    loopname = "block_000"+str(fram)+".png"
                    loopname3 = "block3_000"+str(fram)+".png"
                if fram > 67:
                    fram = 0
                fileloopst = Image.open("image/loop/"+loopname)
                fileloopst = fileloopst.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
                fileloop = ImageTk.PhotoImage(fileloopst)
                fileloop3st = Image.open("image/loop/"+loopname3)
                fileloop3st = fileloop3st.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
                fileloop3 = ImageTk.PhotoImage(fileloop3st)
                canva.delete("loop")
                loop = canva.create_image(925*k,200*k,anchor='ne',image=fileloop,tags="loop")
                if fourth == 1:
                    canva.delete("loop4")
                    loop4 = canva.create_image(1210*k,200*k,anchor='ne',image=fileloop3,tags="loop4")
                time.sleep(0.01)
        canva.delete("loop")
        thirdnames = canva.create_text(880*k,175*k,font = columnFont,anchor='ne',text=thirdname,tags="thi")
        if fourth == 1:
            e=fourthname != " "
            check3 = check2 & e
            if check3 :
                track2.play(loops = 0)
                while loopt == 4:
                    fram += 1
                    if fram < 10:
                        loopname3 = "block_0000"+str(fram)+".png"
                    else:
                        loopname3 = "block_000"+str(fram)+".png"
                    if fram > 67:
                        fram = 0
                    fileloop3st = Image.open("image/loop/"+loopname3)
                    fileloop3st = fileloop3st.resize((int(275*k),int(275*k)),Image.ANTIALIAS)
                    fileloop3 = ImageTk.PhotoImage(fileloop3st)
                    canva.delete("loop4")
                    loop4 = canva.create_image(1210*k,200*k,anchor='ne',image=fileloop3,tags="loop4")
                    time.sleep(0.01)
            canva.delete("loop4")
            fourthnames = canva.create_text(int(1170*k),int(175*k),font = columnFont,anchor='ne',text=fourthname,tags="for")
        track.play(loops = 0)
        track3.stop()
        loopt = 0
        
        givenpos += 125*k
        times += 1
        checktimes = times-1
        if checktimes != 0:
            if checktimes %8 == 0:
                givenpos -= 8*125*k
            #if times > 8:
                #while movespeed < 1000:
                    #movespeed **= 1.05
                    #canva.move("showtext"+str(times-6),movespeed,0)
        canva.delete("showtext"+str(times-8))
        if fourth == 1:
            totalnameshow = canva.create_text(632.5*k,givenpos,font=columnFont2,anchor='n',fill="white",text="第"+str(times)+"个幸运儿:"+totalname,tags="showtext"+str(times))
        else:
            totalnameshow = canva.create_text(500*k,givenpos,font=columnFont2,anchor='n',fill="white",text="第"+str(times)+"个幸运儿:"+totalname,tags="showtext"+str(times))
        timenow = datetime.datetime.now().strftime('%H:%M:%S')
        text_create('['+timenow+']'+'第'+str(times)+'个幸运儿：\t'+totalname)
    except:
        sys.exit()
def chosebe():
    global loopt
    loopt += 1
    if loopt == 1:
        t=threading.Thread(target=chose)
        t.start()
butt=tk.Button(text='开始',bd=0,bg='#e9424a',activebackground='#e9424a',image=filestart,relief="groove",command=chosebe)
canva.pack()
if fourth == 1:
    butt.place(x=632.5*k,y=1750*k,anchor="s")
else:
    butt.place(x=500*k,y=1750*k,anchor="s")
#窗口关闭。
def exitall():
    root.destroy()
    track3.stop()
    sys.exit()
root.protocol('WM_DELETE_WINDOW',exitall)
#第一次启动则显示日志。
dirs = os.listdir("config")
if len(dirs) == 1:
    tkinter.messagebox.showinfo(title='欢迎', message='欢迎使用点名器！\n若这是首次使用点名器，请在list文件夹里修改name.xls导入抽奖名单。\n每次的抽取日志可以在config文件夹里找到。')
root.mainloop()
sys.exit()
