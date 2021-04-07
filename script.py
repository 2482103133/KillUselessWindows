# -*- coding: utf-8 -*-
from win32gui import *
from win32process import *
from datetime import datetime
from time import sleep
import psutil
import PySimpleGUI as sg
import win32con  
import json
import ctypes, sys
import os
import sys
import re
sys.stdout.reconfigure(encoding='utf-8')
#配置文件路径

config_path=os.path.abspath("config.json")

print("Killer starting...")
# 读取配置文件
def read_config():
    with open(config_path,"r",encoding="utf-8") as f:
        config=json.load(f)
        
        return config
def save_config(config):
    with open(config_path,"w",encoding="utf-8") as f:
        json.dump(config,f,ensure_ascii=False, indent=4, sort_keys=True)
#待关闭列表
programs={}
#所有的窗口hwnd
allPrograms=[]

config=read_config()


#遍历所有窗口
def foo(hwnd,mouse):
    allPrograms.append(hwnd)

    #将所有最小化的窗口加入待关闭列表
    if IsWindow(hwnd) and IsWindowEnabled(hwnd) and  IsIconic(hwnd) :
        if not programs.__contains__(hwnd) :
            
            tid, pid=GetWindowThreadProcessId(hwnd)
            # print(pid)
            programs[hwnd]={"name":GetWindowText(hwnd),"className":GetClassName(hwnd),"hwnd":str(hwnd),"iconicTime":datetime.now(),"pid":pid,"process":os.path.basename(psutil.Process(pid).exe())}
    elif IsWindow(hwnd) and IsWindowEnabled(hwnd)  and  (IsWindowVisible(hwnd) or not IsIconic(hwnd)):
        if programs.__contains__(hwnd):
            del programs[hwnd]



#构造程序弹窗,提醒关闭闲置进程
def pop(idles):
    
    if len(idles)==0:
        return

    cbs=[[sg.Text('Choose to close')]]
    row=[]
    config=read_config()
    exceptions=config["exceptions"]
    ifNoSelection=True
    checked=[]
    for idle in idles:
        match=False
        for exp in config["exception_rules"]:
            if(re.match(exp,idle["name"]+"[@]"+idle["className"]+"[@]"+idle["process"])!=None):
                match=True
                break

        unselected=((idle["name"]  in exceptions and idle["className"]==exceptions[idle["name"]]) or  match  )
        if not unselected:
            ifNoSelection=False
        
        cb=sg.Checkbox(idle['process']+"/"+idle["name"], default=not unselected,key="hwnd-"+str(idle["hwnd"]),tooltip=idle["className"]+"  "+idle["process"])
    
        if unselected:
            checked.append(cb)
        else:
            checked.insert(0,cb)
    
    for cb in checked:
        row.append(cb)
        if len(row)==2:
            cbs.append(row)
            row=[]
    if(len(row)>0):
        cbs.append(row)
    
    if(ifNoSelection and config["diable_popup_when_no_recommended_selection"] ):
        return

    cbs.append([sg.Button("Ok"),sg.Button("Cancel"),sg.Button("Quit"),sg.Checkbox("Diable popup when no recommended", default=config["diable_popup_when_no_recommended_selection"],key="disble_popup"),sg.Checkbox("Add unselected to exceptions", default=True,key="add_exceptions")])
    
    window=sg.Window('Found some idle windows for you!', layout=cbs,size=(800, 400),resizable=True)
    event, values = window.read(close=True)
    print("window closed")
    if event == 'Ok' :
        config=read_config() 
        exections={}
        for key in values:
            
            if(key.startswith("hwnd")  ):
                handle=int(key.split("-")[1])
                ex=next(x for x in idles if x["hwnd"]==str(handle))
                #关闭所有被勾选的窗口,并将未勾选的窗口加入到exception列表
                if values[key] :
                    try:
                        print("Killing "+str(ex))
                        
                        PostMessage(handle,win32con.WM_CLOSE,0,0)
                        PostMessage(handle,win32con.WM_QUIT,0,0)
                        PostMessage(handle,win32con.WM_ENDSESSION,0,0)
                        if(ex["name"] in config["exceptions"] and  ex["className"]==config["exceptions"][ex["name"]]):
                            del config["exceptions"][ex["name"]]
                        
                    except Exception as e:
                        print(e)

                else:
                    
                    exections[ex["name"]]=ex["className"]


                  
        #判断是否将未勾选的列表加入到例外列表中, 以后将自动取消勾选位于例外列表中的窗口
        if(values["add_exceptions"]):
            l=config["exceptions"]
            l.update(exections)
            config["exceptions"]=l

        config["diable_popup_when_no_recommended_selection"]=values["disble_popup"]

        save_config(config)
        
    elif event == 'Cancel':
        for key in values:
            
            if(key.startswith("hwnd")  ):

                #将勾选的但是被取消的项目从idle列表中移除
                if values[key] :
                    try:
                        handle=int(key.split("-")[1])
                        ex=next(x for x in idles if x["hwnd"]==str(handle))
                        idles.remove(ex)
                        
                    except Exception as e:
                        print(e)

    elif event == 'Exceptions':
        print(values)
    elif event == 'Quit':
        sys.exit('User ask to quit')


#每隔五秒统计所有窗口最小化闲置的时间, 将闲置时间大于指定秒数的窗口拉出来供你枪毙
while True:
    allPrograms=[]
    EnumWindows(foo, 0)
    now=datetime.now()
    idles=[]
    kets_to_remove=[]
    for key in programs.keys():
        
        program=programs[key]
        if int(program["hwnd"]) not in allPrograms:
            continue

        span=(now-program["iconicTime"]).total_seconds()
        # print(str(span)+":"+str(programs[key]))
        
        if span>=config["max_idle_time"]:
            idles.append(program)
    if(key not in allPrograms):
        kets_to_remove.append(key)
    for key in kets_to_remove:
        del programs[key]
    pop(idles)
    print("total:"+str(programs.keys().__len__()))
    sleep(config["test-interval"])
    


    