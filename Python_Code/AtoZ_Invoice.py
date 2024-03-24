from tkinter import *
from tkinter import ttk
from tkinter import messagebox
import os
import copy
import sqlite3 
from datetime import datetime,date,timedelta
import configparser
from xlsxtpl.writerx import BookWriter
from PIL import Image, ImageTk
import customtkinter as ctk
import time
import sqlite3
from calendar import monthrange
from win32com import client
from win32com.client import Dispatch
from win32api import GetSystemMetrics
from tkinter import filedialog
# import pandas as pd
ctk.set_appearance_mode('Dark')
ctk.set_default_color_theme('blue')
class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + round(self.widget.winfo_rootx() + 55*kW)
        y = y + cy + round(self.widget.winfo_rooty() - 50*kH)
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        label = Label(tw, text=self.text, justify=LEFT,
                      background="#ffffe0", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def CreateToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

def addPositionData():
    if position.get()=='' or rank.get()=='':
        return
    else:
        conn = sqlite3.connect(pathGlobal+'/data.db')
        c = conn.cursor()
        c.execute("""INSERT INTO positions VALUES(:position,:rank,:type)""", { 'position':position.get(),'rank':varR.get(),'type':varTC.get()})
        conn.commit()
        conn.close()
        defaultData('positions')

def addPersonData():
    if name.get()=='' or crewID.get()=='':
        return
    else:
        conn = sqlite3.connect(pathGlobal+'/data.db')
        c = conn.cursor()
        c.execute("""INSERT INTO personalDetails VALUES( :crewID,:crewName,:phoneNumber,:mail)""", {'crewID':crewID.get(), 'crewName':name.get(),'phoneNumber':phone.get(),'mail':mail.get()})
        conn.commit()
        conn.close()
        defaultData('personalDetails')

def addVesselData():
    if name.get()=='':
        return
    else:
        conn = sqlite3.connect(pathGlobal+'/data.db')
        c = conn.cursor()
        c.execute("""INSERT INTO vessels VALUES(:vessel,:company)""", { 'vessel':name.get(),'company':company.get()})
        conn.commit()
        conn.close()
        defaultData('vessels')

def onDoubleClick(table):
    item = trv.item(trv.focus())
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    if table == 'personalDetails':
        nameE.delete(0,END)
        nameE.insert(0,item['values'][2])
        global crewID
        crewID = item['values'][0]
    elif table == 'positions':
        positionE.delete(0,END)
        positionE.insert(0,item['values'][1])
        global positionID
        positionID = item['values'][0]
    else:
        global vesselID
        vesselID=item['values'][0]
        vesselE.delete(0,END)
        vesselE.insert(0,item['values'][1])
    conn.commit()
    conn.close()
    listWindow.destroy()

def dataStyle():
    style.theme_use('alt')
    treestyle = ttk.Style()
    treestyle.theme_use('default')
    treestyle.configure('Treeview', background='#242930', foreground='white', fieldbackground='#242930', borderwidth=0,hover_color='#242930')
    style.configure('Treeview.Heading', background='#242930', foreground='white',fieldbackground='#242930',borderwidth=0,hover_color='#242930')
    treestyle.map('Treeview', background=[('selected', '#506fa3')], foreground=[('selected', 'black')])
    treestyle.map('Treeview.Heading', background=[('selected', '#506fa3')], foreground=[('selected', 'black')])

def correctDate(t):
    
    today = datetime.now()
    if t==1:k=dateOn.get()
    else: k=dateOff.get()
    if k=="":
        return
    k=k.replace(".","")
    if len(k)==10:
        k
    elif len(k)==8:
        k=k[:2]+'/'+k[2:4]+'/'+k[4:8]
    elif len(k)==6:
        k=k[:2]+'/'+k[2:4]+'/20'+k[4:6]
    elif len(k)==4:
        k=k[:2]+'/'+k[2:4]+'/'+str(today.year)
    elif len(k)==2:
        k=k[:2]+'/'+str('%02d' % today.month)+'/'+str(today.year)
    else: k=''
    if  k=='' or int(k[3:5])>12 or monthrange(int(k[6:10]), int(k[3:5]))[1]<int(k[:2]):
        k=''
    if t==1:
        dateOn.delete(0,END)
        dateOn.insert(0,k)
    else:
        dateOff.delete(0,END)
        dateOff.insert(0,k)

def ChooseList(widget,table): 
    widget.unbind('<FocusIn>')
    global listWindow
    listWindow = Toplevel()
    listWindow.focus_force()
    listWindow.grab_set()
    style=ttk.Style(listWindow)
    listWindow.config(bg='#242930')
    listWindow.iconbitmap(os.path.dirname(__file__)+"/Images/ship.ico")
    listWindow.title('AtoZ Invoice')
    listWindow.geometry('%dx%d+%d+%d' % (600, 200, (addWindow.winfo_screenwidth()/2)-300, (addWindow.winfo_screenheight()/2)-100))
    
    if widget==nameE: 
        columns = (' ID',' Number',' Име',' Телефон',' Имейл')
        SearchBar(listWindow,table,1,addPerson,columns)
    elif widget==positionE: 
        columns = (' ID',' Длъжност',' Ранг',' Департамент')
        SearchBar(listWindow,table,1,addPosition,columns)
    else: 
        columns = (' ID',' Кораб')
        SearchBar(listWindow,table,1,addVessel,columns)
    ShowData(listWindow,table,columns,1,onDoubleClick)
    sEntry.bind('<Return>', lambda e: onDoubleClick(table))

def editContract():
    item = trv.item(trv.focus())
    if len(item['values'])==0: return
    addContractData(item)
    varT.set(item['values'][4])
    nameE.delete(0,END)
    nameE.insert(0,item['values'][1])
    nameE.unbind('<FocusIn>')
    positionE.delete(0,END)
    positionE.insert(0,item['values'][2])
    positionE.unbind('<FocusIn>')
    vesselE.delete(0,END)
    vesselE.insert(0,item['values'][6])
    vesselE.unbind('<FocusIn>')
    dateOn.delete(0,END)
    dateOn.insert(0,item['values'][7])
    dateOff.delete(0,END)
    dateOff.insert(0,item['values'][8])
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute("""SELECT vesselChange,crewID,positionID,vesselID,cadet,typeCrew FROM contracts WHERE oid like ?""", (str(item['values'][0]),))
    data = c.fetchall()
    global crewID,positionID,vesselID
    crewID = data[0][1]    
    positionID = data[0][2]    
    vesselID = data[0][3]    
    checkVar.set(int(data[0][0]))
    varC.set(int(data[0][4]))
    varTP.set(data[0][5])
    conn.commit()
    conn.close()

def addContractData(k):
    global addWindow
    addWindow=Toplevel(root,takefocus=True)
    addWindow.grab_set()
    addWindow.focus_force()
    w = 750 # Width 
    h = 400 # Height
    screen_width = addWindow.winfo_screenwidth()  # Width of the screen
    screen_height = addWindow.winfo_screenheight() # Height of the screen
    x = (screen_width/2) - (w/2)
    y = (screen_height/2) - (h/2)
    addWindow.geometry('%dx%d+%d+%d' % (w, h, x, y))
    addWindow.config(bg='#363c45')
    addWindow.iconbitmap(os.path.dirname(__file__)+"/Images/ship.ico")
    image = Image.open(os.path.dirname(__file__)+'/Images/browse.png')
    img = ctk.CTkImage(light_image=image.resize((15, 15)))
    global nameE,positionE,rating,vesselE,department,cadet,typeCrew,dateOn,dateOff,medical,change,typeContract
    
    global varT,varC,varTP
    varTP=ctk.StringVar(value='general')
    varT = ctk.StringVar(value='temporary')
    global checkVar
    checkVar = IntVar(value=0)
    typeCrew = ctk.CTkOptionMenu(addWindow,values=['general', 'riding', 'training'],variable=varTP).place(relwidth = 0.25, relheight = 0.04,relx=0.23,rely=0.13)
    typeCrewL = ctk.CTkLabel(addWindow,text='Тип на работника',anchor=W).place(relwidth = 0.15, relheight = 0.04,relx=0.05,rely=0.13) 
    typeContract = ctk.CTkOptionMenu(addWindow,values=['temporary', 'permament'],variable=varT).place(relwidth = 0.25, relheight = 0.04,relx=0.23,rely=0.2)
    typeL = ctk.CTkLabel(addWindow,text='Tип',bg_color='#363c45',anchor=W).place(relwidth = 0.05, relheight = 0.03,relx=0.05,rely=0.20)
    nameL = ctk.CTkLabel(addWindow,text='Име',bg_color='#363c45',anchor=W).place(relwidth = 0.051, relheight = 0.03,relx=0.05,rely=0.28)
    nameE = ctk.CTkEntry(addWindow)
    nameE.place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.28)
    nameE.bind('<FocusIn>',lambda e:ChooseList(nameE,'personalDetails'))
    addName = ctk.CTkButton(addWindow,image=img,text='',fg_color='#363c45', bg_color='#363c45',hover_color='#363c45',command=lambda:ChooseList(nameE,'personalDetails'),height=10*kH,width=10*kW).place(relwidth = 0.04, relheight = 0.08,relx=0.50,rely=0.276)
    positionL= ctk.CTkLabel(addWindow,text='Позиция',bg_color='#363c45',anchor=W).place(relwidth = 0.1, relheight = 0.03,relx=0.05,rely=0.38)
    positionE = ctk.CTkEntry(addWindow)
    positionE.place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.37)
    positionE.bind('<FocusIn>',lambda e:ChooseList(positionE,'positions'))
    addPosition = ctk.CTkButton(addWindow,image=img,text='',fg_color='#363c45', bg_color='#363c45',hover_color='#363c45',command=lambda:ChooseList(positionE,'positions'),height=10*kH,width=10*kW).place(relwidth = 0.04, relheight = 0.08,relx=0.50,rely=0.366)
    vesselL = ctk.CTkLabel(addWindow,text='Кораб',bg_color='#363c45',anchor=W).place(relwidth = 0.051, relheight = 0.03,relx=0.05,rely=0.47)
    vesselE = ctk.CTkEntry(addWindow)
    vesselE.place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.46)
    vesselE.bind('<FocusIn>',lambda e:ChooseList(vesselE,'vessels'))
    addVessel =ctk.CTkButton(addWindow,text='',image=img,fg_color='#363c45', bg_color='#363c45',hover_color='#363c45',command=lambda:ChooseList(vesselE,'vessels'),height=10*kH,width=10*kW).place(relwidth = 0.04, relheight = 0.08,relx=0.50,rely=0.456)
    varC = IntVar(value=0)
    cadetL = ctk.CTkLabel(addWindow,text='Кадет',bg_color='#363c45',anchor=W).place(relwidth = 0.09, relheight = 0.03,relx=0.05,rely=0.56)
    cadet = ctk.CTkCheckBox(addWindow,text='',bg_color='#363c45',variable=varC).place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.55)
    dateOnL = ctk.CTkLabel(addWindow,text='Дата на качване',bg_color='#363c45',anchor=W).place(relwidth = 0.15, relheight = 0.03,relx=0.05,rely=0.65)
    dateOn = ctk.CTkEntry(addWindow)
    dateOn.place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.64)
    dateOn.bind('<FocusOut>',lambda e:correctDate(1))
    dateOn.bind('<Return>',lambda e:correctDate(0))
    dateOffL = ctk.CTkLabel(addWindow,text='Дата на слизане',bg_color='#363c45',anchor=W).place(relwidth = 0.15, relheight = 0.03,relx=0.05,rely=0.74)
    dateOff = ctk.CTkEntry(addWindow)
    dateOff.place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.73)
    dateOff.bind('<FocusOut>',lambda e:correctDate(0))
    dateOff.bind('<Return>',lambda e:correctDate(0))
    changeL = ctk.CTkLabel(addWindow, text='Смяна на кораб',bg_color='#363c45',anchor=W).place(relwidth = 0.15, relheight = 0.03,relx=0.05,rely=0.83)
    change = ctk.CTkCheckBox(addWindow,text='',bg_color='#363c45',variable=checkVar)
    change.place(relwidth = 0.25, relheight = 0.06,relx=0.23,rely=0.82)
    secretEntry = ctk.CTkEntry(addWindow,state='disable',takefocus=True)
    img = ImageTk.PhotoImage(Image.open(os.path.dirname(__file__)+'/Images/accept.png'))
    if k==0:
        addB = ctk.CTkButton(addWindow,image=img,text='',fg_color='#363c45',hover_color='#363c45',command=lambda:[addContractInfo()])
        addB.place(relwidth = 0.07, relheight = 0.1,relx=0.04,rely=0.88)
    else: 
        addB = ctk.CTkButton(addWindow,image=img,text='',fg_color='#363c45',hover_color='#363c45',command=lambda:[updateContractInfo(k)])
        addB.place(relwidth = 0.07, relheight = 0.1,relx=0.04,rely=0.88)
    img = ImageTk.PhotoImage(Image.open(os.path.dirname(__file__)+'/Images/cancel.png'))
    closeButton = ctk.CTkButton(addWindow,image=img,text='',fg_color='#363c45',hover_color='#363c45',command=addWindow.destroy)
    closeButton.place(relwidth = 0.07, relheight = 0.1,relx=0.1,rely=0.88)
    CreateToolTip(addB, text = 'Добави')
    CreateToolTip(closeButton, text = 'Cancel')
    if k==0: addWindow.bind('<Return>',lambda e:[addContractInfo()])
    else:addWindow.bind('<Return>',lambda e:[updateContractInfo(k)])
    addWindow.bind('<Escape>',lambda e:addWindow.destroy())

def addContractInfo():
    if varT.get()=='' or nameE.get=='' or vesselE.get=='' or positionE.get=='':
        return
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute("""INSERT INTO contracts VALUES( :crewID,:positionID,:type,:vesselID, :on, :off, :change, :medical, :cadet, :typeCrew)""", {'crewID':crewID, 'positionID':positionID,'type':varT.get(), 'vesselID':vesselID,'on': dateOn.get(),'off':dateOff.get(),'change':change.get(),'medical':0,'cadet':varC.get(),'typeCrew':varTP.get()})         
    conn.commit()
    conn.close() 
    ReOpen(-1)

def updateContractInfo(k):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    sql_update_query = """Update contracts set crewID = ?,positionID = ?, contractType = ?, vesselID = ?, dateOn = ?, dateOff = ?, vesselChange = ?, medical = ?, cadet = ?,typeCrew = ? where oid = ?"""
    data = (crewID,positionID,varT.get(),vesselID,dateOn.get(),dateOff.get(),change.get(),0,varC.get(),varTP.get(),k['values'][0])
    c.execute(sql_update_query, data)      
    conn.commit()
    conn.close() 
    ReOpen(k['values'][0]-1)

def ReOpen(i):
    addWindow.destroy()
    dataShow.destroy()
    ShowData(contractAddW,'contracts',(' ID',' Име',' Длъжност',' Ранг',' Тип на договора',' Номер',' Кораб',' Дата на качване',' Дата на слизане'),2,editContract)
    trv.selection_set(trv.get_children()[i]) 
    trv.after(0,lambda: trv.focus_set())
    trv.after(0,lambda: trv.focus(trv.get_children()[i]))

def sortDate(t):
    if(t[0]!=''): return datetime.strptime(t[0], '%d/%m/%Y')
    return datetime.strptime("01/01/3000", '%d/%m/%Y')

def treeview_sort_column(tv, col, reverse, columns):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    if   col == ' Дата на качване' or col ==' Дата на слизане':
        l.sort(reverse=reverse,
                      key=lambda t: sortDate(t))
    elif col == ' ID':
        l.sort(reverse=reverse,
                      key=lambda t: int(t[0]))
    else: 
        l.sort(reverse=reverse)
    for index, (val, k) in enumerate(l):
        tv.move(k, '', index)
    
    for col2 in columns:
        tv.heading(col2, text=col2, command=lambda _col=col2: 
                 treeview_sort_column(tv, _col, False,columns))
    
    tv.heading(col, text=col, command=lambda _col=col: 
                 treeview_sort_column(tv, _col, not reverse,columns))

def update(records):
    for item in trv.get_children():
      trv.delete(item)
    for record in records:
        trv.insert('', 'end',values = record)

def searchData(event,table):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    if event.keysym=="Up" or event.keysym=="Down": return
    if(table=='positions'):
        c.execute("""SELECT oid,* from {} where oid like ? || '%' or position like '%' || ? || '%' or rank like '%' || ? || '%'""".format(table),(sEntry.get(),sEntry.get(),sEntry.get(),))
    elif(table=='personalDetails'):
        c.execute("""SELECT oid,* from personalDetails where oid like ? || '%' or crewName like '%' || ? || '%' or crewID like ? || '%' or phoneNumber like '%' || ? || '%' or mail like '%' || ? || '%'""",(sEntry.get(),sEntry.get(),sEntry.get(),sEntry.get(),sEntry.get(),))
    elif(table=='vessels'):
        c.execute("""SELECT oid,* from vessels where oid like ? || '%' or vessel like '%' || ? || '%'""",(sEntry.get(),sEntry.get()))
    records=c.fetchall()
    update(records)
    if trv.get_children():
        trv.focus(trv.get_children()[0])
        trv.selection_set(trv.get_children()[0])
        sEntry.unbind('<Down>')
        sEntry.unbind('<Up>')
        sEntry.bind('<Down>',lambda e:UpDown(1))
        sEntry.bind('<Up>',lambda e:UpDown(len(trv.get_children())-1))
    conn.commit()
    conn.close()

def UpDown(i):
    if trv.get_children():
        trv.focus(trv.get_children()[i])
        trv.selection_set(trv.get_children()[i])
        trv.see(trv.get_children()[i])
        sEntry.unbind('<Down>')
        sEntry.unbind('<Up>')
        sEntry.bind('<Down>',lambda e:UpDown((i+1)%len(trv.get_children())))
        sEntry.bind('<Up>',lambda e:UpDown((i-1+len(trv.get_children()))%len(trv.get_children())))

def defaultData(table):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    if table=='contracts':
        c.execute("""SELECT contracts.oid,personalDetails.crewName,positions.position,positions.rank,contracts.contractType,personalDetails.crewID,vessels.vessel,contracts.dateOn,contracts.dateOff
                     FROM contracts
                     INNER JOIN personalDetails ON personalDetails.oid=contracts.crewID
                     INNER JOIN positions ON positions.oid = contracts.positionID
                     INNER JOIN vessels ON vessels.oid = contracts.vesselID;""")
    elif table =='prices': c.execute("""SELECT oid,date from {}""".format(table))
    else:
        c.execute("""SELECT oid,* from {}""".format(table))
    # print(c.fetchall())
    records= c.fetchall()
    update(records)
    if len(trv.get_children())!=0:
        trv.selection_set(trv.get_children()[0]) 
        trv.after(0,lambda: trv.focus_set())
        trv.after(0,lambda: trv.focus(trv.get_children()[0]))

def deleteRow(table):
    record = messagebox.askokcancel('','This record will be permamnetly deleted!',icon='warning')
    if record == 0:
        return
    else:
        if table=='contracts':items = trv.selection()
        else:items = trv.selection()
        conn = sqlite3.connect(pathGlobal+'/data.db')
        c = conn.cursor()
        # print(trv.item(item,'values')[0])
        for item in items:
            c.execute("""Delete from {} where oid = ? """.format(table),(trv.item(item,'values')[0],))
        conn.commit()
        conn.close()
        defaultData(table)
        # if(trv)   

def SearchBar(window,table,t,func,columns):
    searchBar = ctk.CTkFrame(window,bg_color='#242930', fg_color='#242930', border_width=0,height=40*kW)
    searchBar.pack(fill='both', expand='no')
    image = Image.open(os.path.dirname(__file__)+'/Images/search.png')
    img = ctk.CTkImage(light_image=image)
    searchIcon = ctk.CTkLabel(searchBar,image=img,text="").place(relx=0.01,rely=0.18)
    global sEntry
    image = Image.open(os.path.dirname(__file__)+'/Images/browse.png')
    img = ctk.CTkImage(light_image=image.resize((15, 15)))
    sEntry = ctk.CTkEntry(searchBar,font=('Arial', 12),fg_color='white',text_color='#242930')
    if t==1:
        sEntry.place(relwidth = 0.15, relheight = 0.8,relx=0.05,rely=0.2)
        aButton = ctk.CTkButton(searchBar,image=img,text='',bg_color='#242930', fg_color='#242930',hover_color='#242930',command=lambda: [func(0),ShowData(window,table,columns,1,onDoubleClick),sEntry.bind('<Return>', lambda e: onDoubleClick(table))],height=10*kH,width=10*kW)
        aButton.place(relx=0.2,rely=0.18)
    else: sEntry.place(relwidth = 0.15, relheight = 0.8,relx=0.03,rely=0.2)
    sEntry.after(0,sEntry.focus_force)
    sEntry.bind('<KeyRelease>',lambda e: searchData(e,table))
    sEntry.bind('<Down>',lambda e: UpDown(0))
    sEntry.bind('<Up>',lambda e: UpDown(0))

def NewWindow(window):
    window.config(bg='#242930')
    window.after(0, lambda:window.state('zoomed'))
    window.geometry('%dx%d+%d+%d' % (w, h, x, y))
    window.after(201,lambda:window.iconbitmap(os.path.dirname(__file__)+"/Images/ship.ico"))
    window.title('AtoZ Invoice')
    root.withdraw()
    global style
    style=ttk.Style(window)
    make_textmenu(window)
    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)

def ShowData(window,table,columns,t,editfunc):
    global dataShow
    dataShow = ctk.CTkFrame(window,bg_color='#242930', fg_color='#242930', border_width=0)
    if table == 'contracts': dataShow.place(relwidth=1,relheight=0.7,relx=0,rely=0.01)
    elif t==1:dataShow.place(relwidth=1,relheight=0.7,relx=0,rely=0.2)
    else: dataShow.place(relwidth=1,relheight=0.7,relx=0,rely=0.04)
    global trv
    
    tree_scroll = ttk.Scrollbar(dataShow, orient='vertical')
    tree_scroll.pack(side=RIGHT, fill=BOTH)
    trv = ttk.Treeview(dataShow, columns=columns,show='headings',yscrollcommand=tree_scroll.set)    
    trv.pack(fill='both',expand='yes',padx=20, pady=10)
    for col in columns:
        trv.heading(col, text=col,anchor=W, command=lambda _col=col: 
                    treeview_sort_column(trv, _col, False,columns))
    tree_scroll.config(command = trv.yview)
    dataStyle()
    defaultData(table)
    trv.column(' ID',width=35,minwidth=35,stretch=NO) 
    for i in range(1,len(columns)):
        trv.column(columns[i],width=100,stretch=YES)
    trv.bind('<Delete>',lambda e: deleteRow(table))
    if t==0:
        trv.bind('<Double-1>',lambda e:[ eButton.place_forget(),editfunc()])
        trv.bind('<Return>',lambda e:[ eButton.place_forget(),editfunc()])
    if t==1:
        trv.bind('<Double-1>',lambda e: editfunc(table))
    if t==2:
        trv.bind('<Double-1>',lambda e: editfunc())
        trv.bind('<Return>',lambda e: editfunc())
    
        
    
    

def editPerson():
    item = trv.item(trv.focus())
    if len(item['values'])==0: return
    crewID.delete(0,END)
    crewID.insert(0,item['values'][1])
    name.delete(0,END)
    name.insert(0,item['values'][2])
    phone.delete(0,END)
    phone.insert(0,item['values'][3])
    mail.delete(0,END)
    mail.insert(0,item['values'][4])
    global uButton
    uButton =ctk.CTkButton(dataAdd, text='Обнови',command=lambda:[updatePersonData(item['values'][0]),uButton.place_forget(),eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.86),crewID.delete(0,END),name.delete(0,END),phone.delete(0,END),mail.delete(0,END)])
    uButton.place(relwidth = 0.037, relheight = 0.135,relx=0.226,rely=0.86)

def updatePersonData(k):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute('UPDATE personalDetails SET crewID=?,crewName=?,phoneNumber=?,mail=? where oid=?',(crewID.get(),name.get(),phone.get(),mail.get(),k))
    conn.commit()
    conn.close()
    defaultData('personalDetails')

def addPerson(flag):
    # Example(root).pack(fill='both', expand=True)
    personAddW = ctk.CTkToplevel(root)
    personAddW.grab_set()
    if flag==1: personAddW.protocol("WM_DELETE_WINDOW",  lambda: [personAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())])
    NewWindow(personAddW) #Window style
    SearchBar(personAddW,'personalDetails',0,0,0) #Search
    columns = (' ID',' Number',' Име',' Телефон',' Имейл') # Show data
    ShowData(personAddW,'personalDetails',columns,0,editPerson)

    global dataAdd,crewID,name,phone,mail,eButton
    dataAdd = ctk.CTkFrame(personAddW,height=250*kH,bg_color='#242930', fg_color='#242930', border_width=0)
    dataAdd.place(relwidth=1,relheight=0.2,relx=0,rely=0.75)
    crewID = ctk.CTkEntry(dataAdd,font=('Arial', 12),fg_color='white',text_color='#242930')
    crewID.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.1)
    crewIDL = ctk.CTkLabel(dataAdd,text='Номер:',text_color='white').place(relwidth = 0.03, relheight = 0.13,relx=0.05,rely=0.1)
    name = ctk.CTkEntry(dataAdd,font=('Arial', 12),fg_color='white',text_color='#242930')
    name.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.3)
    nameL = ctk.CTkLabel(dataAdd,text='Име:',text_color='white').place(relwidth = 0.03, relheight = 0.13,relx=0.05,rely=0.3)
    phone = ctk.CTkEntry(dataAdd,font=('Arial', 12),fg_color='white',text_color='#242930')
    phone.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.5)
    phoneL =ctk.CTkLabel(dataAdd,text='Телефон:',text_color='white').place(relwidth = 0.03, relheight = 0.13,relx=0.05,rely=0.5)
    mail = ctk.CTkEntry(dataAdd,font=('Arial', 12),fg_color='white',text_color='#242930')
    mail.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.7)
    mailL =ctk.CTkLabel(dataAdd,text='Имейл:',text_color='white').place(relwidth = 0.03, relheight = 0.13,relx=0.05,rely=0.7)
    Ivan = ctk.CTkOptionMenu(dataAdd)
    addButton =ctk.CTkButton(dataAdd, text='Добави',command=lambda: [addPersonData(),crewID.delete(0,END),name.delete(0,END),phone.delete(0,END),mail.delete(0,END),ShowData(personAddW,'personalDetails',columns,0,editPerson)],fg_color="green",hover_color="darkgreen").place(relwidth = 0.037, relheight = 0.135,relx=0.1,rely=0.86)
    if flag==1:
        dButton =ctk.CTkButton(dataAdd,text='Изтрий', command=lambda: deleteRow('personalDetails'),fg_color="red",hover_color="darkred").place(relwidth = 0.037, relheight = 0.135,relx=0.142,rely=0.86)
        cButton = ctk.CTkButton(dataAdd,text='Изчисти', command=lambda: [crewID.delete(0,END),name.delete(0,END),phone.delete(0,END),mail.delete(0,END),eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.86),uButton.place_forget()]).place(relwidth = 0.037, relheight = 0.135,relx=0.184,rely=0.86)
        eButton = ctk.CTkButton(dataAdd, text='Редактирай',command=lambda: [ eButton.place_forget(),editPerson()])
        eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.86)
    if flag==1: windowBack =ctk.CTkButton(personAddW,text='Back', command=lambda: [personAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    else: windowBack =ctk.CTkButton(personAddW,text='Back', command=lambda: [personAddW.destroy()]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    

def addContract():
    global contractAddW
    contractAddW= ctk.CTkToplevel(root) 
    contractAddW.grab_set()
    contractAddW.protocol("WM_DELETE_WINDOW",  root.destroy)
    NewWindow(contractAddW) 
    columns = (' ID',' Име',' Длъжност',' Ранг',' Тип на договора',' Номер',' Кораб',' Дата на качване',' Дата на слизане')
    ShowData(contractAddW,'contracts',columns,2,editContract)
    trv.column(' ID',width=35,minwidth=35,stretch=NO) 
    for i in range(1,len(columns)):
        trv.column(columns[i],width=100,stretch=YES) 
    Ivan = ctk.CTkOptionMenu(contractAddW)

    addButton = ctk.CTkButton(contractAddW, text='Добави',command=lambda:[addContractData(0)],fg_color="green",hover_color="darkgreen").place(relwidth = 0.06, relheight = 0.04,relx=0.02,rely=0.9)
    delete =ctk.CTkButton(contractAddW, text='Изтрий',command=lambda: deleteRow('contracts'),fg_color="red",hover_color="darkred").place(relwidth = 0.06, relheight = 0.04,relx=0.1,rely=0.9)
    windowBack =ctk.CTkButton(contractAddW,text='Back', command=lambda: [contractAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    eButton = ctk.CTkButton(contractAddW, text='Редактирай',command=editContract)
    eButton.place(relwidth = 0.06, relheight = 0.04,relx=0.18,rely=0.9)

def editVessel():
    item = trv.item(trv.focus())
    if len(item['values'])==0: 
        eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.63)
        return
    name.delete(0,END)
    name.insert(0,item['values'][1])
    company.delete(0,END)
    company.insert(0,item['values'][2])
    global uButton
    uButton = ctk.CTkButton(dataAdd, text='Обнови',command=lambda:[updateVesselData(item['values'][0]),uButton.place_forget(),eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.63),name.delete(0,END),company.delete(0,END)],width=50*kW)
    uButton.place(relwidth = 0.037, relheight = 0.135,relx=0.226,rely=0.63)

def updateVesselData(k):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute('UPDATE vessels SET vessel=?,company=? where oid=?',(name.get(),company.get(),k))
    conn.commit()
    conn.close()
    defaultData('vessels')

def addVessel(flag): 
    global vesselAddW
    vesselAddW = ctk.CTkToplevel(root)
    vesselAddW.grab_set()
    if flag==1: vesselAddW.protocol("WM_DELETE_WINDOW",  lambda: [vesselAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())])
    NewWindow(vesselAddW) # Window style
    SearchBar(vesselAddW,'vessels',0,0,0) # Search
    columns = (' ID',' Кораб','Компания') # Show data
    ShowData(vesselAddW,'vessels',columns,0,editVessel)
    # Add data
    global dataAdd,name,company,eButton
    dataAdd = ctk.CTkFrame(vesselAddW,height=250*kH,bg_color='#242930', fg_color='#242930', border_width=0)
    dataAdd.place(relwidth=1,relheight=0.2,relx=0,rely=0.76)

    name = ctk.CTkEntry(dataAdd,font=('Arial', 12),fg_color='white',text_color='#242930')
    name.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.2)
    nameL = ctk.CTkLabel(dataAdd,text='Кораб:',text_color='white').place(relwidth = 0.05, relheight = 0.1,relx=0.02,rely=0.22)
    company = ctk.CTkEntry(dataAdd,font=('Arial', 12),fg_color='white',text_color='#242930')
    company.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.4)
    companyL = ctk.CTkLabel(dataAdd,text='Компания:',text_color='white').place(relwidth = 0.05, relheight = 0.1,relx=0.02,rely=0.42)
    Ivan = ctk.CTkOptionMenu(dataAdd)

    if flag==1:
        addButton =ctk.CTkButton(dataAdd, text='Добави',command=lambda: [addVesselData(),name.delete(0,END),company.delete(0,END),ShowData(vesselAddW,'vessels',columns,0,editVessel)],fg_color="green",hover_color="darkgreen").place(relwidth = 0.037, relheight = 0.135,relx=0.1,rely=0.63)
        dButton =ctk.CTkButton(dataAdd,text='Изтрий', command=lambda: deleteRow('vessels'),fg_color="red",hover_color="darkred").place(relwidth = 0.037, relheight = 0.135,relx=0.142,rely=0.63)
        cButton = ctk.CTkButton(dataAdd,text='Изчисти', command=lambda: [name.delete(0,END),company.delete(0,END),eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.63),uButton.place_forget()]).place(relwidth = 0.037, relheight = 0.135,relx=0.184,rely=0.63)
        eButton = ctk.CTkButton(dataAdd, text='Редактирай',command=lambda: [ eButton.place_forget(),editVessel()])
        eButton.place(relwidth = 0.045, relheight = 0.135,relx=0.226,rely=0.63)
    if flag==1:windowBack =ctk.CTkButton(vesselAddW,text='Back', command=lambda: [vesselAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    else: windowBack =ctk.CTkButton(vesselAddW,text='Back', command=lambda: [vesselAddW.destroy()]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    

def editPosition():
    item = trv.item(trv.focus())
    if len(item['values'])==0: 
        eButton.place(relwidth = 0.045, relheight = 0.13,relx=0.226,rely=0.86)
        return
    position.delete(0,END)
    position.insert(0,item['values'][1])
    varR.set(item['values'][2])
    varTC.set(item['values'][3])
    global uButton
    uButton =ctk.CTkButton(dataAdd, text='Обнови',command=lambda:[updatePositionData(item['values'][0]),position.delete(0,END),varR.set(""),varTC.set(""),eButton.place(relwidth = 0.045, relheight = 0.13,relx=0.226,rely=0.86),uButton.place_forget()])
    uButton.place(relwidth = 0.0357, relheight = 0.13,relx=0.226,rely=0.86)

def updatePositionData(k):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute('UPDATE positions SET position=?,rank=?,type=? where oid=?',(position.get(),rank.get(),varTC.get(),k))
    conn.commit()
    conn.close()
    defaultData('positions')

def addPosition(flag):
    # Example(root).pack(fill='both', expand=True)
    global positionAddW
    positionAddW = ctk.CTkToplevel(root)
    positionAddW.grab_set()
    if flag==1:positionAddW.protocol("WM_DELETE_WINDOW", lambda: [positionAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())])
    NewWindow(positionAddW) #Window style
    SearchBar(positionAddW,'positions',0,0,0) #Search
    columns = (' ID',' Длъжност',' Ранг',' Департамент') # Show data
    ShowData(positionAddW,'positions',columns,0,editPosition)
    # Add data
    global dataAdd,position,rank,varTC,varR,eButton

    dataAdd = ctk.CTkFrame(positionAddW,height=250*kH,bg_color='#242930', fg_color='#242930', border_width=0)
    dataAdd.place(relwidth=1,relheight=0.2,relx=0,rely=0.76)

    position = ctk.CTkEntry(dataAdd,width=400*kW,font=('Arial', 12),fg_color='white',text_color='#242930')
    position.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.17)
    positionL = ctk.CTkLabel(dataAdd,text='Длъжност:',text_color='white',anchor=W).place(relwidth = 0.05, relheight = 0.1,relx=0.02,rely=0.19)

    varR = ctk.StringVar(value='')
    rank = ctk.CTkOptionMenu(dataAdd,values=['officer','rating'],variable=varR)
    rank.place(relwidth = 0.17, relheight = 0.13,relx=0.1,rely=0.40)
    rankL =ctk.CTkLabel(dataAdd,text='Ранг:',text_color='white',anchor=W).place(relwidth = 0.05, relheight = 0.1,relx=0.02,rely=0.42)

    varTC = ctk.StringVar(value='')
    typeCrewL = ctk.CTkLabel(dataAdd,text='Департамент:',anchor=W).place(relwidth = 0.08, relheight = 0.1,relx=0.02,rely=0.65)
    typeCrew = ctk.CTkOptionMenu(dataAdd,values=['Deck and Engine','Hotel'],variable=varTC).place(relwidth = 0.168, relheight = 0.13,relx=0.1,rely=0.63)

    addButton =ctk.CTkButton(dataAdd, text='Добави',command=lambda: [addPositionData(),position.delete(0,END),varR.set(""),varTC.set(""), ShowData(positionAddW,'positions',columns,0,editPosition)],fg_color="green",hover_color="darkgreen").place(relwidth = 0.037, relheight = 0.135,relx=0.1,rely=0.86)
    if flag==1:
        dButton  =ctk.CTkButton(dataAdd,text='Изтрий', command=lambda: deleteRow('positions'),fg_color="red",hover_color="darkred").place(relwidth = 0.037, relheight = 0.135,relx=0.142,rely=0.86)
        cButton = ctk.CTkButton(dataAdd,text='Изчисти', command=lambda: [position.delete(0,END),varR.set(""),varTC.set(""),eButton.place(relwidth = 0.045, relheight = 0.13,relx=0.226,rely=0.86),uButton.place_forget()]).place(relwidth = 0.037, relheight = 0.135,relx=0.184,rely=0.86)
        eButton = ctk.CTkButton(dataAdd, text='Редактирай',command=lambda: [ eButton.place_forget(),editPosition()])
        eButton.place(relwidth = 0.045, relheight = 0.13,relx=0.226,rely=0.86)
    if flag==1:windowBack =ctk.CTkButton(positionAddW,text='Back', command=lambda: [positionAddW.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    else: windowBack =ctk.CTkButton(positionAddW,text='Back', command=lambda: [positionAddW.destroy()]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)
    

def browse_button():
    config = configparser.ConfigParser()
    print(os.path.dirname(__file__)+'/config.txt')
    config.read(os.path.dirname(__file__)+'/config.txt')
    filename = filedialog.askdirectory()
    config['config']['path']=filename
    with open(os.path.dirname(__file__)+'/config.txt', 'w') as configfile:
        config.write(configfile)

def createApendix():
    global apendixWindow
    apendixWindow =Toplevel(root,takefocus=True)
    apendixWindow.grab_set()
    NewSmallWindow(apendixWindow)
    global varComp,varM,year,varTP
    varTP = ctk.StringVar(value='general')
    varComp = StringVar(value='')
    first = today.replace(day=1)
    last_month = first - timedelta(days=1)
    varM = StringVar(value=last_month.strftime('%B'))
    typeCrew = ctk.CTkOptionMenu(apendixWindow,values=['general','riding','training'],variable=varTP).place(relwidth = 0.4, relheight = 0.06,relx=0.4,rely=0.1)
    typeCrewL = ctk.CTkLabel(apendixWindow,text='Тип на персонала',anchor=W).place(relwidth = 0.2, relheight = 0.06,relx=0.1,rely=0.1)
    companyL = ctk.CTkLabel(apendixWindow,text='Компания',anchor=W).place(relwidth = 0.15, relheight = 0.06,relx=0.1,rely=0.3)
    companyChoose = ctk.CTkOptionMenu(apendixWindow, variable=varComp,values=['AIDA','Costa','CCSL']).place(relwidth = 0.4, relheight = 0.06,relx=0.4,rely=0.3)
    monthL = ctk.CTkLabel(apendixWindow,text='Месец',anchor=W).place(relwidth = 0.15, relheight = 0.06,relx=0.1,rely=0.5)
    monthChoose = ctk.CTkOptionMenu(apendixWindow,variable = varM, values=['January','February','March','April','May','June','July','August','September','October','November','December']).place(relwidth = 0.4, relheight = 0.06,relx=0.4,rely=0.5)
    windowBack =ctk.CTkButton(apendixWindow,text='Back', command=lambda: [apendixWindow.destroy(),root.after(0, lambda:root.state('zoomed')),root.after(1, lambda:root.deiconify())]).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)

    yearL = ctk.CTkLabel(apendixWindow,text='Година',anchor=W).place(relwidth = 0.15, relheight = 0.06,relx=0.1,rely=0.7)
    year = ctk.CTkEntry(apendixWindow)
    year.insert(0,str(today.year))
    year.place(relwidth = 0.1, relheight = 0.06,relx=0.4,rely=0.7)
    generate = ctk.CTkButton(apendixWindow,text='Генериране на справка',command=generateApendix).place(relwidth = 0.4, relheight = 0.06,relx=0.1,rely=0.8)

def dateCheck(dOn, dOff):
    global date_format,date_format2
    date_format = '%d/%m/%Y'
    date_format2 = '%d.%m.%Y'
    date_obj = datetime.strptime(dOn, date_format)
    if dOff!='':
        date_obj2 = datetime.strptime(dOff, date_format)
    date_obj3 = datetime.strptime(varM.get(), '%B')
    date_obj4 = datetime.strptime(varM.get(), '%B')
    t=0
    t2=0
    if date_obj.year < int(year.get()) or (date_obj.year ==  int(year.get()) and date_obj.month<=date_obj3.month):t=1
    if dOff=='' or date_obj2.year >  int(year.get()) or (date_obj2.year ==  int(year.get()) and date_obj2.month>=date_obj4.month):t2=1
    return t==1 and t2==1

def generateApendix():
    pth = os.path.dirname(__file__)
    fname = os.path.join(pth, './template.xlsx')
    writer = BookWriter(fname)
    writer.jinja_env.globals.update(dir=dir, getattr=getattr)
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    conn.create_function('dateCheck', 2, dateCheck)
    # print(dateCheck('22/06/2023','22/06/2023'))
    c.execute("""SELECT personalDetails.crewName,positions.position,positions.rank,contracts.contractType,personalDetails.crewID,vessels.vessel,contracts.dateOn,contracts.dateOff,contracts.cadet,positions.type,contracts.vesselChange,contracts.typeCrew,vessels.company
                FROM contracts
                INNER JOIN personalDetails ON personalDetails.oid=contracts.crewID
                INNER JOIN positions ON positions.oid = contracts.positionID
                INNER JOIN vessels ON vessels.oid = contracts.vesselID
                WHERE dateCheck(contracts.dateOn,contracts.dateOff)=1
                AND vessels.company = ?
                AND contracts.typeCrew = ?""",(varComp.get(),varTP.get(),))
    data = [dict(zip([column[0] for column in c.description], row)) for row in c.fetchall()]
    print(data)
    c.execute("""SELECT * from prices""")
    currentDate = datetime(int(year.get()),monthK[varM.get()],1)
    priceData = c.fetchall()
    p=0
    while (p<len(priceData) and datetime.strptime(priceData[p][8], date_format)<=currentDate) : p+=1
    p-=1
    for i in range(len(data)):
        object = datetime.strptime(data[i]['dateOn'], date_format)
        if(data[i]['dateOff']==''):
            data[i]['dateOff']=str(monthrange(int(year.get()), monthK[varM.get()])[1])+"/"+str('%02d' % object.month)+"/"+year.get()
        object2=datetime.strptime(data[i]['dateOff'], date_format)
        if object2.month != monthK[varM.get()] or object2.year!=int(year.get()):
            data[i]['dateOff']=str(monthrange(int(year.get()), monthK[varM.get()])[1])+"/"+str('%02d' % object.month)+"/"+year.get()
        elif object.month!=monthK[varM.get()] or object.year!=int(year.get()):
            data[i]['dateOn']="01/"+str('%02d' % object2.month)+"/"+year.get()
        object = datetime.strptime(data[i]['dateOn'], date_format)
        object2=datetime.strptime(data[i]['dateOff'], date_format)
        data[i]['daysOn']=(date(object2.year, object2.month, object2.day)- date(object.year, object.month, object.day)).days+1
        if varComp.get()=='AIDA':
            if data[i]['rank']=='Officer':
                if data[i]['type'] == 'Deck and Engine':
                    data[i]['deploymentFee']=priceData[p][0]
                    data[i]['manningFee']=data[i]['daysOn']/monthrange(int(year.get()), monthK[varM.get()])[1]*priceData[p][2]
                else:
                    data[i]['deploymentFee']=priceData[p][4]
                    data[i]['manningFee']=data[i]['daysOn']/monthrange(int(year.get()), monthK[varM.get()])[1]*priceData[p][6]
            else:
                if data[i]['type'] == 'Deck and Engine':
                    data[i]['deploymentFee']=priceData[p][1]
                    data[i]['manningFee']=data[i]['daysOn']/monthrange(int(year.get()), monthK[varM.get()])[1]*priceData[p][3]
                else:
                    data[i]['deploymentFee']=priceData[p][5]
                    data[i]['manningFee']=data[i]['daysOn']/monthrange(int(year.get()), monthK[varM.get()])[1]*priceData[p][7]
        else:
            if data[i]['rank']=='Officer':
                if data[i]['type'] == 'Deck and Engine':
                    data[i]['deploymentFee']=priceData[p][0]
                    data[i]['manningFee']=round(priceData[p][2]/30,2)*data[i]['daysOn']   
                else:
                    data[i]['deploymentFee']=priceData[p][4]
                    data[i]['manningFee']=round(priceData[p][6]/30,2)*data[i]['daysOn']
            else:
                if data[i]['type'] == 'Deck and Engine':
                    data[i]['deploymentFee']=priceData[p][1]
                    data[i]['manningFee']=round(priceData[p][3]/30,2)*data[i]['daysOn']  
                else:
                    data[i]['deploymentFee']=priceData[p][5]
                    data[i]['manningFee']=round(priceData[p][7]/30,2)*data[i]['daysOn']  
        if data[i]['cadet']==1:
            data[i]['deploymentFee']=priceData[p][1]
            data[i]['manningFee']=0
        if data[i]['vesselChange']==1 or object.month != monthK[varM.get()] or object.year!=int(year.get()):
            data[i]['deploymentFee']=''
    data.sort(key=lambda t:t['crewName'].split(' ')[1])
    for i in range(len(data)):data[i]['number']=i+1
    person_info2 = {}    
    person_info2['contracts'] = data
    person_info2['monthEn'] = varM.get()
    person_info2['monthBg'] = monthN[varM.get()]
    person_info2['endDate'] = date(int(year.get()),monthK[varM.get()],monthrange(int(year.get()), monthK[varM.get()])[1]).strftime(date_format2)
    person_info2['number'] = varComp.get()+varTP.get()[0]+str(int(monthK[varM.get()]/10))+str(monthK[varM.get()]%10)+year.get()
    print(str(monthK[varM.get()]/10))
    person_info2['year'] = year.get()
    print(data)
    payloads = [person_info2]
    writer.render_book(payloads=payloads)
    fileName = varComp.get()+"_"+varTP.get()[0]+"_"+str(int(monthK[varM.get()]/10))+str(monthK[varM.get()]%10)+"."+year.get()
    print(fileName)
    config = configparser.ConfigParser()
    config.read(os.path.dirname(__file__)+'/config.txt')
    print(config.sections())
    path = config.get('config', 'path')
    
    fname = os.path.join(path, fileName+'.xlsx')
    print(fname)
    writer.save(fname)
    excel = client.Dispatch('Excel.Application') 
    sheets = excel.Workbooks.Open(path+'/'+fileName+'.xlsx') 
    work_sheets = sheets.Worksheets[0] 
    print(path)
    if  os.path.exists(path+'/'+fileName+'.pdf'):os.remove(path+'/' + fileName+'.pdf')
    work_sheets.ExportAsFixedFormat(0,path+'/'+fileName+'.pdf') 
    path+='/'+fileName+'.xlsx'
    sheets.Close(True,path)

def NewSmallWindow(window):
    window.focus_force()
    window.geometry('%dx%d+%d+%d' % (600, 400, (root.winfo_screenwidth()/2)-300, (root.winfo_screenheight()/2)-200))
    window.config(background='#363c45')
    window.iconbitmap(os.path.dirname(__file__)+"/Images/ship.ico")
    window.title('AtoZ Invoice')
    global style
    style=ttk.Style(window)    
    make_textmenu(window)
    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)

def LoadPrices():
    
    item=trv.item(trv.focus())
    addPrice(item['values'][0])
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute("""SELECT oid,* FROM prices WHERE oid={}""".format(str(item['values'][0])))
    data=c.fetchall()
    print(data)
    deOfficerDeploy.insert(0,data[0][1])
    deRatingDeploy.insert(0,data[0][2])
    deOfficerMann.insert(0,data[0][3])
    deRatingMann.insert(0,data[0][4])
    hOfficerDeploy.insert(0,data[0][5])
    hRatingDeploy.insert(0,data[0][6])
    hOfficerMann.insert(0,data[0][7])
    hRatingMann.insert(0,data[0][8])
    dateP.insert(0,data[0][9])
    conn.commit()
    conn.close()

def updatePrice(k):
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute('UPDATE prices SET deOfficerDeploy=?,deRatingDeploy=?,deOfficerMann=?,deRatingMann=?,hOfficerDeploy=?,hRatingDeploy=?,hOfficerMann=?,hRatingMann=?,date=? where oid=?',(deOfficerDeploy.get(),deRatingDeploy.get(),deOfficerMann.get(),deRatingMann.get(),hOfficerDeploy.get(),hRatingDeploy.get(),hOfficerMann.get(),hRatingMann.get(),dateP.get(),k))
    conn.commit()
    conn.close()
    defaultData('prices')
    addPriceDataW.destroy()

def addPriceData():
    conn = sqlite3.connect(pathGlobal+'/data.db')
    c = conn.cursor()
    c.execute("""INSERT INTO prices VALUES( :deOfficerDeploy,:deRatingDeploy,:deOfficerMann,:deRatingMann,:hOfficerDeploy,:hRatingDeploy,:hOfficerMann,:hRatingMann,:date)""", {'deOfficerDeploy':deOfficerDeploy.get(), 'deRatingDeploy':deRatingDeploy.get(),'deOfficerMann':deOfficerMann.get(),'deRatingMann':deRatingMann.get(),'hOfficerDeploy':hOfficerDeploy.get(), 'hRatingDeploy':hRatingDeploy.get(),'hOfficerMann':deOfficerMann.get(),'hRatingMann':deRatingMann.get(),'date':dateP.get()})
    conn.commit()
    conn.close()
    defaultData('prices')
    addPriceDataW.destroy()

def addPrice(k):
    global addPriceDataW
    addPriceDataW = Toplevel()
    addPriceDataW.grab_set()
    NewSmallWindow(addPriceDataW)
    make_textmenu(addPriceDataW)
    addPriceDataW.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    addPriceDataW.bind_class("Entry", "<Control-a>", callback_select_all)

    deL = ctk.CTkLabel(addPriceDataW,text='Deck and Engine',anchor=W).place(relx=0.43,rely=0.01)
    deOfficerL = ctk.CTkLabel(addPriceDataW,text='Officer',anchor=CENTER).place(relx=0.1,rely=0.2)
    deRatingL = ctk.CTkLabel(addPriceDataW,text='Rating',anchor=CENTER).place(relx=0.1,rely=0.3)
    deDeplyL = ctk.CTkLabel(addPriceDataW,text='Deployment',anchor=CENTER).place(relx=0.3,rely=0.1)
    deMannL = ctk.CTkLabel(addPriceDataW,text='Manning',anchor=CENTER).place(relx=0.61,rely=0.1)
    hL = ctk.CTkLabel(addPriceDataW,text='          Hotel',anchor=W).place(relx=0.43,rely=0.41)
    hOfficerL = ctk.CTkLabel(addPriceDataW,text='Officer',anchor=CENTER).place(relx=0.1,rely=0.6)
    hRatingL = ctk.CTkLabel(addPriceDataW,text='Rating',anchor=CENTER).place(relx=0.1,rely=0.7)
    hDeplyL = ctk.CTkLabel(addPriceDataW,text='Deployment',anchor=CENTER).place(relx=0.3,rely=0.5)
    hMannL = ctk.CTkLabel(addPriceDataW,text='Manning',anchor=CENTER).place(relx=0.61,rely=0.5)

    global deOfficerDeploy,deOfficerMann,deRatingDeploy,deRatingMann,hOfficerDeploy,hOfficerMann,hRatingDeploy,hRatingMann,dateP
    deOfficerDeploy = ctk.CTkEntry(addPriceDataW)
    deOfficerDeploy.place(relx=0.3,rely=0.21,relwidth=0.12,relheight=0.05)
    deRatingDeploy = ctk.CTkEntry(addPriceDataW)
    deRatingDeploy.place(relx=0.3,rely=0.31,relwidth=0.12,relheight=0.05)
    deOfficerMann = ctk.CTkEntry(addPriceDataW)
    deOfficerMann.place(relx=0.61,rely=0.21,relwidth=0.12,relheight=0.05)
    deRatingMann = ctk.CTkEntry(addPriceDataW)
    deRatingMann.place(relx=0.61,rely=0.31,relwidth=0.12,relheight=0.05)
    hOfficerDeploy = ctk.CTkEntry(addPriceDataW)
    hOfficerDeploy.place(relx=0.3,rely=0.61,relwidth=0.12,relheight=0.05)
    hRatingDeploy = ctk.CTkEntry(addPriceDataW)
    hRatingDeploy.place(relx=0.3,rely=0.71,relwidth=0.12,relheight=0.05)
    hOfficerMann = ctk.CTkEntry(addPriceDataW)
    hOfficerMann.place(relx=0.61,rely=0.61,relwidth=0.12,relheight=0.05)
    hRatingMann = ctk.CTkEntry(addPriceDataW)
    hRatingMann.place(relx=0.61,rely=0.71,relwidth=0.12,relheight=0.05)
    dateL = ctk.CTkLabel(addPriceDataW,text='Дата:').place(relx=0.05,rely=0.9)
    dateP = ctk.CTkEntry(addPriceDataW)
    dateP.place(relx=0.13,rely=0.9)


    if k==0:addButton =ctk.CTkButton(addPriceDataW, text='Добави',command=addPriceData,fg_color="green",hover_color="darkgreen").place(relx=0.5,rely=0.9)
    else: addButton = ctk.CTkButton(addPriceDataW, text='Добави',command=lambda: updatePrice(k),fg_color="green",hover_color="darkgreen").place(relx=0.5,rely=0.9)


def Price():
    addPriceW = Toplevel()
    addPriceW.grab_set()
    NewSmallWindow(addPriceW)
    columns = (' ID',' Валидно от')
    ShowData(addPriceW,'prices',columns,3,0)
    # dataShow.an="#363c45"
    dataShow.configure( bg_color='#363c45',fg_color='#363c45')
    treestyle = ttk.Style()
    treestyle.theme_use('default')
    treestyle.configure('Treeview', background='#363c45', foreground='white', fieldbackground='#363c45', borderwidth=0,hover_color='#363c45')
    style.configure('Treeview.Heading', background='#363c45', foreground='white',fieldbackground='#363c45',borderwidth=0,hover_color='#363c45')
    
    addButton =ctk.CTkButton(addPriceW, text='Добави',command=lambda: addPrice(0),fg_color="green",hover_color="darkgreen").place(relwidth = 0.2, relheight = 0.08,relx=0.1,rely=0.86)
    dButton  =ctk.CTkButton(addPriceW,text='Изтрий', command=lambda: deleteRow('prices'),fg_color="red",hover_color="darkred").place(relwidth = 0.2, relheight = 0.08,relx=0.4,rely=0.86)
    trv.bind('<Double-1>',lambda e:LoadPrices())

def setting():
    settingsWindow =Toplevel(root,takefocus=True)
    settingsWindow.grab_set()
    NewSmallWindow(settingsWindow)
    browse = ctk.CTkButton(settingsWindow,text='Път за експортиране', command=browse_button).place(relwidth = 0.4, relheight = 0.06,relx=0.1,rely=0.2)
    priceCorection = ctk.CTkButton(settingsWindow,text='Добавяне на ценоразпис',command=Price).place(relwidth = 0.4, relheight = 0.06,relx=0.1,rely=0.4)

def make_textmenu(root):
	global the_menu
	the_menu = Menu(root, tearoff=0)
	the_menu.add_command(label="Cut")
	the_menu.add_command(label="Copy")
	the_menu.add_command(label="Paste")
	the_menu.add_separator()
	the_menu.add_command(label="Select all")   

def show_textmenu(event):
	e_widget = event.widget
	the_menu.entryconfigure("Cut",command=lambda: e_widget.event_generate("<<Cut>>"))
	the_menu.entryconfigure("Copy",command=lambda: e_widget.event_generate("<<Copy>>"))
	the_menu.entryconfigure("Paste",command=lambda: e_widget.event_generate("<<Paste>>"))
	the_menu.entryconfigure("Select all",command=lambda: e_widget.select_range(0, 'end'))
	the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)

def callback_select_all(event):
	# select text after 50ms
	root.after(50, lambda:event.widget.select_range(0, 'end'))

root = ctk.CTk()
root.iconbitmap(os.path.dirname(__file__)+"./Images/ship.ico")
root.title('AtoZ Invoice')
config = configparser.ConfigParser()
#config.read(r'config.txt')
#filename = config.get('config', 'path')
pathGlobal = os.path.dirname(__file__)
root.after(0, lambda:root.state('zoomed'))
root.config(background='#242930')
# root.attributes('-fullscreen', True)
w = round(GetSystemMetrics(0)*0.75)
h = round(GetSystemMetrics(1)*0.75)
screen_width = root.winfo_screenwidth()  # Width of the screen
screen_height = root.winfo_screenheight() # Height of the screen
x = (screen_width/2) - (w/2)
y = (screen_height/2) - (h/2)
root.geometry('%dx%d+%d+%d' % (w, h, x, y))
p=root.winfo_screenheight()
kW = w/1920
kH = h/1080
print(os.path.dirname(__file__)+'/config.txt')
monthK = {'January': 1, 'February': 2, 'March': 3, 'April': 4, 'May': 5, 'June': 6, 'July': 7, 'August': 8, 'September': 9, 'October': 10, 'November': 11, 'December': 12}
monthN = {'January': 'Януари', 'February': 'Февруари', 'March': 'Март', 'April': 'Април', 'May': 'Май', 'June': 'Юни', 'July': 'Юли', 'August': 'Август', 'September': 'Септември', 'October': 'Октомври', 'November': 'Ноември', 'December': 'Декември'}
today = datetime.now()
personAdd =ctk.CTkButton(root,text='Добавяне на човек',font=('Arial',17),border_width=2,border_color='#242930',command=lambda: addPerson(1))
personAdd.place(relwidth = 0.25, relheight = 0.06,relx=0.07,rely=0.22)
contractAdd =ctk.CTkButton(root,text='Добавяне на договор',font=('Arial',17),border_width=2,border_color='#242930',command=addContract).place(relwidth = 0.25, relheight = 0.06,relx=0.07,rely=0.29)
vesselAdd =ctk.CTkButton(root,text='Добавяне на кораб',font=('Arial',17),border_width=2,border_color='#242930',command=lambda: addVessel(1)).place(relwidth = 0.25, relheight = 0.06,relx=0.07,rely=0.36)
positionAdd = ctk.CTkButton(root,text='Добавяне на позиция',font=('Arial',17),border_width=2,border_color='#242930',command=lambda: addPosition(1)).place(relwidth = 0.25, relheight = 0.06,relx=0.07,rely=0.43)
apendixMake = ctk.CTkButton(root,text='Справки',font=('Arial',17),border_width=2,border_color='#242930',command=createApendix).place(relwidth = 0.25, relheight = 0.06,relx=0.07,rely=0.50)
image = PhotoImage(file=os.path.dirname(__file__)+'/Images/settings.png')
image=image.subsample(8,8)
settings = ctk.CTkButton(root,image=image,text='',fg_color='#242930',hover_color='#242930',command=setting).place(relx=0,rely=0.9)
windowExit = ctk.CTkButton(root,text='Exit', command=root.quit).place(relwidth = 0.15, relheight = 0.06,relx=0.40,rely=0.9)


# conn = sqlite3.connect(pathGlobal+'/data.db')
# for i inprint(kW,kH)
# c = conn.cursor()
# # c.execute("""Create Table Ivan2(ime int)""")
# df  = pd.read_csv('test.txt')
# df.columns = df.columns.str.strip()
# c.execute('DELETE FROM personalDetails')
# # c.execute("""CREATE TABLE personalDetails (
# # crewID int,
# # crewName text,
# # phoneNumber int,
# # mail text
# # );""")
# df.to_sql('personalDetails',conn, if_exists='append',index=False)
# c.execute("""SELECT * FROM positions""")

# conn.commit()
# conn.close()




# CREATE TABLE POSITIONS(
#     POSITIONID INT,
#     POSITION text,
#     RANK text
# );
        
# CREATE TABLE VESSELS(
#     VESSELID INT,
#     VESSEL text
# );
root.mainloop()
