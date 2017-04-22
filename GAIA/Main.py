import tkinter as tk
import sys
import tkinter.messagebox as mb
#import os
import pymssql
#from werkzeug._compat import iterlistvalues
from argcomplete.compat import str
from twisted.conch.test.test_helper import HEIGHT
from openpyxl.styles.borders import Side
import codecs

class DBSource:
    def __init__(self,servername,dbusername,dbpsw,dbname):
        self.servername = servername
        self.dbusername = dbusername
        self.dbpsw = dbpsw
        self.dbname = dbname
ehr = DBSource("localhost","sa","1qaz2wsx","ZYBXSCeHR_DB_170207")
        
def connWithDB (DBSource,orderNeed):    
    conn=pymssql.connect(DBSource.servername,DBSource.dbusername,DBSource.dbpsw,database=DBSource.dbname)
    DBlist = []
    cursor=conn.cursor()    
    cursor.execute(orderNeed)
    row=cursor.fetchone()
    while row:
        #print("readline:%s"%(row[0]))
        DBlist.append(row[0])
        row=cursor.fetchone()
    conn.close()
    return DBlist
'''    
ehr = DBSource("localhost","sa","1qaz2wsx","ZYBXSCeHR_DB_170207")
orderNeed = """select top 10 truename from psnaccount"""
selectDataFromPsnaccount = connWithDB(ehr,orderNeed)
'''
#调整窗口位置居中
'''
def get_screen_size(window):
    return window.winfo_screenwidth(), window.winfo_screenheight()

def get_window_size(window):
    return window.winfo_reqwidth(), window.winfo_reqheight()
'''
def center_window(root, width, height):
    screenwidth = root.winfo_screenwidth()
    screenheight = root.winfo_screenheight()
    width, height = root.maxsize() #把窗口重写成了最大
    size = '%dx%d+%d+%d' % (width, height-65,-10 ,0 )
    print (size)
    root.geometry(size)
'''
class Hide():
    def __init__(self,Master):
        Master.withdraw()
'''        
def quit_all():
    sys.exit(0)
'''
class Show():
    def __init__(self,Master):
        Master.update()
        Master.deiconify()
'''
 
class Create_main_window:
    def __init__(self,Master):
        center_window(Master, 1366, 768)  # 设置窗口居中，设置宽度和高度       
        Master.title("BenQGuru eHR Ass For Gaia7303")
        Master.resizable(True, True)
        Master.protocol("WM_DELETE_WINDOW", quit_all)
        self.mainframe=tk.Frame(Master)
        self.CONframe=tk.Frame(Master)
        
        self.label_payrollmode = tk.Label(self.CONframe,text='加载数据库配置')
        self.label_payrollmode.grid(row=0,column=0,sticky=tk.W)
        self.label_payrollmode = tk.Label(self.CONframe,text='服务器名称')
        self.label_payrollmode.grid(row=0,column=31,sticky=tk.W)
        self.label_payrollmode = tk.Label(self.CONframe,text='用户名')
        self.label_payrollmode.grid(row=0,column=33,sticky=tk.W)
        self.label_payrollmode = tk.Label(self.CONframe,text='用户密码')
        self.label_payrollmode.grid(row=0,column=35,sticky=tk.W)                        
        self.label_payrollmode = tk.Label(self.CONframe,text='数据库名')
        self.label_payrollmode.grid(row=0,column=37,sticky=tk.W)        
        self.label_payrollmode = tk.Label(self.CONframe,text='数据源名称')
        self.label_payrollmode.grid(row=0,column=39,sticky=tk.W)   
        self.entry_servername_var = tk.StringVar()
        self.entry_dbusername_var = tk.StringVar()
        self.entry_dbpsw_var = tk.StringVar()
        self.entry_dbname_var = tk.StringVar()
        self.entry_sourcename_var = tk.StringVar()
        self.entry_servername = tk.Entry(self.CONframe,textvariable=self.entry_servername_var)
        self.entry_dbusername = tk.Entry(self.CONframe,textvariable=self.entry_dbusername_var)
        self.entry_dbpsw = tk.Entry(self.CONframe,textvariable=self.entry_dbpsw_var)
        self.entry_dbname = tk.Entry(self.CONframe,textvariable=self.entry_dbname_var)
        self.entry_sourcename = tk.Entry(self.CONframe,textvariable=self.entry_sourcename_var)
        self.entry_servername.insert(-1," ")
        self.entry_dbusername.insert(-1," ")
        self.entry_dbpsw.insert(-1," ")
        self.entry_dbname.insert(-1," ")
        self.entry_sourcename.insert(-1," ")
        self.entry_servername.grid(row=0,column=32,sticky=tk.W+tk.E)
        self.entry_dbusername.grid(row=0,column=34,sticky=tk.W+tk.E)
        self.entry_dbpsw.grid(row=0,column=36,sticky=tk.W+tk.E)
        self.entry_dbname.grid(row=0,column=38,sticky=tk.W+tk.E)
        self.entry_sourcename.grid(row=0,column=40,sticky=tk.W+tk.E)

        
              
        #Master.iconbitmap('E:\工作文档\1608-中银\98客户环境\eHR\Web\Content\images\home\logo-gaia.png')        
#窗口滚动问题没解决 mainScrolly=tk.Scrollbar(Master)  mainScrolly.pack(side="right",fill="y") #mainScrolly.grid(row=22,column=1,sticky=tk.W)  scrolly.config(command=lb.yview)        
        self.label_payrollmode = tk.Label(self.mainframe,text='加载资源')
        self.label_payrollmode.grid(row=10,column=0,sticky=tk.W)
#此处点击按钮从数据库中更新公用代码        
        self.button_reoption = tk.Button(self.mainframe,text='加载公用代码')#command=self.show_root2
        self.button_reoption.grid(row=10,column=1,sticky=tk.E)
        
        self.label_option = tk.Label(self.mainframe,text='薪资公用代码')
        self.label_option.grid(row=11,column=0,sticky=tk.W)
        self.label_option = tk.Label(self.mainframe,text='薪资主逻辑')
        self.label_option.grid(row=12,column=0,sticky=tk.W)
        self.label_option = tk.Label(self.mainframe,text='计薪模式')
        self.label_option.grid(row=12,column=1,sticky=tk.W)


        selectMainlogic = """select mainlogicname from PAYCALCULATEMAINLOGIC"""
        db_mainlogic = connWithDB(ehr,selectMainlogic)
        selectPatternName = """SELECT PATTERNNAME FROM PAYACCOUNTPATTERN WHERE ISDELETED=0"""
        db_patternName = connWithDB(ehr,selectPatternName)
        
        self.listbox_option_mainlogic = tk.Listbox(self.mainframe,selectmode="browse",height=4)
        lb = self.listbox_option_mainlogic
        for item in db_mainlogic:
             lb.insert(-1,item)
        self.listbox_option_mainlogic.grid(row=13,column=0)
            
        self.listbox_option_patternName = tk.Listbox(self.mainframe,selectmode="browse",height=4)
        lb = self.listbox_option_patternName
        for item in db_patternName:
             lb.insert(-1,item)
        self.listbox_option_patternName.grid(row=13,column=1)

        self.label_option = tk.Label(self.mainframe,text='工作流配置')
        self.label_option.grid(row=20,column=0,sticky=tk.W)
#新建showAllSP列表框        
        self.label_option = tk.Label(self.mainframe,text='子流程阶段调用中的sp')
        self.label_option.grid(row=21,column=0,sticky=tk.W)                

        scrolly_showALLSP=tk.Scrollbar(self.mainframe,orient="vertical") 
        scrolly_showALLSP.grid(row=22,column=1,sticky=tk.N+tk.S+tk.W) 
#canva空间   convas_test = tk.Canvas(self.mainframe) convas_test.grid(row=23,column=0)        
        selectShowALLSP = """select sp_name from gbpm.fm_mdprocedure"""
        db_ShowALLSP = connWithDB(ehr,selectShowALLSP)
        self.listbox_flower_showAllSP = tk.Listbox(self.mainframe,selectmode="browse",yscrollcommand=scrolly_showALLSP.set)#,yscrollcommand=scrolly.set
        lb = self.listbox_flower_showAllSP
        for item in db_ShowALLSP:
             lb.insert(-1,item)
        self.listbox_flower_showAllSP.grid(row=22,column=0)    
        scrolly_showALLSP.config(command=lb.yview)    
#读取所有SP

#获取写入地址            
        self.label_option = tk.Label(self.mainframe,text='写入文件地址')
        self.label_option.grid(row=200,column=0,sticky=tk.W)
        self.entry_rootfile_var = tk.StringVar()
        self.entry_rootfile = tk.Entry(self.mainframe,textvariable=self.entry_rootfile_var)
        self.entry_rootfile.insert(-1,"E:\\workspace\\tmpfile")
        #self.entry_rootfile.place(x=50,y=50)
        self.entry_rootfile.grid(row=201,column=0,sticky=tk.W+tk.E)
        # 设置按钮
        self.but_readShowAllSP = tk.Button(self.mainframe, text="读取SP脚本", width=10,command=self.readShowALLSP(db_ShowALLSP))#
        self.but_readShowAllSP.grid(row=202, column=0,sticky=tk.W+tk.E+tk.N)
        self.CONframe.place(x=0,y=0)
        self.mainframe.place(x=0,y=20)
         
    def readShowALLSP(self,db_ShowALLSP):
        mb.showinfo('SP开始提取')
        db_SPContentList = []
        for item in db_ShowALLSP:
            if item[:4] != "gbpm":
                #print(item[:4])
                item = "gbpm." + item
            selectHelptext = """sp_helptext '"""+item+"""'"""
            print(selectHelptext)
            strHead = """\n-------------------TIP """+item+"""\nBEGIN----BEGINTIP"""
            db_SPContentList.append(strHead)
            db_helptext = connWithDB(ehr,selectHelptext)
            for item in db_helptext:
                db_SPContentList.append(item)
            strFoot = """\nEND----BEGINTIP"""
            db_SPContentList.append(strFoot)
        mb.showinfo('SP提取完毕')             
        filename_showALLSP = 'showALLSP.txt'
        filetruepath_showALLSP = str(self.entry_rootfile_var.get())+'/'+filename_showALLSP
        print(filetruepath_showALLSP)
        fileobj = open( filetruepath_showALLSP, 'w')
        for item in db_SPContentList:
            fileobj.write(item)
                
if __name__ == "__main__":
    count = 0
    root1 = tk.Tk()
    main_window = Create_main_window(root1)
    root1.mainloop()
     
        