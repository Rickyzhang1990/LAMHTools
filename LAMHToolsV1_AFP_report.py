#!/usr/bin/env python3
#-*-coding:utf8-*-
## author:kaka Zhang 202407

from tkinter import *
from tkinter import filedialog, font
import time
import datetime
import sys,os
import polars as pl
import math
from tkcalendar import DateEntry
import tkinter.messagebox
import traceback
from reportimage import create_overlay_pdf,Examinate


class LAMH_GUI:
    def __init__(self,init_window_name):
        self.init_window = init_window_name
        ## batch model
        self.path = StringVar()
        self.CYcut = 28
        self.new_window = None
        self.batch_window  = None
        ## single model


    def selectPath(self):
        self.batch_window.attributes('-topmost',False)
        file_path = filedialog.askopenfilename()
        self.batch_window.attributes('-topmost',True)
        if file_path:
            self.path.set(file_path)

    def error(self, errmessage): #错误弹窗
        tkinter.messagebox.showinfo(title= "提示", message= errmessage)

    def uptime(self): #更新时间
        TimeLabel["text"] = datetime.datetime.now().strftime( '%Y-%m-%d %H:%M:%S')
        self.init_window.after(1000, self.uptime)



#设置窗口
    def set_init_window(self):
        font3 = font.Font(size=10,weight='bold')
        self.init_window.title("LAMH Tools") #窗口标题栏
        if os.path.exists('莱盟健康LOGO.ico'):
            self.init_window.iconbitmap('莱盟健康LOGO.ico')
        else:
            pass
        self.init_window.geometry( '500x400+300+200') #窗口大小，300x300是窗口的大小，+600是距离左边距的距离，+300是距离上边距的距离
        self.init_window.resizable(width=FALSE, height=FALSE) #限制窗口大小，拒绝用户调整边框大小
        Label(self.init_window, text="肝癌联合检测分析软件", bg="AliceBlue", fg="Blue",
              font=font.Font(size=18,weight="bold",)).place(relx=0.25, rely=0.15)
        Label(self.init_window,text="大健康版",bg="AliceBlue", fg="Blue",font=font.Font(size=12,weight="bold")
             ).place(relx=0.44,rely=0.25)
        Button(self.init_window, text="单人模式", font=font.Font(size=16,weight="bold"),relief = tkinter.RAISED,command=self.SingleModel).place(relx=0.25,rely=0.35,relwidth=0.5,relheight=0.15)
        Button(self.init_window, text="报告模式", font=font.Font(size=16,weight="bold"),relief = tkinter.RAISED,command=self.BatchModel).place(relx=0.25,rely=0.55,relwidth=0.5,relheight=0.15)
        self.init_window[ "bg"] = "AliceBlue"#窗口背景
        self.init_window.attributes( "-alpha", 0.95) #虚化，值越小虚化程度越高



        global TimeLabel
        TimeLabel = Label(self.init_window,text="%s"% (datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')),bg= "AliceBlue",font=font3) #标签组件，显示时间
        TimeLabel.place(relx= 0.33, rely= 0.75)
        self.init_window.after(1000, self.uptime)
        Label(self.init_window, text="©Copyright LAMH", fg="red",
              font=font.Font(size=8, weight="normal", underline=True, slant="italic")).place(relx=0.39, rely=0.9)

    def SingleModel(self):
        ## some new parameter
        self.riskscore = StringVar()
        self.result = StringVar()
#        self.note   = StringVar()
        self.name   = StringVar()
        self.gender = StringVar()
        self.age    = StringVar()
        self.fam    = StringVar()
        self.rox    = StringVar()
        self.vic    = StringVar()
        self.cy5    = StringVar()
        self.afp    = StringVar()

        if self.new_window is None or not self.new_window.winfo_exists():
            self.new_window = Toplevel(self.init_window)
            self.new_window.title("单人模式")
            self.new_window.geometry("500x400+400+200")
            self.new_window.resizable(width=TRUE, height=TRUE) # 窗口标题栏
            self.new_window[ "bg"] = "AliceBlue"#窗口背景
            self.new_window.attributes( "-alpha", 0.95) #虚化，值越小虚化程度越高
            self.new_window.attributes('-topmost', 1)
            if os.path.exists('莱盟健康LOGO.ico'):
                self.new_window.iconbitmap('莱盟健康LOGO.ico')
            else:
                pass
        ## 病人信息
            Label(self.new_window, text="病人信息", bg="AliceBlue",font=font.Font(size=14, weight="bold", underline=True)).place(relx=0.01,
                                                                                                         rely=0.01)
            Label(self.new_window, text=f"姓名",bg="AliceBlue").place(relx=0.02, rely=0.08, )
            name = Entry(self.new_window,textvariable=self.name)
            name.place(relx=0.1, rely=0.08, relwidth=0.2)
            Label(self.new_window, text=f"性别",bg="AliceBlue").place(relx=0.32, rely=0.08)
            gender = Entry(self.new_window,textvariable=self.gender)
            gender.place(relx=0.40, rely=0.08, relwidth=0.2)
            Label(self.new_window, text="年龄",bg="AliceBlue").place(relx=0.64, rely=0.08)
            age = Entry(self.new_window,textvariable=self.age)
            age.place(relx=0.72, rely=0.08, relwidth=0.2)
            Label(self.new_window, text="病人编号",bg="AliceBlue").place(relx=0.02, rely=0.16)
            Entry(self.new_window).place(relx=0.16, rely=0.16, relwidth=0.25)
            Label(self.new_window, text="联系方式" ,bg="AliceBlue").place(relx=0.43, rely=0.16)
            Entry(self.new_window).place(relx=0.57, rely=0.16, relwidth=0.35)
        ## 送检科室及信息
            Label(self.new_window, text="医院信息", bg="AliceBlue",font=font.Font(size=14, weight="bold", underline=True)).place(relx=0.01,
                                                                                                         rely=0.22)
            Label(self.new_window, text="送检科室",bg="AliceBlue").place(relx=0.02, rely=0.3)
            Entry(self.new_window).place(relx=0.16, rely=0.3, relwidth=0.27)
            Label(self.new_window, text="检测申请人",bg='AliceBlue').place(relx=0.44, rely=0.3)
            Entry(self.new_window).place(relx=0.615, rely=0.3, relwidth=0.30)
            Label(self.new_window, text="送检日期",bg="AliceBlue").place(relx=0.02, rely=0.38)
            DateEntry(self.new_window, width=12, background='darkblue', foreground='white', borderwidth=2).place(relx=0.16,
                                                                                                        rely=0.38,
                                                                                                        relwidth=0.4)
        ## 检测信息
            Label(self.new_window, text="检测信息",bg="AliceBlue", font=font.Font(size=14, weight="bold", underline=True)
              ).place(relx=0.01,rely=0.48)
            Label(self.new_window, text="FAM",bg="AliceBlue").place(relx=0.02, rely=0.56)
            fam = Entry(self.new_window,textvariable=self.fam)
            fam.place(relx=0.1, rely=0.56, relwidth=0.13)
            Label(self.new_window, text="ROX",bg="AliceBlue").place(relx=0.25, rely=0.56)
            rox = Entry(self.new_window,textvariable=self.rox)
            rox.place(relx=0.34, rely=0.56, relwidth=0.13)
            Label(self.new_window, text="VIC",bg="AliceBlue").place(relx=0.50, rely=0.56)
            vic = Entry(self.new_window,textvariable=self.vic)
            vic.place(relx=0.57, rely=0.56, relwidth=0.12)
            Label(self.new_window, text="CY5",bg="AliceBlue").place(relx=0.70, rely=0.56)
            cy5 = Entry(self.new_window,textvariable=self.cy5)
            cy5.place(relx=0.78, rely=0.56, relwidth=0.13)
            Label(self.new_window, text="AFP",bg="AliceBlue").place(relx=0.02, rely=0.63)
            afp = Entry(self.new_window,textvariable=self.afp)
            afp.place(relx=0.10, rely=0.63, relwidth=0.13)

        ##检测结果
            Button(self.new_window, text="计算风险值",bg="AliceBlue",fg="Navy", command=self.singledecision, font=font.Font(size=14, weight="bold")).place(
                                                                                                       relx=0.01,rely=0.72)
            Label(self.new_window, text="风险值",bg="AliceBlue").place(relx=0.01, rely=0.83)
            Entry(self.new_window, textvariable=self.riskscore,state='readonly').place(relx=0.1, rely=0.83, relwidth=0.34)
            Label(self.new_window, text="判定结果",bg="AliceBlue").place(relx=0.45, rely=0.83)
            Entry(self.new_window, textvariable=self.result,state='readonly').place(relx=0.57, rely=0.83, relwidth=0.34)
#            Label(self.new_window, text="注意事项",bg="AliceBlue").place(relx=0.01, rely=0.88)
#            Entry(self.new_window, textvariable = self.note,state='readonly').place(relx=0.15, rely=0.88, relwidth=.4)
            Label(self.new_window, text="©Copyright LAMH", fg="red",bg="AliceBlue",
              font=font.Font(size=8, weight="normal", underline=True, slant="italic")).place(relx=0.37, rely=0.93)
    def is_number(self,input_str):
        try:
            float(input_str)
            return True
        except ValueError:
            return False
    def checkinput(self):
        errorlist = []
#        print(self.name.get())
        if self.name.get() == "":
            errorlist.append("姓名")
        if not self.gender.get() in ['男','女']:
            errorlist.append("性别[男/女]")
        if self.age.get() == "":
            errorlist.append("年龄")
        if self.fam.get() != "" and self.is_number(self.fam.get()):
            pass
        else:
            errorlist.append("FAM")
        if self.rox.get() != "" and self.is_number(self.rox.get()):
            pass
        else:
            errorlist.append("ROX")
        if self.vic.get() != "" and self.is_number(self.vic.get()):
            pass
        else:
            errorlist.append("VIC")
        if self.cy5.get() != "" and self.is_number(self.cy5.get()) and float(self.cy5.get()) <= self.CYcut:
            pass
        elif float(self.cy5.get()) > self.CYcut:
            errorlist.append("Ct(Cy5)>28,检测结果不合格")
        else:
            errorlist.append("CY5")
        if self.afp.get() != "" and self.is_number(self.afp.get()):
            pass
#        elif not int(self.afp.get()) in [0,1]:
#            errorlist.append("AFP阳性为1阴性为0")
        else:
            errorlist.append("AFP")
        if len(errorlist) == 0:
            return True
        else:
            self.new_window.attributes('-topmost',False)
            errorinfo = "\n".join(errorlist)
            self.error(f"请正确输入以下信息：\n{errorinfo}")
            self.new_window.attributes('-topmost', True)
            return False
    def singledecision(self):
        '''
        :return:
        '''
#        print(self.fam,self.cy5)
        if self.checkinput():
            vfam = float(self.fam.get())
            vrox = float(self.rox.get())
            vvic = float(self.vic.get())
            vcy5 = float(self.cy5.get())
            vgender = 1 if self.gender.get() == "男" else 0
            ROX_delta, VIC_delta, FAM_delta = vrox - vcy5, vvic - vcy5, vfam - vcy5
#            if self.afp.get() == "":
#                lrscore = 1.796 - 0.115 * ROX_delta - 0.033 * VIC_delta  - 0.138 * FAM_delta + 1.272 * vgender
#                riskscore = max(100/(1 + math.exp(0.2 * (ROX_delta - 6.05))),
#                                100/(1 + math.exp(0.2 * (VIC_delta - 2.93))),
#                                100/(1 + math.exp(0.2 * (FAM_delta - 5.49))),
#                                100/(1 + math.exp(-2 * (lrscore - 0.65)))
#                                )
#                riskscore = round(riskscore,2)
#                result = "阳性" if riskscore >= 50 else "阴性"
#                note  = "不包含AFP风险值"
#                print(riskscore,result ,note)
#                self.riskscore.set(str(riskscore))
#                self.result.set(result)
#                self.note.set(note)
#            else:
            vafp = int(self.afp.get())
            if vafp in [0,1]:
                famcut = 4.53  #OTX1
                roxcut = 5.58  #RNF135
                viccut = 2.93  #H3C7
                lrcut  = 0.974
#                note   = "包含AFP风险值"
                lrscore = 0.654 - 0.102 * ROX_delta - 0.007 * VIC_delta - 0.104 * FAM_delta + 1.375 * vgender + 1.9 * vafp
                riskscore = max(100 / (1 + math.exp(0.3 * (ROX_delta - roxcut))),
                               100 / (1 + math.exp(0.2 * (VIC_delta - viccut))),
                               100 / (1 + math.exp(0.3 * (FAM_delta - famcut))),
                               100 / (1 + math.exp(-1.15 * (lrscore - lrcut)))
                               )
                riskscore = round(riskscore, 2)
                result = "高风险" if riskscore >= 50 else "低风险"
                self.riskscore.set(str(riskscore))
                self.result.set(result)
#                self.note.set(note)
            else:
                self.new_window.attributes('-topmost', False)
                self.error("AFP结果应当为有条带输入1,无条带输入0")
                self.new_window.attributes('-topmost', True)




    def BatchModel(self):
        font1 = font.Font(size=12)
        font2 = font.Font(size=16,weight="bold")
        if self.batch_window is None or not self.batch_window.winfo_exists():
            self.batch_window = Toplevel(self.init_window)
            self.batch_window.title("报告模式") #窗口标题栏
            self.batch_window.attributes('-topmost', 1)
            if os.path.exists('莱盟健康LOGO.ico'):
                self.batch_window.iconbitmap('莱盟健康LOGO.ico')
            else:
                pass
            self.batch_window.geometry('500x400+400+200') #窗口大小，300x300是窗口的大小，+600是距离左边距的距离，+300是距离上边距的距离
            self.batch_window.resizable(width=FALSE, height=FALSE) #限制窗口大小，拒绝用户调整边框大小
            Label(self.batch_window, text= "报告模式结果判定",bg="AliceBlue",fg = "Blue",font=font2).place(relx= 0.31,rely= 0.15) #标签组件，用来显示内容，place里的x、y是标签组件放置的位置
            Label(self.batch_window, text= "检测结果文件:",bg= "AliceBlue",font=font.Font(size=12)).place(relx= 0.06,rely= 0.3625) #标签组件，用来显示内容，place里的x、y是标签组件放置的位置
            Entry(self.batch_window, textvariable=self.path).place(relwidth = 0.48,relheight = 0.075,relx= 0.28,rely= 0.35) #输入控件，用于显示简单的文本内容
            Button(self.batch_window, text= "路径选择", command=self.selectPath,bg= "AliceBlue",font=font1).place(relx= 0.78,rely= 0.35,relwidth=0.14) #按钮组件，用来触发某个功能或函数
            Button(self.batch_window, text= "生成结果及报告",bg= "AliceBlue",font=font1,command=self.decision).place(relwidth= 0.48,relheight= 0.1,relx= 0.28,rely= 0.55) #按钮组件，用来触发某个功能或函数
            self.batch_window[ "bg"] = "AliceBlue"#窗口背景
            self.batch_window.attributes( "-alpha", 0.95) #虚化，值越小虚化程度越高
            Label(self.batch_window, text="©Copyright LAMH", fg="red",bg="AliceBlue",
              font=font.Font(size=8, weight="normal", underline=True, slant="italic")).place(relx=0.4, rely=0.95)
#获取当前时间
    def get_current_time(self):
        current_time = time.strftime( '%Y-%m-%d %H:%M:%S',time.localtime(time.time))
        return current_time

    def isnumeric_pl(self,s):
        try:
            pl.Series(s).cast(pl.Float64)
            return True
        except ValueError:
            return False

    def decision(self):
        '''
        :return:result of liver cancer detection
        '''
        if self.path.get() == "":
            self.batch_window.attributes("-topmost",False)
            self.error("文件位置为空，请选择检测结果文件")
            self.batch_window.attributes("-topmost",True)
        else:
            dataname = os.path.basename(self.path.get())
            dirname  = os.path.dirname(self.path.get())
            filename = dataname.rstrip(".xls") if dataname.endswith("xls") else dataname.rstrip(".xlsx")
            ROXcut = 5.58 ## RNF135
            VICcut = 2.93 ## H3C7
            FAMcut = 4.53 ## OTX1
            LRcut  = 0.974 ## LR model score
            genderdict = {"男":1,"女":0}
            data = pl.read_excel(self.path.get())
            data.columns = [i.strip().upper() for i in data.columns]
            if "AFP" in data.columns:
#        print(data.columns)
                data = data.with_columns((pl.col('性别').map_dict(genderdict)).alias("gender"))
                data = data.with_columns(
                [pl.col(col).str.to_uppercase().alias(col) if data[col].dtype == pl.Utf8 else pl.col(col).alias(col) for col
                 in ['FAM', 'VIC', 'ROX', 'CY5']])
                data = data.with_columns(
                [pl.col(col).str.replace("NOCT", 40).cast(pl.Float64).alias(col) for col in ['FAM', 'VIC', 'ROX', 'CY5']])
                mask_FAM = self.isnumeric_pl(data['FAM'])
                mask_ROX = self.isnumeric_pl(data['ROX'])
                mask_VIC = self.isnumeric_pl(data['VIC'])
                mask_CY5 = self.isnumeric_pl(data['CY5'])
                mask_AFP = self.isnumeric_pl(data['AFP'])
                check_na_expr = (
                (pl.col("性别").is_not_null())&
                (pl.col("FAM").is_not_null())&
                (pl.col("ROX").is_not_null())&
                (pl.col('VIC').is_not_null())&
                (pl.col("CY5").is_not_null())&
                (pl.col('AFP').is_not_null())&
                (pl.col('AFP').is_in([0,1])) &
                (pl.col("性别").is_in(['男','女']))&
                mask_FAM &
                mask_ROX &
                mask_VIC &
                mask_CY5 &
                mask_AFP &
                (pl.col("CY5") <= 28)
                )
                data_ok = data.filter(check_na_expr)
                data_bad = data.filter(~check_na_expr)
#                print(data)
                data_ok = data_ok.with_columns([(pl.col('ROX') - pl.col('CY5')).alias("deltaROX"),
                                  (pl.col('FAM') - pl.col('CY5')).alias("deltaFAM"),
                                  (pl.col('VIC') - pl.col('CY5')).alias("deltaVIC")])
#        data = data.with_columns(pl.struct(['性别', 'FAM', 'VIC', 'ROX', 'CY5']).map_elements(
#            lambda x: ",".join([i + " has NA" for i in [col for col, value in x.items() if str(value) == ""]])).alias(
#            "Note"))
        ## ok data
                data_ok = data_ok.with_columns([(0.654 - 0.102 * pl.col('deltaROX') - 0.007 * pl.col('deltaVIC') - 0.104 * pl.col(
                "deltaFAM") + 1.375 * pl.col("gender") + 1.9 * pl.col("AFP")).alias("LR")])
  #      data = data.with_columns(pl.struct(['deltaRNF135', 'deltaOTX1', 'deltaH3C7', 'LRModel']).map_elements(
  #                              lambda x : "Pos" if any([x['deltaRNF135'] <= 6.05, x['deltaH3C7'] <= 2.63, x['deltaOTX1'] <= 5.49,
  #                                                      x['LRModel'] >0.666]) else "Neg").alias("qPCR_result"))
                data_ok = data_ok.with_columns(pl.col("deltaVIC").map_elements(
                lambda x: 100 / (1 + math.exp(0.2 * (x - VICcut)))).alias("VIC_score"))
                data_ok = data_ok.with_columns(pl.col("deltaROX").map_elements(
                lambda x: 100 / (1 + math.exp(0.3 * (x - ROXcut)))).alias("ROX_score"))
                data_ok = data_ok.with_columns(pl.col("deltaFAM").map_elements(
                lambda x: 100 / (1 + math.exp(0.3 * (x - FAMcut)))).alias("FAM_score"))
                data_ok = data_ok.with_columns(pl.col("LR").map_elements(
                lambda x: 100 / (1 + math.exp(-1.15 * (x - LRcut)))).alias("LR_score"))
                data_ok = data_ok.with_columns(pl.max_horizontal(['ROX_score', 'VIC_score', 'FAM_score', 'LR_score']).alias("风险值"))
#            data_ok = data_ok.with_columns(pl.when(
#                        pl.col("CY5") <=28
#                        ).then(
#                        pl.col("风险值") * 0
#                        ).otherwise(
#                        pl.col("风险值") * 1
#                        )
#                    )
                data_ok  = data_ok.with_columns(pl.col("风险值").round(2).alias("风险值"))
                data_ok = data_ok.with_columns(pl.col('风险值').map_elements(lambda x:"高风险" if x >= 50 else "低风险").alias("判定结果"))
                note_ok = pl.Series(name = "Note", values= [None] * data_ok.shape[0],dtype=pl.Utf8)
#            data_ok = data_ok.with_columns(pl.col('note').map_elements(lambda x:"实验不合格" if pl.col("CY5") > 28 else None).alias("note"))
                data_ok = data_ok.with_columns(note_ok)
                selectcol = ['姓名', '性别', '年龄', '病人编号', '联系方式', '检测科室', '检测申请人', '检测日期', 'AFP',
                         'FAM','VIC', 'ROX', 'CY5', '风险值', '判定结果', 'Note']
                for row in data_ok.iter_rows():
                    person = Examinate()
                    person.name = row[data_ok.columns.index('姓名')]
                    person.age = row[data_ok.columns.index('年龄')]
                    person.gender = row[data_ok.columns.index('性别')]
                    person.date1 = row[data_ok.columns.index('检测日期')]
                    person.date2 = row[data_ok.columns.index('检测日期')]
                    person.date3 = row[data_ok.columns.index('检测日期')]
                    person.code = row[data_ok.columns.index('病人编号')]
                    person.blood = "血浆"
                    person.riskscore = row[data_ok.columns.index('风险值')]
                    person.hospital  = row[data_ok.columns.index('检测科室')]
                    outpath = os.path.dirname(str(self.path.get()))
                    pdf = create_overlay_pdf(person,outpath)
                    os.makedirs(outpath +"/report/",exist_ok=True)
                    with open(outpath + f"/report/{person.code}_{person.name}_{person.date3}_肝癌qPCR基因检测报告.pdf",
                              'wb') as output_pdf:
                        pdf.write(output_pdf)
        ## data bad
                if not data_bad.is_empty():
                    riskscore = pl.Series(name="风险值",values=[None] * data_bad.shape[0],dtype=pl.Float64)
                    anresult  = pl.Series(name="判定结果", values=[None] * data_bad.shape[0],dtype=pl.Utf8)
                    data_bad  = data_bad.with_columns([riskscore,anresult])
                    data_bad  = data_bad.with_columns(pl.struct(['性别','FAM','ROX','VIC','CY5','AFP']).map_elements(
                    lambda x: [col for col, value in x.items() if value is None]).alias("Note")
                                        )
                    data_bad  = data_bad.with_columns(pl.col('Note').map_elements(lambda x:"请输入:" + ",".join(x)).alias("Note"))
                    data_bad  = data_bad.with_columns(pl.when(pl.col("CY5") >28)
                                            .then("Ct(Cy5)>28, 检测结果不合格")
                                            .otherwise(pl.col("Note"))
                                            .alias("Note")
                                            )
#
                    data_all  = pl.concat([data_ok.select(selectcol),data_bad.select(selectcol)])
                else:
                    data_all  = data_ok.select(selectcol)
                if not os.path.exists(f"{dirname}/{filename}_qPCR_result.xlsx"):
                    data_all.write_excel(f"{dirname}/{filename}_qPCR_result.xlsx")
                    self.batch_window.attributes('-topmost',False)
                    tkinter.messagebox.showinfo(title="提示", message="结果已生成")
                    self.batch_window.attributes('-topmost', True)
                else:
                    self.batch_window.attributes('-topmost', False)
                    self.error("结果文件已存在,请先删除同名文件")
                    self.batch_window.attributes('-topmost', True)
            else:
                self.batch_window.attributes('-topmost', False)
                self.error("大健康版需要AFP结果")
                self.batch_window.attributes('-topmost',True)



if __name__ == '__main__':
    try:
        init_window = Tk() # 实例化出一个父窗口
        LAMH_PORTAL = LAMH_GUI(init_window)
        LAMH_PORTAL.set_init_window()
        init_window.mainloop()
    except Exception as e:
        with open(sys.path[0] +"/error_log.txt",'w') as f:
            f.write(traceback.format_exc())
