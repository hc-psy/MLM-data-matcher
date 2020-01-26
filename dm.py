#!/usr/bin/env python
# coding: utf-8



from tkinter import Frame, Tk, BOTH, Text, Menu, END, filedialog
import tkinter as tk
import xlwings as xw
from xlwings.constants import DeleteShiftDirection

class Gui(Frame):
    def __init__(self, parent):
        Frame.__init__(self, parent)   
        self.book = None
        self.parent = parent        
        self.initUI()
    
    def initUI(self):
            
        def onOpen1():

            ftypes = [('Excel files (.xls, .xlsx)', ('*.xlsx', '*.xls')), ('All files', '*')]
            dlg = filedialog.Open(self, filetypes = ftypes)
            fl = dlg.show()

            if fl != '':
                T.insert(END,fl)
                b = str(T.get(1.0,END)).replace('/',"\\\\")
                self.book = xw.Book(fl)
        
        def generatesb_name_in_sp():
            temp = int(n_sb1.get())
            
            
            for i in range(temp):
                tk.Label(p2, text="主管填第{}部屬姓名欄".format(i+1)).grid(row=2*i,column=0)
                sb_name_in_sp1.append(tk.Entry(p2))
                sb_name_in_sp1[-1].grid(row=2*i,column=1)
                sp_rate_sb1.append([])
                tk.Label(p2, text="主管評第{}部屬變項欄(起)".format(i+1)).grid(row=1+2*i,column=0)
                sp_rate_sb1[-1].append(tk.Entry(p2))
                sp_rate_sb1[-1][-1].grid(row=1+2*i,column=1)
                tk.Label(p2, text="主管評第{}部屬變項欄(迄)".format(i+1)).grid(row=1+2*i,column=2)
                sp_rate_sb1[-1].append(tk.Entry(p2))
                sp_rate_sb1[-1][-1].grid(row=1+2*i,column=3)

        def run():
            sheet = self.book.sheets[0]
            sheet.activate()
            complete = self.book.sheets.add(name="Complete", after=self.book.sheets[-1])
            
            #inputs
            headerbool = headerchecker.get()   
            n = int(n1.get())           
            n_sb = int(n_sb1.get())    
            sp_or_sb = str(sp_or_sb1.get())     
            sp_name = str(sp_name1.get())      
            sb_name = str(sb_name1.get())      
            sp_name_in_sb = str(sp_name_in_sb1.get()) 
            sb_name_in_sp = [str(i.get()) for i in sb_name_in_sp1]
            sb_rate_sp = [str(sb_rate_sp1.get()),str(sb_rate_sp2.get())]
            sb_self_rate = [str(sb_self_rate1.get()),str(sb_self_rate2.get())]
            sb_geo = [str(sb_geo1.get()),str(sb_geo2.get())]
            sp_rate_sb = [[str(j.get()) for j in i] for i in sp_rate_sb1]
            sp_self_rate = [str(sp_self_rate1.get()),str(sp_self_rate2.get())]
            sp_geo = [str(sp_geo1.get()),str(sp_geo2.get())]
        


            counter = 0
            pid = 0
            gid = 1
            
            #header
            if headerbool:
                i = 1
                complete[counter,0].value = "ID"
                complete[counter,1].value = 'Group ID'
                complete[counter,2].value = "主管"+sheet.range('{}{}'.format(sp_name,i)).value
                complete[counter,3].value = '部屬'+sheet.range('{}{}'.format(sb_name,i)).value
                
                if sb_rate_sp[0] and sb_rate_sp[1]:
                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sb_rate_sp[0],i,sb_rate_sp[1],i)).value
                if sb_self_rate[0] and sb_self_rate[1]:
                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sb_self_rate[0],i,sb_self_rate[1],i)).value
                if sb_geo[0] and sb_geo[1]:
                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sb_geo[0],i,sb_geo[1],i)).value
                
                complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sp_rate_sb[0][0],i,sp_rate_sb[0][1],i)).value
                if sp_self_rate[0] and sp_self_rate[1]:
                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sp_self_rate[0],i,sp_self_rate[1],i)).value
                if sp_geo[0] and sp_geo[1]:
                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sp_geo[0],i,sp_geo[1],i)).value

                counter += 1
                
            gidtemp = False    
            supnum = 0
            
            #check each sup
            for i in range(headerbool+1,headerbool+n+1):
                if sheet.range('{}{}'.format(sp_or_sb,i)).value == 1:
                    supnum+=1
                    
                    if gidtemp:
                        gid+=1
                    gidtemp = False
                    
                    for j in range(n_sb):
                        complete[counter,2].value = sheet.range('{}{}'.format(sp_name,i)).value
                        for k in range(headerbool+1,headerbool+n+1):
                            if sheet.range('{}{}'.format(sb_name_in_sp[j],i)).value == None:
                                break
                            if sheet.range('{}{}'.format(sb_name_in_sp[j],i)).value == sheet.range('{}{}'.format(sb_name,k)).value and sheet.range('{}{}'.format(sp_name_in_sb,k)).value == sheet.range('{}{}'.format(sp_name,i)).value:
                                gidtemp = True
                                pid +=1
                                complete[counter,0].value = pid
                                complete[counter,1].value = gid
                                complete[counter,3].value = sheet.range('{}{}'.format(sb_name,k)).value
                                
                                sheet.range('{}{}'.format(sb_name_in_sp[j],i)).color = (225,225,0)
                                sheet.range('{}{}'.format(sb_name,k)).color = (225,225,0)
                                
                                if sb_rate_sp[0] and sb_rate_sp[1]:
                                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sb_rate_sp[0],k,sb_rate_sp[1],k)).value
                                if sb_self_rate[0] and sb_self_rate[1]:
                                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sb_self_rate[0],k,sb_self_rate[1],k)).value
                                if sb_geo[0] and sb_geo[1]:
                                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sb_geo[0],k,sb_geo[1],k)).value
                                
                                complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sp_rate_sb[j][0],i,sp_rate_sb[j][1],i)).value
                                if sp_self_rate[0] and sp_self_rate[1]:
                                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sp_self_rate[0],i,sp_self_rate[1],i)).value
                                if sp_geo[0] and sp_geo[1]:
                                    complete[counter,complete[counter,0].end("right").column].value = sheet.range('{}{}:{}{}'.format(sp_geo[0],i,sp_geo[1],i)).value

                                break
                        
                        counter+=1
            
            #delete blank rows
            x, y = headerbool+1, headerbool+1
            while y <= headerbool+1+supnum*n_sb:
                if complete.range('A{}'.format(x)).value == None:
                    complete.range('{}:{}'.format(x,x)).api.Delete(DeleteShiftDirection.xlShiftUp)
                    y+=1
                    continue
                x+=1
                y+=1
        
        #GUI
        self.parent.title("Data Matcher by Hao-Cheng Lo 2017")
        self.pack()

        menubar = Menu(self.parent)
        self.parent.config(menu=menubar)
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Open", command=onOpen1)
        menubar.add_cascade(label="File", menu=fileMenu)
        menubar.add_cascade(label="Information")
        
        intro = '''歡迎使用一對多配對小工具，請注意下列事項
        =========================================================================
        1. 請將主管與部屬的資料剪貼至同一個Excel表中(第一個工作表)
        2. 主管和部屬的資料，儘量使用不同欄(這是讓你方便查找)
        3. 請在下列輸入欄名(如:A,B,C)，其中[必]為必填欄位
        4. 如果某類變項只有一欄，請於起迄填入相同欄名(如:M;M)
        5. 如果沒有主管評部屬變項，請於起迄欄名皆填入(XFD)
        6. 被成功配對的部屬問卷，部屬問卷中的部屬姓名會被標記為黃底
        7. 被成功配對的主管問卷，主管問卷中的部屬姓名會被標記為黃底'''
        
        header_label = tk.Label(self.parent, text=intro)
        header_label.pack()
        
        headerchecker = tk.IntVar()
        
        headercheckerbutton = tk.Checkbutton(self.parent, text = "資料包含標題",variable = headerchecker)
        headercheckerbutton.pack()
        
        # data
        data = Frame(self.parent)
        data.pack(side='top')
        
        tk.Label(data, text='未配對樣本數[必]').grid(row=0,column=0)
        n1 = tk.Entry(data)
        n1.grid(row=0,column=1)
        
        tk.Label(data, text='一對幾對偶(最大10)[必]').grid(row=0,column=2)
        n_sb1 = tk.Entry(data)
        n_sb1.grid(row=0,column=3)

        tk.Label(data, text='主管部屬判別欄[必]').grid(row=1,column=0)
        sp_or_sb1 = tk.Entry(data)
        sp_or_sb1.grid(row=1,column=1)
        
        tk.Label(data, text="主管自評姓名欄[必]").grid(row=1,column=2)
        sp_name1 = tk.Entry(data)
        sp_name1.grid(row=1,column=3)
        
        tk.Label(data, text="部屬自評姓名欄[必]").grid(row=4,column=0)
        sb_name1 = tk.Entry(data)
        sb_name1.grid(row=4,column=1)
        
        tk.Label(data, text="部屬填主管姓名欄[必]").grid(row=4,column=2)
        sp_name_in_sb1 = tk.Entry(data)
        sp_name_in_sb1.grid(row=4,column=3)
        
        tk.Label(data, text="部屬評主管變項欄(起)").grid(row=5,column=0)
        sb_rate_sp1 = tk.Entry(data)
        sb_rate_sp1.grid(row=5,column=1)
        tk.Label(data, text="部屬評主管變項欄(迄)").grid(row=5,column=2)
        sb_rate_sp2 = tk.Entry(data)
        sb_rate_sp2.grid(row=5,column=3)
        
        tk.Label(data, text="部屬自評變項欄(起)").grid(row=6,column=0)
        sb_self_rate1 = tk.Entry(data)
        sb_self_rate1.grid(row=6,column=1)
        tk.Label(data, text="部屬自評變項欄(迄)").grid(row=6,column=2)
        sb_self_rate2 = tk.Entry(data)
        sb_self_rate2.grid(row=6,column=3)
        
        tk.Label(data, text="部屬基本變項欄(起)").grid(row=7,column=0)
        sb_geo1 = tk.Entry(data)
        sb_geo1.grid(row=7,column=1)
        tk.Label(data, text="部屬基本變項欄(迄)").grid(row=7,column=2)
        sb_geo2 = tk.Entry(data)
        sb_geo2.grid(row=7,column=3)
        
        tk.Label(data, text="主管自評變項欄(起)").grid(row=8,column=0)
        sp_self_rate1 = tk.Entry(data)
        sp_self_rate1.grid(row=8,column=1)
        tk.Label(data, text="主管自評變項欄(迄)").grid(row=8,column=2)
        sp_self_rate2 = tk.Entry(data)
        sp_self_rate2.grid(row=8,column=3)
        
        tk.Label(data, text="主管基本變項欄(起)").grid(row=9,column=0)
        sp_geo1 = tk.Entry(data)
        sp_geo1.grid(row=9,column=1)
        tk.Label(data, text="主管基本變項欄(迄)").grid(row=9,column=2)
        sp_geo2 = tk.Entry(data)
        sp_geo2.grid(row=9,column=3)
        
        sb_name_in_sp1 = []
        sp_rate_sb1 = []
        
        open_sp_rate_sb = tk.Button(self.parent, text='主管評部屬變項，請先輸入一對幾對偶[必]', command=generatesb_name_in_sp)
        open_sp_rate_sb.pack()
        
        p2 = Frame(self.parent)
        p2.pack()
        
        operate = tk.Button(self.parent, text='檢查無誤，開始配對', command=run)
        operate.pack()
        
        T = Text(self.parent, height=2, width=100)
        T.pack()

def main():

    root = Tk()
    gui = Gui(root)
    root.geometry("600x850")
    root.mainloop()  

if __name__ == '__main__':
    main()  





