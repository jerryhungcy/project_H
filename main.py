

import tkinter as tk
import multiprocessing as mp
import tkinter.ttk as ttk
import subprocess
import os

ScriptData_2 =    {"Rfiles" : ["BCI", "FSI", "RAI", "shadow rate"],
                "BCI":[ "Australia","Brazil", "China", "japane", "Korean", "Turkey", "USA"],
                "FSI":["Brazil", "China", "Indonesia", "Japan", "Korean", "Malaysia",  "South Africa", "Thailand", "USA" ],
                "RAI":["China", "world"],
                "shadow rate":["China"]}

ScriptData = {"Rfiles":["BCI", "FSI", "RAI", "Shadow Rate"],
       "BCI":["Australia", "Brazil", "China", "Japan", "Korea", "Turkey", "USA"],
       "FSI":["Brazil", "China", "Indonesia", "Japan", "Korea", "Malaysia", "S. Africa", "Thailand", "USA"],
       "RAI":["China", "USA"],
       "Shadow Rate":["China"],
       "BCIAustralia":["AUSBCI.R", ["AusBCI.xlsx"]],
       "BCIBrazil":["BRABCI.R", ["BrazilBCI.xlsx"]],
       "BCIChina":["CBCI.R", ["MonthlyChinaBCI.xlsx","DailyChinaBCI.xlsx"]],
       "BCIJapan":["JAPBCI.R", ["JapanBCI.xlsx"]],
       "BCIKorea":["KOR.R", ["KoreaBCI.xlsx"]],
       "BCITurkey":["TURBCI.R", ["TurkeyBCI.xlsx"]],
       "BCIUSA":["USBCI.R", ["USDailyBCI.xlsx"]],
       "FSIChina":["CFSI.R", ["ChinaFSI.xlsx"]],
       "FSIUSA":["USFSI.R", ["USFSI.xlsx"]],
       "FSIBrazil":["BRAFSI.R", ["BRAFSI.xlsx"]],
       "FSIIndonesia":["INDFSI.R", ["INDFSI.xlsx"]],
       "FSIJapan":["JAPFSI.R",["JAPFSI.xlsx"]],
       "FSIKorea":["KOR.R", ["KORFSI.xlsx"]],
       "FSIMalaysia":["MALFSI.R", ["MALFSI.xlsx"]],
       "FSIS. Africa":["SAFFSI.R", ["SAFFSI.xlsx"]],
       "FSIThailand":["THAFSI.R", ["THAFSI.xlsx"]],
       "RAIChina":["CRAI.R", ["ChinaRAI.xlsx"]],
       "RAIUSA":["USRAI.R", ["USRAI.xlsx"]],
       "Shadow RateChina":["Shadow Rate.R", ["Daily Shadow Rate.xlsx", "ShadowRate-1DAheadForecast.xlsx"]]
}

R_path = "C:/Program Files/R/R-4.2.0/bin/x64/RScript.exe"


class App:
    def __init__(self, win):
        self.win = win

        self.Labels = []
        self.f1 = tk.Frame(win, bg = "green", height = 500, width = 500)
        self.f1.pack(side = 'right')
        self.f1.propagate(0)

        self.f2 = tk.Frame(win, bg = "blue", height = 500, width = 300)
        self.f2.pack(side = 'right')
        self.f2.propagate(0)

        self.cbScript = ttk.Combobox(self.f1, width= 15, state='readonly')
        self.cbScript.pack(side = 'left', anchor=tk.NW, padx= 10, pady= 2)
        self.cbScript['value'] = ScriptData["Rfiles"]
        self.cbScript.current(0)

        self.cbContry = ttk.Combobox(self.f1, width= 15, state='readonly')
        self.cbContry.pack(side = 'left', anchor=tk.NW, padx= 10, pady= 2)
        self.cbContry['value'] = ScriptData[ self.cbScript.get()]
        self.cbContry.current(0)

        self.btn = tk.Button(self.f1, text = "run", bg = "light gray",  font=("Arial", 12,"bold"), width = 8, command = self.runProcess)
        self.btn.pack(side = 'left', anchor= tk.NW, padx= 10, pady= 2)

        self.cbScript.bind('<<ComboboxSelected>>', self.set_cbContry)
    
        

    def set_cbContry(self, event):
        self.cbContry.pack(side = 'left', anchor=tk.NW)
        self.cbContry['value'] = ScriptData[ self.cbScript.get()]
        self.cbContry.current(0)
    
    def loop(self):
        self.win.mainloop()

    def showExcel(self, files):
        '''df = pd.read_excel(filename)
        f = Figure(figsize=(9,5), dpi=100)
        ax = f.add_subplot(111)
        df.plot(kind="line",ax=ax)
        '''
        for f in files:
            print(f)
            os.system("start EXCEL.EXE {}".format(f))
        

    def runProcess(self):
        #R_dst = "{} {}.r".format(self.cbContry.get(), self.cbScript.get())
        d =  self.cbScript.get() + self.cbContry.get()
        R_dst = ScriptData[d][0]
        t = "{} {}\n{}".format(self.cbContry.get(), self.cbScript.get(),R_dst)
        btn = tk.Button(self.f2, state = 'disable', text = t+"\nloading", fg = 'green', font=("Arial", 12,"bold"), width= 15, command = lambda:self.showExcel(ScriptData[d][1]))
        btn.pack(side = 'top', padx= 10, pady= 2)
        self.f2.update_idletasks()
        try:
            p = subprocess.run([R_path, R_dst,] , shell = True, check = True)
            if p.stdout == None:
                btn['text'] = t +'\nfinish'
                btn['state'] = 'normal'
        except subprocess.CalledProcessError:
            print("something went wrong")
            btn['text'] = t + '\nerror'
                  
       


if __name__ == '__main__':
    mp.freeze_support()
    win = tk.Tk()
    app = App(win)
    app.loop()