

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
       "BCIAustralia":["AUSBCI.gau", ["AusBCI.xlsx"]],
       "BCIBrazil":["BRABCI.gau", ["BrazilBCI.xlsx"]],
       "BCIChina":["CBCI.R", ["MonthlyChinaBCI.xlsx","DailyChinaBCI.xlsx"]],
       "BCIJapan":["JAPBCI.gau", ["JapanBCI.xlsx"]],
       "BCIKorea":["KOR.gau", ["KoreaBCI.xlsx"]],
       "BCITurkey":["TURBCI.gau", ["TurkeyBCI.xlsx"]],
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
       "Shadow RateChina":["Shadow Rate.R", ["\"Daily Shadow Rate.xlsx\"", "\"ShadowRate-1DAheadForecast.xlsx\""]]
}

#R_path = "C:/Program Files/R/R-4.2.0/bin/x64/RScript.exe"

R_path = "C:/Program Files/R/R-3.6.1/bin/x64/RScript.exe"
G_path = "C:/gauss19/gauss.exe"
G_file_path = "C:/Users/chueng/Google Drive/PRC1/"


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

        self.current_dir = os.getcwd().replace('\\','/')
        
    
        

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
        print(R_dst)
        t = "{} {}\n{}".format(self.cbContry.get(), self.cbScript.get(),R_dst)
        btn = tk.Button(self.f2, state = 'disable', text = t+"\nloading", fg = 'green', font=("Arial", 12,"bold"), width= 15, command = lambda:self.showExcel(ScriptData[d][1]))
        btn.pack(side = 'top', padx= 10, pady= 2)
        self.f2.update_idletasks()
        
        try:
            if d == "BCIChina":
                p = subprocess.run([R_path, "CBCI.R", "&&", G_path, G_file_path+"CBCI_M.gau", "&&", G_path, G_file_path+"CBCI_D.gau",] , shell = True, check = True)
            elif d == "BCIUSA":
                p = subprocess.run([R_path, "USBCI.R", "&&", G_path, G_file_path+"USBCI.gau", ] , shell = True, check = True)
            elif self.cbScript.get() == "BCI":
                G_file = G_file_path + R_dst
                p = subprocess.run([G_path, G_file, ] , shell = True, check = True)
                #p = subprocess.run([R_path, R_dst, "&&", R_path, "shadow Rate.R", ] , shell = True, check = True)
            else :
                p = subprocess.run([R_path, R_dst,] , shell = True, check = True)
            if p.stdout == None:
                btn['text'] = t +'\nfinished'
                btn['state'] = 'normal'
        except subprocess.CalledProcessError:
            print("something went wrong")
            btn['text'] = t + '\nfailed'
            btn['fg'] = 'red'
                  
       


if __name__ == '__main__':
    mp.freeze_support()
    win = tk.Tk()
    app = App(win)
    app.loop()
