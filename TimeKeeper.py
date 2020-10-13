import tkinter as tk
from datetime import datetime, date
import re
import pandas
from openpyxl import load_workbook


class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.timeIN = 0
        self.timeOut = 0
        self.hoursWorked = 0
        self.minsWorked = 0
        self.date = 0
        self.master = master
        self.pack()
        self.create_widgets()

    def create_widgets(self):
        self.header = tk.Label(self,
                               text="CLOCK IN : CLOCK OUT",
                               height=25,
                               width=200)
        self.header.pack()

        self.timeIn = tk.Button(self,
                                text="TIME IN",
                                command=self.punchIN)

        self.timeIn.pack(side='left')

        self.timeOut = tk.Button(self,
                                 text="TIME OUT",
                                 command=self.punchOUT)

        self.timeOut.pack(side='right')

    def round(self, val):
        if val < 25: return 25
        if val < 50: return 50
        if val < 75: return 75

    def punchIN(self):
        t = datetime.now()
        self.timeIN = t.strftime("%H:%M:%S")

    def punchOUT(self):
        if self.timeIN == 0:
            self.forgot = tk.Label(self,
                                   text="FORGOT TO CLOCK IN \n ENTER TIME IN")
            self.forgot.pack()

            self.forgotEnter = tk.Entry(self)
            self.forgotEnter.pack()
            root.bind('<Return>', self.fix(self.forgotEnter.get()))
        t = datetime.now()
        self.timeOut = t.strftime("%H:%M:%S")
        self.cal(self.timeIN, self.timeOut)

    def fix(self, entry):
        if str(entry).isnumeric():
            if len(entry) < 4:
                entry = '0' + entry
            self.timeIN = entry[:2] + ':' + entry[2:] + ':00'

    def cal(self, timein, timeout):
        tempout = int(re.sub('[^0-9]','',timeout))
        tempin = int(re.sub('[^0-9]','',timein))
        tempdif = tempout - tempin
        self.hoursWorked = int(tempdif / 3600)
        self.minsWorked = tempdif / 60 / 60 * 100
        self.minsWorked = round(self.minsWorked)
        self.header["text"] = "YOU HAVE WORKED\n" + str(self.hoursWorked) + ':' + str(self.minsWorked)
        self.appendToXlsx(self.timeIN, self.timeOut, self.hoursWorked, self.minsWorked)

    def appendToXlsx(idk, In, Out, Hours, Mins, startrow=None, sheet_name='Sheet1', **to_excel_kwargs):
        path = r'C:\Users\Keith Rich\Documents\PythonScripts\TimeCard\TimeInTimeOut.xlsx'
        book = load_workbook(path)
        writer = pandas.ExcelWriter(path, engine = 'openpyxl')
        writer.book = book
        if startrow is None and sheet_name in writer.book.sheetnames:
            startrow = writer.book[sheet_name].max_row
        writer.sheets = {ws.title:ws for ws in writer.book.worksheets}

        if startrow is None:
            startrow = 0

        df = {"IN": [In], "OUT": [Out], "HOURS:MINS": [str(Hours) + ":" + str(Mins)]}
        info = pandas.DataFrame(data=df, columns=['IN', 'OUT', 'HOURS:MINS'], index=[str(datetime.date(datetime.now()))])

        info.to_excel(writer, 'Sheet1', startrow=startrow, **to_excel_kwargs)
        writer.save()
        writer.close()


root = tk.Tk()
app = Application(master=root)
app.mainloop()
