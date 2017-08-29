import subprocess
import time
import xlwt
import pyoo
import os
import tkinter.messagebox
from tkinter import *
from pymongo import MongoClient


class MiningCoins(Frame):

    def __init__(self, master):
        try:
            #subprocess.Popen(['service mongod start'], shell=True)
            #time.sleep(5)
            client = MongoClient('mongodb://localhost:27017/')
            db = client['tbp']
        except ConnectionRefusedError:
            tkinter.messageBox.showinfo('MongoDB connection error', 'Nije moguće se spojiti na MongoDB!')
        frame = Frame(master)
        frame.grid()
        label = Label(frame, text = "Dobrodosli!")
        label.grid(row=0, column=1)
        btnCoin = Button(frame, text = "Prikaz po coinovima", command=lambda: self.coin(db))
        btnCoin.grid(row=1)
        btnKat = Button(frame, text = "Prikaz po kategorijama", command=lambda: self.category(db))
        btnKat.grid(row=2)
        btnGKat = Button(frame, text="Graficki prikaz za kategorije", command=self.grafKat)
        btnGKat.grid(row=2, column=1)
        lbGod = Listbox(frame, height=3, exportselection=0)
        lbGod.insert(0, "2017")
        lbGod.insert(1, "2016")
        lbGod.insert(2, "2015")
        lbGod.grid(row = 3)
        lbKat = Listbox(frame, height=6, exportselection=0)
        lbKat.insert(0, "Open")
        lbKat.insert(1, "High")
        lbKat.insert(2, "Low")
        lbKat.insert(3, "Close")
        lbKat.insert(4, "Volume")
        lbKat.insert(5, "Market Cap")
        lbKat.grid(row=3, column=1)
        lbCoi = Listbox(frame, height=5, exportselection=0)
        lbCoi.insert(0, "Bitcoin")
        lbCoi.insert(1, "Ethereum")
        lbCoi.insert(2, "Litecoin")
        lbCoi.insert(3, "Monero")
        lbCoi.insert(4, "Ripple")
        lbCoi.grid(row=3, column=2)
        btnP = Button(frame, text="Napravi pojedinačnu analizu", command=lambda: self.analiza(db, lbGod.get(ACTIVE), lbKat.get(ACTIVE), lbCoi.get(ACTIVE)))
        btnP.grid(row=4, column=1)
        btnQuit = Button(frame, text="Izađi", command=frame.quit)
        btnQuit.grid(row=5, column=4)

    def coin(self, db):
        bit = db.bitcoin.find()
        eth = db.ethereum.find()
        lit = db.litecoin.find()
        mon = db.monero.find()
        rip = db.ripple.find()

        cursors = list([bit, eth, lit, mon, rip])

        wb = xlwt.Workbook()

        sheetB = wb.add_sheet("Bitcoin")
        sheetE = wb.add_sheet("Etherium")
        sheetL = wb.add_sheet("Litecoin")
        sheetM = wb.add_sheet("Monero")
        sheetR = wb.add_sheet("Ripple")

        sheet = list([sheetB, sheetE, sheetL, sheetM, sheetR])

        for s in sheet:
            s.write(0, 0, 'Datum')
            s.write(0, 1, 'Open')
            s.write(0, 2, 'High')
            s.write(0, 3, 'Low')
            s.write(0, 4, 'Close')
            s.write(0, 5, 'Volume')
            s.write(0, 6, 'Market Cap')


        i = 0
        for c in cursors:
            j = 1
            for row in c:
                sheet[i].write(j, 0, str(row['Date']))
                sheet[i].write(j, 1, float(row['Open']))
                sheet[i].write(j, 2, float(row['High']))
                sheet[i].write(j, 3, float(row['Low']))
                sheet[i].write(j, 4, float(row['Close']))
                sheet[i].write(j, 5, str(row['Volume']))
                sheet[i].write(j, 6, str(row['Market Cap']))
                j = j + 1
            i = i + 1

        reportDir = os.path.dirname(os.path.abspath('reports')) + "/reports/"
        wb.save(reportDir + "coinovi.ods")
        subprocess.call(['/usr/bin/localc', reportDir + 'coinovi.ods'])

    def category(self, db):
        lim = db.ethereum.count()

        eth = db.ethereum.find()
        bit = db.bitcoin.find().limit(lim)
        lit = db.litecoin.find().limit(lim)
        mon = db.monero.find().limit(lim)
        rip = db.ripple.find().limit(lim)

        cursors = list([bit, eth, lit, mon, rip])

        wb = xlwt.Workbook()

        sheetO = wb.add_sheet("Open")
        sheetH = wb.add_sheet("High")
        sheetL = wb.add_sheet("Low")
        sheetC = wb.add_sheet("Close")
        sheetV = wb.add_sheet("Volume")
        sheetMC = wb.add_sheet("Market Cap")

        sheet = list([sheetO, sheetH, sheetL, sheetC, sheetV, sheetMC])

        for s in sheet:
            s.write(0, 0, 'Datum')
            s.write(0, 1, 'Bitcoin')
            s.write(0, 2, 'Ethereum')
            s.write(0, 3, 'Litecoin')
            s.write(0, 4, 'Monero')
            s.write(0, 5, 'Ripple')

        i = 1
        for row in bit:
            for s in sheet:
                s.write(i, 0, str(row['Date']))
            sheet[0].write(i, 1, float(row['Open']))
            sheet[1].write(i, 1, float(row['High']))
            sheet[2].write(i, 1, float(row['Low']))
            sheet[3].write(i, 1, float(row['Close']))
            sheet[4].write(i, 1, int(str(row['Volume']).replace(',', '')))
            sheet[5].write(i, 1, int(str(row['Market Cap']).replace(',', '')))
            i = i + 1

        j = 1
        for c in cursors:
            i = 1
            for row in c:
                sheet[0].write(i, j, float(row['Open']))
                sheet[1].write(i, j, float(row['High']))
                sheet[2].write(i, j, float(row['Low']))
                sheet[3].write(i, j, float(row['Close']))
                sheet[4].write(i, j, int(str(row['Volume']).replace(',','')))
                sheet[5].write(i, j, int(str(row['Market Cap']).replace(',','')))
                i = i + 1
            j = j + 1

        reportDir = os.path.dirname(os.path.abspath('reports')) + "/reports/"
        wb.save(reportDir + "kategorije.ods")
        subprocess.call(['/usr/bin/localc', reportDir + 'kategorije.ods'])

    def grafKat(self):
            try:
                subprocess.Popen(['soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless'],
                                 shell=True)
                time.sleep(2)
                desktop = pyoo.Desktop('localhost', 2002)
                reportKat = os.path.dirname(os.path.abspath('reports')) + "/reports/kategorije.ods"
                doc = desktop.open_spreadsheet(reportKat)

                sheet1 = doc.sheets[0]
                sheet2 = doc.sheets[1]
                sheet3 = doc.sheets[2]
                sheet4 = doc.sheets[3]
                sheet5 = doc.sheets[4]
                sheet6 = doc.sheets[5]

                chart = sheet1.charts.create('Open', sheet1[2:50, 7:18], sheet1[0:740, 0:6], row_header=TRUE, col_header=TRUE)
                diagram = chart.change_type(pyoo.LineDiagram)
                diagram.y_axis.title = "USD"
                diagram.y_axis.logarithmic = True

                chart = sheet2.charts.create('High', sheet2[2:50, 7:18], sheet2[0:740, 0:6], row_header=TRUE, col_header=TRUE)
                diagram = chart.change_type(pyoo.LineDiagram)
                diagram.y_axis.title = "USD"
                diagram.y_axis.logarithmic = True

                chart = sheet3.charts.create('Low', sheet3[2:50, 7:18], sheet3[0:740, 0:6], row_header=TRUE, col_header=TRUE)
                diagram = chart.change_type(pyoo.LineDiagram)
                diagram.y_axis.title = "USD"
                diagram.y_axis.logarithmic = True

                chart = sheet4.charts.create('Close', sheet4[2:50, 7:18], sheet4[0:740, 0:6], row_header=TRUE, col_header=TRUE)
                diagram = chart.change_type(pyoo.LineDiagram)
                diagram.y_axis.title = "USD"
                diagram.y_axis.logarithmic = True

                chart = sheet5.charts.create('Volume', sheet5[2:28, 7:18], sheet5[0:740, 0:6], row_header=TRUE, col_header=TRUE)
                diagram = chart.change_type(pyoo.LineDiagram)
                diagram.y_axis.title = "USD"
                diagram.y_axis.logarithmic = True

                chart = sheet6.charts.create('Market Cap', sheet6[2:28, 7:18], sheet6[0:740, 0:6], row_header=TRUE, col_header=TRUE)
                diagram = chart.change_type(pyoo.LineDiagram)
                diagram.y_axis.title = "USD"
                diagram.y_axis.logarithmic = True

                reportDir = os.path.dirname(os.path.abspath('reports')) + "/reports/"
                doc.save(reportDir + 'grafovi.ods')
            except:
                tkinter.messagebox.showinfo('Dokument ne postoji', 'Potrebno je prvo generirati prikaz kategorija!')

    def analiza(self, db, god, cat, coi):

        if coi == 'Bitcoin':
            cursor = db.bitcoin.find({'Date': {'$regex': god}}, {'Date' : 1, cat: 1})

        elif coi == 'Ethereum':
            cursor = db.ethereum.find({'Date': {'$regex': god}}, {'Date' : 1, cat: 1})

        elif coi == 'Litecoin':
            cursor = db.litecoin.find({'Date': {'$regex': god}}, {'Date' : 1, cat: 1})

        elif coi == 'Monero':
            cursor = db.monero.find({'Date': {'$regex': god}}, {'Date' : 1, cat: 1})

        elif coi == 'Ripple':
            cursor = db.ripple.find({'Date': {'$regex': god}}, {'Date' : 1, cat: 1})

        wb = xlwt.Workbook()

        sheet = wb.add_sheet(coi)

        sheet.write(0, 0, 'Datum')
        sheet.write(0, 1, cat)


        i = 1
        for row in cursor:
            sheet.write(i, 0, str(row['Date']))
            if (cat == 'Volume' or cat == 'Market Cap'):
                sheet.write(i, 1, int(str(row[cat]).replace(',', '')))
            else:
                sheet.write(i, 1, float(row[cat]))
            i = i + 1

        reportDir = os.path.dirname(os.path.abspath('reports')) + "/reports/"
        wb.save(reportDir + "analiza.ods")

        subprocess.Popen(['soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless'],
                         shell=True)
        time.sleep(2)
        desktop = pyoo.Desktop('localhost', 2002)
        reportKat = os.path.dirname(os.path.abspath('reports')) + "/reports/analiza.ods"
        doc = desktop.open_spreadsheet(reportKat)

        sheet1 = doc.sheets[0]

        chart = sheet1.charts.create(cat, sheet1[2:20, 4:20], sheet1[0:cursor.count()+1, 0:2], row_header=TRUE, col_header=TRUE)
        diagram = chart.change_type(pyoo.LineDiagram)
        diagram.y_axis.title = "USD"
        diagram.y_axis.logarithmic = True

        reportDir = os.path.dirname(os.path.abspath('reports')) + "/reports/"
        doc.save(reportDir + 'analiza.ods')

root = Tk()
root.title("TBP Projekt")
mc = MiningCoins(root)
root.mainloop()