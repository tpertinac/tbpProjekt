import subprocess
import xlwt
import pyoo
import time
from pymongo import MongoClient
from tkinter import *

class MiningCoins(Frame):

    def __init__(self, master):
        Frame.__init__(self, master)
        self.grid()
        self.createWidgets()

    def createWidgets(self):
        self.label = Label(self, text = "Dobrodosli!")
        self.label.grid()
        self.btnCoin = Button(self, text = "Prikaz po coinovima")
        self.btnCoin["command"] = self.coin
        self.btnCoin.grid()
        self.btnKat = Button(self, text = "Prikaz po kategorijama")
        self.btnKat["command"] = self.category
        self.btnKat.grid()
        self.btnGKat = Button(self, text="Graficki prikaz za kategorije")
        self.btnGKat["command"] = self.grafKat
        self.btnGKat.grid()

    def coin(self):
        client = MongoClient('mongodb://localhost:27017/')
        db = client['tbp']
        eth = db.ethereum.find()
        bit = db.bitcoin.find()
        lit = db.litecoin.find()
        mon = db.monero.find()
        rip = db.ripple.find()
        wb = xlwt.Workbook()

        sheetB = wb.add_sheet("Bitcoin")
        sheetE = wb.add_sheet("Etherium")
        sheetL = wb.add_sheet("Litecoin")
        sheetM = wb.add_sheet("Monero")
        sheetR = wb.add_sheet("Ripple")

        sheetB.write(0, 0, 'Datum')
        sheetE.write(0, 0, 'Datum')
        sheetL.write(0, 0, 'Datum')
        sheetM.write(0, 0, 'Datum')
        sheetR.write(0, 0, 'Datum')

        sheetB.write(0, 1, 'Open')
        sheetE.write(0, 1, 'Open')
        sheetL.write(0, 1, 'Open')
        sheetM.write(0, 1, 'Open')
        sheetR.write(0, 1, 'Open')

        sheetB.write(0, 2, 'High')
        sheetE.write(0, 2, 'High')
        sheetL.write(0, 2, 'High')
        sheetM.write(0, 2, 'High')
        sheetR.write(0, 2, 'High')

        sheetB.write(0, 3, 'Low')
        sheetE.write(0, 3, 'Low')
        sheetL.write(0, 3, 'Low')
        sheetM.write(0, 3, 'Low')
        sheetR.write(0, 3, 'Low')

        sheetB.write(0, 4, 'Close')
        sheetE.write(0, 4, 'Close')
        sheetL.write(0, 4, 'Close')
        sheetM.write(0, 4, 'Close')
        sheetR.write(0, 4, 'Close')

        sheetB.write(0, 5, 'Volume')
        sheetE.write(0, 5, 'Volume')
        sheetL.write(0, 5, 'Volume')
        sheetM.write(0, 5, 'Volume')
        sheetR.write(0, 5, 'Volume')

        sheetB.write(0, 6, 'Market Cap')
        sheetE.write(0, 6, 'Market Cap')
        sheetL.write(0, 6, 'Market Cap')
        sheetM.write(0, 6, 'Market Cap')
        sheetR.write(0, 6, 'Market Cap')

        i = 1
        for row in bit:
            sheetB.write(i, 0, str(row['Date']))
            sheetB.write(i, 1, float(row['Open']))
            sheetB.write(i, 2, float(row['High']))
            sheetB.write(i, 3, float(row['Low']))
            sheetB.write(i, 4, float(row['Close']))
            sheetB.write(i, 5, str(row['Volume']))
            sheetB.write(i, 6, str(row['Market Cap']))
            i = i +1

        i = 1
        for row in eth:
            sheetE.write(i, 0, str(row['Date']))
            sheetE.write(i, 1, float(row['Open']))
            sheetE.write(i, 2, float(row['High']))
            sheetE.write(i, 3, float(row['Low']))
            sheetE.write(i, 4, float(row['Close']))
            sheetE.write(i, 5, str(row['Volume']))
            sheetE.write(i, 6, str(row['Market Cap']))
            i = i + 1

        i = 1
        for row in lit:
            sheetL.write(i, 0, str(row['Date']))
            sheetL.write(i, 1, float(row['Open']))
            sheetL.write(i, 2, float(row['High']))
            sheetL.write(i, 3, float(row['Low']))
            sheetL.write(i, 4, float(row['Close']))
            sheetL.write(i, 5, str(row['Volume']))
            sheetL.write(i, 6, str(row['Market Cap']))
            i = i + 1

        i = 1
        for row in mon:
            sheetM.write(i, 0, str(row['Date']))
            sheetM.write(i, 1, float(row['Open']))
            sheetM.write(i, 2, float(row['High']))
            sheetM.write(i, 3, float(row['Low']))
            sheetM.write(i, 4, float(row['Close']))
            sheetM.write(i, 5, str(row['Volume']))
            sheetM.write(i, 6, str(row['Market Cap']))
            i = i + 1

        i = 1
        for row in rip:
            sheetR.write(i, 0, str(row['Date']))
            sheetR.write(i, 1, float(row['Open']))
            sheetR.write(i, 2, float(row['High']))
            sheetR.write(i, 3, float(row['Low']))
            sheetR.write(i, 4, float(row['Close']))
            sheetR.write(i, 5, str(row['Volume']))
            sheetR.write(i, 6, str(row['Market Cap']))
            i = i + 1

        wb.save("/home/tomek/Documents/coinovi.ods")
        subprocess.call(['/usr/bin/localc', '/home/tomek/Documents/coinovi.ods'])

    def category(self):
        client = MongoClient('mongodb://localhost:27017/')
        db = client['tbp']
        lim = db.ethereum.count()
        eth = db.ethereum.find()
        bit = db.bitcoin.find().limit(lim)
        lit = db.litecoin.find().limit(lim)
        mon = db.monero.find().limit(lim)
        rip = db.ripple.find().limit(lim)
        wb = xlwt.Workbook()
        sheetO = wb.add_sheet("Open")
        sheetH = wb.add_sheet("High")
        sheetL = wb.add_sheet("Low")
        sheetC = wb.add_sheet("Close")
        sheetV = wb.add_sheet("Volume")
        sheetMC = wb.add_sheet("Market Cap")

        sheetO.write(0, 0, 'Datum')
        sheetO.write(0, 1, 'Bitcoin')
        sheetO.write(0, 2, 'Ethereum')
        sheetO.write(0, 3, 'Litecoin')
        sheetO.write(0, 4, 'Monero')
        sheetO.write(0, 5, 'Ripple')

        sheetH.write(0, 0, 'Datum')
        sheetH.write(0, 1, 'Bitcoin')
        sheetH.write(0, 2, 'Ethereum')
        sheetH.write(0, 3, 'Litecoin')
        sheetH.write(0, 4, 'Monero')
        sheetH.write(0, 5, 'Ripple')

        sheetL.write(0, 0, 'Datum')
        sheetL.write(0, 1, 'Bitcoin')
        sheetL.write(0, 2, 'Ethereum')
        sheetL.write(0, 3, 'Litecoin')
        sheetL.write(0, 4, 'Monero')
        sheetL.write(0, 5, 'Ripple')

        sheetC.write(0, 0, 'Datum')
        sheetC.write(0, 1, 'Bitcoin')
        sheetC.write(0, 2, 'Ethereum')
        sheetC.write(0, 3, 'Litecoin')
        sheetC.write(0, 4, 'Monero')
        sheetC.write(0, 5, 'Ripple')

        sheetV.write(0, 0, 'Datum')
        sheetV.write(0, 1, 'Bitcoin')
        sheetV.write(0, 2, 'Ethereum')
        sheetV.write(0, 3, 'Litecoin')
        sheetV.write(0, 4, 'Monero')
        sheetV.write(0, 5, 'Ripple')

        sheetMC.write(0, 0, 'Datum')
        sheetMC.write(0, 1, 'Bitcoin')
        sheetMC.write(0, 2, 'Ethereum')
        sheetMC.write(0, 3, 'Litecoin')
        sheetMC.write(0, 4, 'Monero')
        sheetMC.write(0, 5, 'Ripple')

        i = 1
        for row in bit:
            sheetO.write(i, 0, str(row['Date']))
            sheetO.write(i, 1, float(row['Open']))
            sheetH.write(i, 0, str(row['Date']))
            sheetH.write(i, 1, float(row['High']))
            sheetL.write(i, 0, str(row['Date']))
            sheetL.write(i, 1, float(row['Low']))
            sheetC.write(i, 0, str(row['Date']))
            sheetC.write(i, 1, float(row['Close']))
            sheetV.write(i, 0, str(row['Date']))
            sheetV.write(i, 1, int(str(row['Volume']).replace(',','')))
            sheetMC.write(i, 0, str(row['Date']))
            sheetMC.write(i, 1, int(str(row['Market Cap']).replace(',','')))
            i = i + 1

        i = 1
        for row in eth:
            sheetO.write(i, 2, float(row['Open']))
            sheetH.write(i, 2, float(row['High']))
            sheetL.write(i, 2, float(row['Low']))
            sheetC.write(i, 2, float(row['Close']))
            sheetV.write(i, 2, int(str(row['Volume']).replace(',','')))
            sheetMC.write(i, 2, int(str(row['Market Cap']).replace(',','')))
            i = i + 1

        i = 1
        for row in lit:
            sheetO.write(i, 3, float(row['Open']))
            sheetH.write(i, 3, float(row['High']))
            sheetL.write(i, 3, float(row['Low']))
            sheetC.write(i, 3, float(row['Close']))
            sheetV.write(i, 3, int(str(row['Volume']).replace(',','')))
            sheetMC.write(i, 3, int(str(row['Market Cap']).replace(',','')))
            i = i + 1

        i = 1
        for row in mon:
            sheetO.write(i, 4, float(row['Open']))
            sheetH.write(i, 4, float(row['High']))
            sheetL.write(i, 4, float(row['Low']))
            sheetC.write(i, 4, float(row['Close']))
            sheetV.write(i, 4, int(str(row['Volume']).replace(',','')))
            sheetMC.write(i, 4, int(str(row['Market Cap']).replace(',','')))
            i = i + 1

        i = 1
        for row in rip:
            sheetO.write(i, 5, float(row['Open']))
            sheetH.write(i, 5, float(row['High']))
            sheetL.write(i, 5, float(row['Low']))
            sheetC.write(i, 5, float(row['Close']))
            sheetV.write(i, 5, int(str(row['Volume']).replace(',','')))
            sheetMC.write(i, 5, str(row['Market Cap']))
            i = i + 1

        wb.save("/home/tomek/Documents/kategorije.ods")
        subprocess.call(['/usr/bin/localc', '/home/tomek/Documents/kategorije.ods'])

    def grafKat(self):
            subprocess.Popen(['soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless'], shell=True)
            time.sleep(2)
            desktop = pyoo.Desktop('localhost', 2002)
            doc = desktop.open_spreadsheet('/home/tomek/Documents/kategorije.ods')
            sheet1 = doc.sheets[0]
            sheet2 = doc.sheets[1]
            sheet3 = doc.sheets[2]
            sheet4 = doc.sheets[3]
            sheet5 = doc.sheets[4]
            sheet6 = doc.sheets[5]

            chart = sheet1.charts.create('Open', sheet1[2:50, 7:18], sheet1[1:739, 1:6])
            diagram = chart.change_type(pyoo.LineDiagram)
            diagram.y_axis.title = "USD"
            diagram.y_axis.logarithmic = True

            chart = sheet2.charts.create('High', sheet2[2:50, 7:18], sheet2[1:739, 1:6])
            diagram = chart.change_type(pyoo.LineDiagram)
            diagram.y_axis.title = "USD"
            diagram.y_axis.logarithmic = True

            chart = sheet3.charts.create('Low', sheet3[2:50, 7:18], sheet3[1:739, 1:6])
            diagram = chart.change_type(pyoo.LineDiagram)
            diagram.y_axis.title = "USD"
            diagram.y_axis.logarithmic = True

            chart = sheet4.charts.create('Close', sheet4[2:50, 7:18], sheet4[1:739, 1:6])
            diagram = chart.change_type(pyoo.LineDiagram)
            diagram.y_axis.title = "USD"
            diagram.y_axis.logarithmic = True

            chart = sheet5.charts.create('Volume', sheet5[2:28, 7:18], sheet5[1:739, 1:6])
            diagram = chart.change_type(pyoo.LineDiagram)
            diagram.y_axis.title = "USD"
            diagram.y_axis.logarithmic = True

            chart = sheet6.charts.create('Market Cap', sheet6[2:28, 7:18], sheet6[1:739, 1:6])
            diagram = chart.change_type(pyoo.LineDiagram)
            diagram.y_axis.title = "USD"
            diagram.y_axis.logarithmic = True

            doc.save('/home/tomek/Documents/grafovi.ods')

root = Tk()
root.geometry("200x130")
root.title("TBP Projekt")
MiningCoins(root)
root.mainloop()