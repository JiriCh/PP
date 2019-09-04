
import os
import sys
import time
import shutil
import sqlite3
import smtplib
import getpass
import requests
import calendar
import unidecode
import pythoncom
from win32com import client
from email import encoders
from PyQt5.QtGui import *
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from mailmerge import MailMerge
from operator import itemgetter
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from matplotlib.figure import Figure
from pandas import DataFrame, concat, merge, read_excel, read_csv
from email.mime.multipart import MIMEMultipart
from datetime import datetime, timedelta, date
from email.utils import COMMASPACE, formatdate
from dateutil.relativedelta import relativedelta

# Set fixed commission rate here:
FixedCommission = 0.1

sysDate = datetime.now().strftime('%Y-%m-%d')
dt = datetime.strptime(sysDate,'%Y-%m-%d')
start = dt - timedelta(days=dt.weekday())
end = start + timedelta(days=7)
weekStart = start.strftime('%Y-%m-%d %H:%M:%S')
weekEnd = end.strftime('%Y-%m-%d %H:%M:%S')
mS = dt - timedelta(days=int(datetime.now().strftime('%d')) - 1)
pmS = mS - relativedelta(months=1)
pmonthStart = pmS.strftime('%Y-%m-%d %H:%M:%S')
monthStart = mS.strftime('%Y-%m-%d %H:%M:%S')
mE = dt + timedelta(days=1 + calendar.monthrange(int(datetime.now().strftime('%Y')),int(datetime.now().strftime('%m')))[1]-int(datetime.now().strftime('%d')))
pmE = mS
monthEnd = mE.strftime('%Y-%m-%d %H:%M:%S')
pmonthEnd = pmE.strftime('%Y-%m-%d %H:%M:%S')
ipmE = dt - timedelta(days=int(datetime.now().strftime('%d')) - 0.99999)
ipmonthEnd = ipmE.strftime('%Y-%m-%d %H:%M:%S')
dDate = (date.today() + timedelta(days=7)).strftime('%Y-%m-%d')

inputs = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\Locked\\Inputs\\'
desktop = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\'
locked = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\Locked\\'
backup = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\Locked\\BackUp\\'
tempfolder = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\Locked\\TempFolder\\'
solditems = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\Locked\\SoldItemsInvoices\\'
rents = 'C:\\Users\\' + getpass.getuser() + '\\Desktop\\Locked\\RentsInvoices\\'
   
# PDF files
def PDF_creator(obj, word):
    filename = os.fsdecode(obj)
    if obj.endswith(".docx"):
        out_name = obj.replace("docx", r"pdf")
        in_file = tempfolder + filename
        if 'R-' in filename:
            out_file = rents + out_name
        if 'S-' in filename:
            out_file = solditems + out_name
        doc = word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=17)
        doc.Close()
        os.remove(in_file)

def sql_query(sql, args):
    if 'SELECT' in sql:
        global sql_obj, headers
        if args is None:
            args = []
        conn = sqlite3.connect(locked + '\\DataBase.db')
        c = conn.cursor()
        c.execute(sql, args)
        sql_obj = c.fetchall()
        headers = list(map(lambda x: x[0], c.description))
        c.close()
        conn.close()
    if 'INSERT' in sql or 'UPDATE' in sql or 'DELETE' in sql:
        conn = sqlite3.connect(locked + '\\DataBase.db')
        c = conn.cursor()
        c.execute(sql, args)
        conn.commit()
        c.close()
        conn.close()

month_dict = {
    '-01-' : ' Jan',
    '-02-' : ' Feb',
    '-03-' : ' Mar',
    '-04-' : ' Apr',
    '-05-' : ' May',
    '-06-' : ' Jun',
    '-07-' : ' Jul',
    '-08-' : ' Aug',
    '-09-' : ' Sep',
    '-10-' : ' Oct',
    '-11-' : ' Nov',
    '-12-' : ' Dec'
    }

cost_dict = {
    ' ' : 'NAP01',
    'Material' : 'MAT01',
    'Employee salary' : 'EMS01',
    'Sellers payoff' : 'SPO01',
    'Rent' : 'RNT01',
    'Postage' : 'POS01',
    'Other' : 'RES01'
    }

sql_dict = {
    'Table_of_widgets' : {'sql' : 'SELECT Category, Label, Product_desc, Product_id, Unit_price from Stock where Usage_flag not like "N"', 'args' : None},
    'SavePurchase' : {'sql' : 'INSERT INTO Purchases VALUES(?,?,?,?,?,?,?,?,?)', 'args': None},
    'SaveCosts' : {'sql' : 'INSERT INTO Costs VALUES(?,?,?,?,?,?,?,?,?,?,?)', 'args' : None},
    'CancelCheck' : {'sql' : 'SELECT Identifier from purchases', 'args' : None},
    'CancelReq' : {'sql' : 'UPDATE Purchases SET Flag = "Y" where Identifier = ?', 'args' : None},
    'LastPurchase' : {'sql' : 'SELECT Category, Label, Product_desc, Store, Amount, purchases.Product_id, purchases.Unit_price, Total_price, Purchase_time, Identifier, Flag, Payment_type from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Purchase_time = (SELECT max(Purchase_time) from purchases)', 'args': None},
    'PurchasesToday' : {'sql' : 'SELECT Category, Label, Product_desc, Store, Amount, purchases.Product_id, purchases.Unit_price, Total_price, Purchase_time, Identifier, Flag, Payment_type from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Purchase_time like ?', 'args' : ['%'+sysDate+'%']},
    'PurchasesThisWeek' : {'sql' : 'SELECT Category, Label, Product_desc, Store, Amount, purchases.Product_id, purchases.Unit_price, Total_price, Purchase_time, Identifier, Flag, Payment_type from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Purchase_time >= ? and Purchase_time < ?', 'args' : (weekStart, weekEnd)},
    'PurchasesThisMonth' : {'sql' : 'SELECT Category, Label, Product_desc, Store, Amount, purchases.Product_id, purchases.Unit_price, Total_price, Purchase_time, Identifier, Flag, Payment_type from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Purchase_time >= ? and Purchase_time < ?', 'args' : (monthStart, monthEnd)},
    'PurchasesLastMonth' : {'sql' : 'SELECT Category, Label, Product_desc, Store, Amount, purchases.Product_id, purchases.Unit_price, Total_price, Purchase_time, Identifier, Flag, Payment_type from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Purchase_time >= ? and Purchase_time < ?', 'args' : (pmonthStart, pmonthEnd)},
    'AllPurchases' : {'sql' : 'SELECT Category, Label, Product_desc, Store, Amount, purchases.Product_id, purchases.Unit_price, Total_price, Purchase_time, Identifier, Flag, Payment_type from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id', 'args' : None},
    'LastCost' : {'sql' : 'SELECT * from Costs where Cost_time = (select max(Cost_time) from costs)', 'args' : None},
    'ThisWeekCosts' : {'sql' : 'SELECT * from Costs where Cost_time >= ? and Cost_time < ?', 'args' : (weekStart, weekEnd)},
    'ThisMonthCosts' : {'sql' : 'SELECT * from Costs where Cost_time >= ? and Cost_time < ?', 'args' : (monthStart, monthEnd)},
    'LastMonthCosts' : {'sql' : 'SELECT * from Costs where Cost_time >= ? and Cost_time < ?', 'args' : (pmonthStart, pmonthEnd)},
    'AllCosts' : {'sql' : 'SELECT * from Costs', 'args' : None},
    'DisplayActStock' : {'sql' : 'SELECT * from Stock where Product_id not in ("post01")', 'args' : None},
    'Dups_check' : {'sql' : 'SELECT Category, Label, Product_desc, Product_id, Unit_price from Stock', 'args' : None},
    'UpdateProd' : {'sql' : 'UPDATE Stock SET Category = ?, Product_id = ?, Product_desc = ?, Label = ?, Unit_price = ?, Stock = ?, Added_by = ?, Usage_flag = ? where Product_id = ?', 'args' : None},
    'RegUpProd' : {'sql' : 'INSERT INTO Stock VALUES(?,?,?,?,?,?,?,?)', 'args' : None},
    'DisplayZero' : {'sql' : 'SELECT * from Stock where Product_id = ?', 'args' : ['']},
    'DisplayModStock' : {'sql' : 'SELECT * from Stock', 'args' : None},
    'LogIn' : {'sql' : 'SELECT * from users', 'args' : None},
    'ChangeP' : {'sql' : 'UPDATE users SET Password = ? where user = ?', 'args' : None},
    'RemoveU' : {'sql' : 'DELETE from users where user = ?', 'args' : None},
    'AddU' : {'sql' : 'INSERT INTO users VALUES(?,?)', 'args' : None},
    'addInfo' : {'sql' : 'SELECT Category, Label, Product_desc from Stock where Product_id = ?', 'args' : None},
    'Dynamic_sql' : {'sql' : 'SELECT Category, Label as Seller, Store, ? as Selected_period, sum(Amount) as Total_amount, sum(Total_price) as Total_sales from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Flag is not "Y" and Purchase_time like ? group by Category, Label, Store', 'args' : None},
    'Dynamic_sql_R' : {'sql' : 'SELECT "Nájemné" as Category, Rent_history.Seller, Rent_types.Place as Store, ? as Selected_period, sum(Rent_history.Rent_amount) as Total_amount, sum(Rent_history.Rent_value) as Total_sales from Rent_history LEFT JOIN Rent_types on Rent_history.Rent_ID = Rent_types.Rent_ID where Outstanding is "N" and Rent_history.Event_time like ? group by Category, Seller, Store', 'args' : None},
    'Dynamic_sql_OL' : {'sql' : 'SELECT Category, Label as Seller, Store, ? as Selected_period, sum(Amount) as Total_amount, sum(Total_price) as Total_sales from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Flag is not "Y" and Label in ("BonMotýl", "PP", "Lezecké dárky", "Česká pošta", "Výprodej") and Purchase_time like ? group by Category, Label, Store', 'args' : None},
    'Aggregated_view' : {'sql' : 'SELECT Category, Label as Seller, Store, ? as Selected_period, sum(Amount) as Total_amount, sum(Total_price) as Total_sales from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Flag is not "Y" group by Category, Label, Store', 'args' : None},
    'Aggregated_view_R' : {'sql' : 'SELECT "Nájemné" as Category, Rent_history.Seller, Rent_types.Place as Store, ? as Selected_period, sum(Rent_history.Rent_amount) as Total_amount, sum(Rent_history.Rent_value) as Total_sales from Rent_history LEFT JOIN Rent_types on Rent_history.Rent_ID = Rent_types.Rent_ID where Outstanding is "N" group by Category, Seller, Store', 'args' : None},
    'Aggregated_view_OL' : {'sql' : 'SELECT Category, Label as Seller, Store, ? as Selected_period, sum(Amount) as Total_amount, sum(Total_price) as Total_sales from Purchases LEFT JOIN Stock on stock.Product_id = purchases.Product_id where Flag is not "Y" and Label in ("BonMotýl", "PP", "Lezecké dárky", "Česká pošta", "Výprodej") group by Category, Label, Store', 'args' : None},
    'PMSoldItems' : {'sql' : 'SELECT stock.Label, stock.Product_desc, sum(Purchases.Amount) as Number_of_pieces, case when Rent_payers.Comment = "Commission" then (1 - "%s") * sum(Purchases.Total_price) else sum(Purchases.Total_price) end as Total_value from Purchases LEFT JOIN Stock on Purchases.Product_ID = stock.Product_id LEFT JOIN Rent_payers on stock.Label = Rent_payers.Seller where Stock.Label not in ("Česká pošta", "PP", "BonMotýl", "Lezecké dárky") and ? <= purchases.Purchase_time and purchases.Purchase_time < ? group by stock.Label, stock.Product_desc' %FixedCommission, 'args' : (pmonthStart, pmonthEnd)},
    'Rent_addressee' : {'sql' : 'SELECT Rent_payers.Seller, Sellers.EM_address, Sellers.Full_name, Sellers.Street, Sellers.Town, Sellers.Postcode, Sellers.Country, Sellers.ID_number, Rent_payers.Contract_since, Sellers.Phone_number, sum(Rent_payers.Rent_Amount * Rent_types.Rent_value) as Paid_rent from Rent_payers LEFT JOIN Sellers on Rent_payers.Seller = Sellers.Seller LEFT JOIN Rent_types on Rent_payers.Rent_id = Rent_types.Rent_id where Rent_payers.Active_payer is "Y" and Rent_payers.Rent_amount is not 0 and Rent_payers.Comment is not "Commission" group by Rent_payers.Seller, Sellers.EM_address, Sellers.Full_name, Sellers.Street, Sellers.Town, Sellers.Postcode, Sellers.Country, Sellers.ID_number, Rent_payers.Contract_since, Sellers.Phone_number', 'args' : None},
    'rActData' : {'sql' : 'SELECT Rent_payers.Rent_ID, Rent_payers.Rent_Amount, Rent_payers.Seller, "Bank Transfer" as Payment_type, "Y" as Outstanding, (Rent_payers.Rent_Amount * Rent_types.Rent_value) as Rent_value from Rent_payers join Rent_types on Rent_payers.Rent_ID = Rent_types.Rent_ID where Rent_payers.Active_payer is "Y" and Rent_payers.Rent_amount is not 0 and Rent_payers.Comment is not "Commission"', 'args' : None},
    'NewRents' : {'sql' : 'INSERT INTO Rent_history VALUES(?,?,?,?,?,?,?)', 'args': None},
    'sPayOff' : {'sql' : 'SELECT stock.Label, "Sellers payoff" as Item, "CZK" as Unit, "1" as Volume, "Bank Transfer" as Payment_type, "Costs" as Type, "SPO01" as Cost_ID, case when Rent_payers.Comment = "Commission" then (1 - "%s") * sum(Purchases.Total_price) else sum(Purchases.Total_price) end as Amount_paid from Purchases LEFT JOIN Stock on Purchases.Product_ID = stock.Product_id LEFT JOIN Rent_payers on stock.Label = Rent_payers.Seller where Stock.Label not in ("Česká pošta", "PP", "BonMotýl", "Lezecké dárky") and ? <= purchases.Purchase_time and purchases.Purchase_time < ? group by stock.Label, Item, Unit, Volume, Payment_type, Type, Cost_ID' %FixedCommission, 'args' : (pmonthStart, pmonthEnd)},
    'NewCosts' : {'sql' : 'INSERT INTO Costs VALUES(?,?,?,?,?,?,?,?,?,?,?)', 'args': None},
    'All_sellers' : {'sql' : 'SELECT Sellers.Seller, Sellers.EM_address from Sellers where Sellers.EM_address is not Null', 'args' : None},
    'InsertRentPayers' : {'sql' : 'INSERT into Rent_payers VALUES(?,?,?,?,?,?)', 'args' : None},
    'InsertSellers' : {'sql' : 'INSERT into Sellers VALUES(?,?,?,?,?,?,?,?,?,?,?,?)', 'args' : None},
    'ValidateSeller' : {'sql' : 'SELECT distinct Seller from Sellers', 'args' : None},
    'DisplayModPayers' : {'sql' : 'SELECT * from Rent_payers', 'args' : None},
    'DisplayModRents' : {'sql' : 'SELECT * from Rent_history where Outstanding = "Y"', 'args' : None},
    'UpdateRents' : {'sql' : 'UPDATE Rent_history SET Rent_ID = ?, Rent_amount = ?, Seller = ?, Event_time = ?, Rent_value = ?, Payment_type = ?, Outstanding = ? where Seller = ? and Event_time = ?', 'args' : None},
    'ChangeModPayers' : {'sql' : 'INSERT into Rent_payers VALUES(?,?,?,?,?,?)', 'args' : None},
    'DeleteModPayers' : {'sql' : 'DELETE from Rent_payers where Seller is not ?', 'args' : None},
    'DistinctPayers' : {'sql' : 'SELECT distinct Seller from Rent_payers', 'args' : None},
    'InsertNewPayer' : {'sql' : 'INSERT into Rent_payers VALUES(?,?,?,?,?,?)', 'args' : None},
    'DisplayRes' : {'sql' : 'SELECT * from Stock where Product_id not in ("post01") LIMIT 5', 'args' : None},
    'CheckStock' : {'sql' : 'SELECT * from Stock where Product_id  in (%s)', 'args' : None}
            }

class myCombo(QComboBox):
    def __init__(self, parent = None):
        super(myCombo, self).__init__(parent)
        self.setStyleSheet("QComboBox{"
                           "font-size:11px;"
                           "color:black;"
                           "background-color:white;"
                           "border:1px solid black;"
                           "padding:1px;""}")

        self.setEditable(True)
        self.lineEdit().setAlignment(Qt.AlignCenter)
        self.lineEdit().setReadOnly(True)

class mySpinbox(QSpinBox):
    def __init__(self, parent = None):
        super(mySpinbox, self).__init__(parent)
        self.lineEdit().setAlignment(Qt.AlignCenter)
        self.lineEdit().setReadOnly(True)

class CheckableComboBox(QComboBox):
    def __init__(self, parent = None):
        super(CheckableComboBox, self).__init__(parent)
        self.view().pressed.connect(self.handleItemPressed)
        self.setModel(QStandardItemModel(self))

        self.setStyleSheet("QComboBox{"
                           "font-size:11px;"
                           "color:black;"
                           "background-color:white;"
                           "border:1px solid black;"
                           "padding:1px;""}")

        self.setEditable(True)
        self.lineEdit().setAlignment(Qt.AlignCenter)
        self.lineEdit().setReadOnly(True)
        self.setMinimumSize(150, 40)

    def handleItemPressed(self, index):
        item = self.model().itemFromIndex(index)
        if item.text() != ' ' and item.checkState() == Qt.Checked:
            item.setCheckState(Qt.Unchecked)
        if item.text() != ' ' and item.checkState() != Qt.Checked:
            item.setCheckState(Qt.Checked)

class log_in(QMainWindow):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setMinimumSize(QSize(340, 140))
        self.setWindowTitle("Log-in")

        self.userName = QLabel(self)
        self.userName.setText('User name: ')
        self.userName.move(20, 20)

        self.nameInput = QLineEdit(self)
        self.nameInput.setFixedWidth(200)
        self.nameInput.move(110, 20)

        self.passWord = QLabel(self)
        self.passWord.setText('Password: ')
        self.passWord.move(20, 60)

        self.passInput = QLineEdit(self)
        self.passInput.setFixedWidth(200)
        self.passInput.setEchoMode(QLineEdit.Password)
        self.passInput.move(110, 60)

        okButton = QPushButton('OK', self)
        okButton.clicked.connect(self.okClicked)
        okButton.resize(100, 32)
        okButton.move(110, 100)

        cancelB = QPushButton('Cancel', self)
        cancelB.clicked.connect(self.cancelClicked)
        cancelB.resize(100, 32)
        cancelB.move(210, 100)

        self.Esc = QShortcut(QKeySequence("Escape"), self)
        self.Esc.activated.connect(self.cancelClicked)

        self.Enter = QShortcut(QKeySequence(Qt.Key_Return), self)
        self.Enter.activated.connect(self.okClicked)
        sql_query(sql_dict['LogIn']['sql'], sql_dict['LogIn']['args'])
        self.logDict = dict(sql_obj)

    def okClicked(self):
        if self.nameInput.text() in self.logDict.keys() and self.passInput.text() == self.logDict[self.nameInput.text()]:
            global uNm
            uNm = self.nameInput.text()
            self.initialW = appMW()
            self.initialW.showMaximized()
            self.close()
        else:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Invalid User name and/or Password, please re-try.', QMessageBox.Ok)
            msgbox.exec()

    def cancelClicked(self):
        self.close()

class appMW(QMainWindow):
    def __init__(self, parent = None):
        super(appMW, self).__init__(parent)
        self.title = 'Aplikace Princezna Pampeliska - you are logged in as: ' + uNm
        self.initUI()

    def initUI(self):
        self.setWindowTitle(self.title)
        
        mainMenu = self.menuBar()

        Purchase = mainMenu.addMenu('Purchases')
        Purchase.addAction('Open Purchase Template').triggered.connect(lambda arg, sql = 'Table_of_widgets': self.on_display(Window(sql)))
        Purchase.addAction('Cancel Realized Purchase').triggered.connect(self.on_CCB)
        subMenu = QMenu('History of Purchases', self)
        Purchase.addMenu(subMenu)

        subMenu.addAction('The Last Purchase(s)').triggered.connect(lambda arg, sql = 'LastPurchase': self.on_display(TBWindow(sql)))
        subMenu.addAction('Purchases Today').triggered.connect(lambda arg, sql = 'PurchasesToday': self.on_display(TBWindow(sql)))
        subMenu.addAction('Purchases This Week').triggered.connect(lambda arg, sql = 'PurchasesThisWeek': self.on_display(TBWindow(sql)))
        subMenu.addAction('Purchases This Month').triggered.connect(lambda arg, sql = 'PurchasesThisMonth': self.on_display(TBWindow(sql)))
        subMenu.addAction('Purchases Last Month').triggered.connect(lambda arg, sql = 'PurchasesLastMonth': self.on_display(TBWindow(sql)))
        subMenu.addAction('All Purchases').triggered.connect(lambda arg, sql = 'AllPurchases': self.on_display(TBWindow(sql)))

        Costs = mainMenu.addMenu('Costs')
        Costs.addAction('Open Cost Template').triggered.connect(lambda arg, sql = 'SaveCosts': self.on_display(Window(sql)))
        subM = QMenu('History of Costs', self)
        Costs.addMenu(subM)
        subM.addAction('The Last Cost(s)').triggered.connect(lambda arg, sql = 'LastCost': self.on_display(TBWindow(sql)))
        subM.addAction('Costs This Week').triggered.connect(lambda arg, sql = 'ThisWeekCosts': self.on_display(TBWindow(sql)))
        subM.addAction('Costs This Month').triggered.connect(lambda arg, sql = 'ThisMonthCosts': self.on_display(TBWindow(sql)))
        subM.addAction('Costs Last Month').triggered.connect(lambda arg, sql = 'LastMonthCosts': self.on_display(TBWindow(sql)))
        subM.addAction('All Costs').triggered.connect(lambda arg, sql = 'AllCosts': self.on_display(TBWindow(sql)))

        Warehouse = mainMenu.addMenu('Warehouse')
        Warehouse.addAction('Display Actual Stock').triggered.connect(lambda arg, sql = 'DisplayActStock': self.on_display(TBWindow(sql)))
        Warehouse.addAction('Modify Stock/Product').triggered.connect(lambda arg, sql = 'DisplayModStock': self.on_display(TBWindow(sql)))
        Warehouse.addAction('Add Product(s)').triggered.connect(lambda arg, sql = 'DisplayZero': self.on_display(TBWindow(sql)))
        Warehouse.addAction('Import Product(s)').triggered.connect(lambda arg , var = 'importProds': self.TCWarning(var))

        Analyses = mainMenu.addMenu('Overview')
        Analyses.addAction('Dynamic Overview of Sales').triggered.connect(lambda arg, glVar = None: self.on_display(Example(glVar)))
        Analyses.addAction('Dynamic Overview of Rents').triggered.connect(lambda arg, glVar = 'Rent': self.on_display(Example(glVar)))

        Reports = mainMenu.addMenu('Reports')
        Reports.addAction('Monthly Sales/Rents').triggered.connect(lambda arg , var = 'SReport': self.TCWarning(var))

##        Analyses.addAction('Total Sales/Label').triggered.connect(self.on_ViewAll)
##        Analyses.addAction('Sales Customized Overview').triggered.connect(self.on_selfAdj)
##        Analyses.addAction('Plotted Sales of Own Labels').triggered.connect(self.on_DP)
##        Analyses.addAction('Plotted Total Sales/Costs/Profit').triggered.connect(self.on_DTP)

        Tools = mainMenu.addMenu('Tools')
        Tools.addAction('Add New Seller').triggered.connect(lambda arg, glVar = None: self.on_display(Form(glVar)))
        Tools.addAction('Add New Rent Payer').triggered.connect(lambda arg, glVar = 'AddModPayer': self.on_display(Form(glVar)))
        Tools.addAction('Update Rent Payers').triggered.connect(lambda arg, sql = 'DisplayModPayers': self.on_display(TBWindow(sql)))
        Tools.addAction('Update Rents').triggered.connect(lambda arg, sql = 'DisplayModRents': self.on_display(TBWindow(sql)))
        Tools.addAction('Backup the DB').triggered.connect(lambda arg, bType = 'Regular': self.on_BackUp(bType))
        Tools.addAction('Backup the DB Online').triggered.connect(lambda arg, bType = 'On-line': self.on_BackUp(bType))
        Tools.addAction('Add New User').triggered.connect(lambda arg, sql = 'Add': self.on_display(UserForm(sql)))
        Tools.addAction('Remove Existing User').triggered.connect(lambda arg, sql = 'Remove': self.on_display(UserForm(sql)))
        Tools.addAction('Change User Password').triggered.connect(lambda arg, sql = 'Change': self.on_display(UserForm(sql)))
        Tools.addAction('Update Application').triggered.connect(lambda arg , var = 'update': self.TCWarning(var))
##        NewOrder = mainMenu.addMenu('Orders')

        sizeObject = QDesktopWidget().screenGeometry()
        w, h = sizeObject.width(), sizeObject.height()
        oImage = QImage(inputs + '\\logo.jpg')
        sImage = oImage.scaled(QSize(w, h))
        palette = QPalette()
        palette.setBrush(10, QBrush(sImage))
        self.setPalette(palette)

    def updateApp(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to update this application?' , QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            try:
                shutil.copyfile(locked + 'source.pyw', backup + 'source' + datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + '.pyw')
                req = requests.get('https://raw.githubusercontent.com/JiriCh/PP/master/source.pyw')
                with open(locked + 'source.pyw', 'w+', encoding='utf-8') as r:
                    r.write(req.text)
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The application has been updated successfully, please restart it to use its new functionality.', QMessageBox.Ok)
                msgbox.exec()
            except Exception as e:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The application update failed, please see the details below: %s' %e, QMessageBox.Ok)
                msgbox.exec()

    def TCWarning(self, var):
        if uNm in ('Admin', 'Jana') and var == 'SReport':
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'This action might take few minutes, do you wish to proceed?' , QMessageBox.Yes|QMessageBox.No)
            res = msgbox.exec()
            if res == QMessageBox.Yes:
                self.view = SRreport()
                self.view.show()
        elif uNm in ('Admin', 'Jana') and var == 'importProds':
            self.importProds()
        elif uNm in ('Admin', 'Jana') and var == 'update':
            self.updateApp()
        else:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'This user does not have authorization for that action.', QMessageBox.Ok)
            msgbox.exec()

    def on_CCB(self):
        text, okPressed = QInputDialog.getText(self, "Purchase Cancelation", "Enter identifier of the purchase you would like to cancel:", QLineEdit.Normal, "")
        sql_query(sql_dict['CancelCheck']['sql'], sql_dict['CancelCheck']['args'])
        
        if (text,) in sql_obj and okPressed:
            sql_query(sql_dict['CancelReq']['sql'], [text])
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Selected purchase has been sucessfully canceled.', QMessageBox.Ok)
            msgbox.exec()

        if (text,) not in sql_obj and okPressed:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Selected purchase has never been realized, please re-check the identifier of the purchase and re-enter it.', QMessageBox.Ok)
            msgbox.exec()
    
    def on_display(self, arg):
        if arg.__class__.__name__ == 'Example' or arg.__class__.__name__ == 'Form':
            if uNm in ('Admin', 'Jana'):
                self.view = arg
                self.view.show()
            else:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'This user does not have authorization for that action.', QMessageBox.Ok)
                msgbox.exec()
        else:
            if arg.sql in ('ThisMonthCosts', 'ThisWeekCosts', 'LastCost', 'AllCosts', 'LastMonthCosts', 'SaveCosts', 'DisplayModStock', 'DisplayModPayers', 'DisplayModRents', 'DisplayZero') and uNm not in ('Admin', 'Jana'):
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'This user does not have authorization for that action.', QMessageBox.Ok)
                msgbox.exec()
            else:
                if arg.__class__.__name__ == 'UserForm':
                    self.view = arg
                    self.view.show()
                else:
                    self.view = arg
                    self.view.showMaximized()

    def on_BackUp(self, bType):
        if bType == 'Regular':
            try:
                backupdir = backup
                dbfile = locked + '\\DataBase.db'
                backup_file = os.path.join(backupdir, datetime.now().strftime('%Y-%m-%d_%H-%M-%S') + os.path.basename(dbfile))
                shutil.copyfile(dbfile, backup_file)
                
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The DataBase has been successfully backed-up.', QMessageBox.Ok)
                msgbox.exec()
            except:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The DataBase has not been backed-up, some problem occurred.', QMessageBox.Ok)
                msgbox.exec()
        if bType == 'On-line':
            try:
                send_to = 'info@princezna-pampeliska.cz'
                subject = 'Automatically generated email - DataBase copy'
                message = 'This message was automatically generated and sent by Princezna Pampeliska application. Attached might be found the DB copy.'
                files = [locked + 'DataBase.db']
                server = 'smtp.gmail.com'
                port = 587
                username = 'jana.klapetkova@gmail.com'#'tmp.testem@gmail.com'
                password = 'Bpv15v5blc'#'ocepo123'

                msg = MIMEMultipart()
                msg['From'] = username
                msg['To'] = send_to
                msg['Date'] = formatdate(localtime=True)
                msg['Subject'] = subject

                msg.attach(MIMEText(message))

                for p in files:
                    part = MIMEBase('application', "octet-stream")
                    with open(p, 'rb') as file:
                        part.set_payload(file.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition',
                                    'attachment; filename="{}"'.format(os.path.basename(p)))
                    msg.attach(part)

                smtp = smtplib.SMTP(server, port)
                smtp.ehlo()
                smtp.starttls()
                smtp.login(username, password)
                smtp.sendmail(username, send_to, msg.as_string())
                smtp.quit()

                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The DataBase has been sucessfully backed-up.', QMessageBox.Ok)
                msgbox.exec()

            except:      
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'An error occurred, the DataBase has not been backed-up.', QMessageBox.Ok)
                msgbox.exec()

    def importProds(self):
        fname, fEnd = QFileDialog.getOpenFileName(self, "Save File", desktop, "Data Files(*.xls; *.xlsx; *.csv)")
        if fname:
            if 'xlsx' in fname or 'xls' in fname:
                try:
                    res = read_excel(fname)
                    res.columns = 'Category', 'Product_id', 'Product_desc', 'Label', 'Unit_price', 'Stock', 'Added_by', 'Usage_flag'
                    sql = res.values.astype(str).tolist()
                    if all(True if len(x) == 8 and 'nan' not in x else False for x in sql):
                        self.w = TBWindow(sql)
                        self.w.showMaximized()
                    else:
                        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The data you would like to insert into the Database has incorrect dimension, it needs to have 8 columns.', QMessageBox.Ok)
                        msgbox.exec()
                except Exception as e:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'An unexpected error occurred, this action cannot be completed. The error details are: %s.' %e, QMessageBox.Ok)
                    msgbox.exec()
            elif 'csv' in fname:
                try:
                    res = read_csv(fname)
                    res.columns = 'Category', 'Product_id', 'Product_desc', 'Label', 'Unit_price', 'Stock', 'Added_by', 'Usage_flag'
                    sql = res.values.astype(str).tolist()
                    if all(True if len(x) == 8 and 'nan' not in x else False for x in sql):
                        self.w = TBWindow(sql)
                        self.w.showMaximized()
                    else:
                        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The data you would like to insert into the Database has incorrect dimension, it needs to have 8 columns.', QMessageBox.Ok)
                        msgbox.exec()
                except Exception as e:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'An unexpected error occurred, this action cannot be completed. The error details are: %s.' %e, QMessageBox.Ok)
                    msgbox.exec()

class Window(QMainWindow):
    def __init__(self, sql, parent = None):
        super(Window, self).__init__(parent)

        self.setWindowTitle('Aplikace Princezna Pampeliska')
        menu = self.menuBar()

        self.sql = sql

        if self.sql == 'Table_of_widgets':
            sql_query(sql_dict[sql]['sql'], sql_dict[sql]['args'])
            menu.addAction('Save and Close').triggered.connect(self.on_SCPT)
            menu.addAction('Cancel and Close').triggered.connect(self.on_CCPT)

            paymentType = QMenu('Payment Type', self)
            self.group = QActionGroup(paymentType)
            texts = ["Cash", "Noncash Payment", "Cash on Delivery", "Bank Transfer"]
            for text in texts:
                action = QAction(text, paymentType, checkable=True, checked=text==texts[0])
                paymentType.addAction(action)
                self.group.addAction(action)
            self.group.setExclusive(True)
            menu.addMenu(paymentType)
            menu.addAction('Display Total Value').triggered.connect(self.on_DTV)
            
            self.b = QCheckBox('Create Receipt    ',self)
            self.nameL = QCheckBox('Postage:')
            self.EditL = QLineEdit(self)
            self.EET = QCheckBox('EET    ')
            self.EditL.setFixedWidth(100)
            onlyInteger = QIntValidator()
            self.EditL.setValidator(onlyInteger)
            self.addRow = QPushButton('Add Row')
            self.addRow.clicked.connect(self.addR)
            self.spacer = QWidget()
            self.spacer.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            toolBar = self.addToolBar('oneAct')
            toolBar.addWidget(self.EET)
            toolBar.addWidget(self.b)
            toolBar.addWidget(self.nameL)
            toolBar.addWidget(self.EditL)
            toolBar.addWidget(self.spacer)
            toolBar.addWidget(self.addRow)

            self.table = Table_of_widgets(self)
            self.setCentralWidget(self.table)
        else:
            menu.addAction('Save and Close').triggered.connect(self.on_SaveCosts)
            menu.addAction('Cancel and Close').triggered.connect(self.onCancelReq)
            menu.addAction('Display Total Amount Paid').triggered.connect(self.on_sumCheck)
            
            self.table = ToWspec(self)
            self.setCentralWidget(self.table)

    def addR(self):
        self.table.addRow()

    def on_SCPT(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to save the purchase template and close it?' , QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            source = self.table.slot()
            PaymentType = self.group.checkedAction().text()
            if len(source) > 0:
                if any([(source.index(x),x.index(y)) for x in source for y in x if (y == ' ' or y == '')]) or (self.nameL.isChecked() and self.EditL.text() == '') or (not(self.nameL.isChecked()) and self.EditL.text() != ''):
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The purchase template is incomplete, please check it.', QMessageBox.Ok)
                    msgbox.exec()
                else:
                    if  self.nameL.isChecked():
                        postage = [source[0][0], '1', 'post01', self.EditL.text(), self.EditL.text()]
                        source.append(postage)
                    for index, element in enumerate(source):
                        sysTime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        letter = chr(index+97)
                        identifier = sysTime.replace(' ',letter)
                        stornoFlag = 'N'
                        element.extend([sysTime,identifier,stornoFlag,PaymentType])
                        
                    checkList = []
                    for row in source:
                        if row[1].isdigit() and row[3].isdigit() and row[4].isdigit():
                            checkList.append(1)

                    if checkList == []:
                        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'At least one input in Unit price and/or Total price field is not of a numeric format.', QMessageBox.Ok)
                        msgbox.exec()
                    else:
                        #this needs to be implemented
                        if self.EET.isChecked():
                            print('self.EET.isChecked()', source)
                            self.close()
                        else:
                            for element in source:
                                sql_query(sql_dict['SavePurchase']['sql'], (element))

                            if self.b.isChecked():

                                DocTemplate = inputs + '\\PurchaseTemplate.docx'

                                lines = []
                                pSum = []
                                reOrder = [1,0,2,3]

                                getID = []
                                for row in source:
                                    getID.append(row[2])

                                addInfo = []
                                for item in getID:
                                    sql_query(sql_dict['addInfo']['sql'], [item])
                                    addInfo.append(list(*sql_obj))

                                temp = []
                                for index, element in enumerate(source):
                                    fin = addInfo[index] + source[index]
                                    temp.append(fin)
                                
                                for i in temp:
                                    eList = []
                                    ppSum = []
                                    for index, element in enumerate(i):
                                        if index in (2,4,6,7):
                                            if index in (6,7):
                                                eList.append(format(int(element),'.2f'))
                                            else:
                                                eList.append(element)
                                        if index == 7:
                                            ppSum.append(int(element))
                                    fList = [eList[j] for j in reOrder]
                                    lines.append(fList)
                                    pSum.append(ppSum)

                                Total = int(str([sum(i) for i in zip(*pSum)]).strip('[]'))
                                DPHperc = 0.21
                                DPH = DPHperc * Total
                                woDPH = Total - DPH
                                            
                                keys = ['NPieces', 'ProdDesc', 'UnitPrice', 'TotalPrice']
                                dictList = []

                                for i in lines:
                                    dictList.append(dict(zip(keys, i)))

                                paymentDD = {'Cash' : 'Hotově' , 'Noncash Payment' : 'Bezhotovostně', 'Cash on Delivery' : 'Na dobírku', 'Bank Transfer' : 'Bankovním převodem'}

                                document = MailMerge(DocTemplate)
                                document.merge(
                                    BusinessName = 'Princezna Pampeliška - Jana Klapetková',
                                    PurchaseID = 'PN' + str(round(datetime.today().timestamp())),
                                    CashID = '1',
                                    Cashier = uNm,
                                    paymentMethod = paymentDD.get(self.group.checkedAction().text()),
                                    iDPH = str(format(DPH,'.2f')),
                                    iwoDPH = str(format(woDPH,'.2f')),
                                    iDPHperc = str(format(round(DPHperc * 100),'.2f')),
                                    iTotal = str(format(Total,'.2f')),
                                    TimeStamp = '{:%d-%b-%Y}'.format(date.today()))
                                document.merge_rows('NPieces', dictList)
                                document.write(desktop + 'PN' + str(round(datetime.today().timestamp())) + '.docx')

                            self.close()               

            else:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Nothing has been selected within the purchase template.', QMessageBox.Ok)
                msgbox.exec()

    def on_CCPT(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to cancel this purchase and close the template?', QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            self.close()

    def on_SaveCosts(self):
        source = self.table.slot()
        if len(source) > 0:

            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to save the cost(s) and close this template?', QMessageBox.Yes|QMessageBox.No)
            res = msgbox.exec()
            if res == QMessageBox.Yes:
                if any([(source.index(x),x.index(y)) for x in source for y in x if (y == ' ' or y == '' or y == '0')]):
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The purchase template is incomplete, please check it.', QMessageBox.Ok)
                    msgbox.exec()
                else:
                    for index, element in enumerate(source):
                        sysTime = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                        vType = 'Costs'
                        Outstanding = 'N'
                        letter = chr(index+97)
                        identifier = sysTime.replace(' ',letter)
                        element.extend([sysTime,identifier,vType])
                        element.append(cost_dict[element[0]])
                        element.append(Outstanding)
                        sql_query(sql_dict[self.sql]['sql'], (element))
                        
                    self.close()                    
        else:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The template is empty, total amount paid is 0 Kc.', QMessageBox.Ok)
            msgbox.exec()

    def onCancelReq(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to cancel the action and close this template?', QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            self.close()

    def on_sumCheck(self):
        source = self.table.slot()
        if len(source) > 0:
            if any([(source.index(x),x.index(y)) for x in source for y in x if (y == ' ' or y == '' or y == '0')]):
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The purchase template is incomplete, please check it.', QMessageBox.Ok)
                msgbox.exec()
            else:
                rfSource = DataFrame(source)
                col2SumUp = rfSource[4].apply(int)
                purchaseValue = col2SumUp.sum()
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Total value of this purchase is: %s Kc.' % purchaseValue, QMessageBox.Ok)
                msgbox.exec()
        else:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The template is empty, total amount paid is 0 Kc.', QMessageBox.Ok)
            msgbox.exec()

    def on_DTV(self):
        source = self.table.slot()
        if len(source) > 0:
            if any([(source.index(x),x.index(y)) for x in source for y in x if (y == ' ' or y == '')]):
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The purchase template is incomplete, please check it.', QMessageBox.Ok)
                msgbox.exec()
            else:
                checkList = []
                for row in source:
                    if row[1].isdigit() and row[3].isdigit() and row[4].isdigit():
                        checkList.append(1)
                if checkList == []:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'At least one input in Unit price and/or Total price field is not of a numeric format.', QMessageBox.Ok)
                    msgbox.exec()
                else:
                    rfSource = DataFrame(source)
                    col2SumUp = rfSource[4].apply(int)
                    purchaseValue = col2SumUp.sum()
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Total value of this purchase is: %s Kc.' % purchaseValue, QMessageBox.Ok)
                    msgbox.exec()
        else:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The template is empty, total value of this purchase is 0 Kc.', QMessageBox.Ok)
            msgbox.exec()
                
class Table_of_widgets(QTableWidget):
    def __init__(self, parent = None):
        super(Table_of_widgets, self).__init__(parent)
        
        self.rC = 1
        self.cC = 8

        global df
        df = DataFrame(sql_obj)
        fin = {}
        for i in df:
            fin[i] = df[i]
            fin[i] = df[i].drop_duplicates()
            fin[i] = list(fin[i])
            fin[i].insert(0,' ')

        self.setColumnCount(self.cC)
        self.setRowCount(self.rC)
        self.setHorizontalHeaderLabels(['Section',
                                        'Label',
                                        'Product description',
                                        'Store',
                                        'Amount',
                                        'Product ID',
                                        'Unit price',
                                        'Total price'])
        self.horizontalHeader().setFixedHeight(80)
        self.verticalHeader().hide()

        shops = [' ', 'Holesovice', 'Fler', 'E-shop', 'Sashe', 'Facebook', 'Actions', 'Other']

        colW0 = 6.5*len(max(fin[0], key=len))
        colW1 = 7*len(max(fin[1], key=len))
        colW2 = 6*len(max(fin[2], key=len))
        colW3 = 8.5*len(max(shops, key=len))
        colW4 = 7.5*len(max(fin[3], key=len))

        self.setColumnWidth(0,colW0)
        self.setColumnWidth(1,colW1)
        self.setColumnWidth(2,colW2)
        self.setColumnWidth(3,colW3)
        self.setColumnWidth(4,55)
        self.setColumnWidth(5,colW4)

        for i in (6,7):
            self.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
            
        self.setStyleSheet( "QHeaderView::section{"
                            "border-top:1px solid #D8D8D8;"
                            "border-left:1px solid #D8D8D8;"
                            "border-right:2px solid #D8D8D8;"
                            "border-bottom: 3px solid #D8D8D8;"
                            "background-color:white;"
                            "padding:4px;""}")
                
        self.offer1 = fin[0]
        self.offer2 = fin[1]
        self.offer3 = fin[2]
        self.offer4 = shops
        self.offer5 = fin[3]
        
        for i in range(self.rC):

            self.comboA = myCombo()
            self.comboB = myCombo()
            self.comboC = myCombo()
            self.comboD = myCombo()
            self.comboE = myCombo()
            self.spinBox = mySpinbox()

            self.comboA.addItems(sorted(self.offer1))
            self.comboB.addItems(sorted(self.offer2))
            self.comboC.addItems(sorted(self.offer3))
            self.comboD.addItems(self.offer4)
            self.comboE.addItems(sorted(self.offer5))

            self.setCellWidget(i, 0, self.comboA)
            self.setCellWidget(i, 1, self.comboB)
            self.setCellWidget(i, 2, self.comboC)
            self.setCellWidget(i, 3, self.comboD)
            self.setCellWidget(i, 4, self.spinBox)
            self.setCellWidget(i, 5, self.comboE)
            self.setItem(i, 6, QTableWidgetItem(' '))
            self.setItem(i, 7, QTableWidgetItem(' '))

            self.comboA.currentTextChanged.connect(lambda text, row=i: self.onComboACurrentTextChanged(text, row))
            self.comboB.currentTextChanged.connect(lambda text, row=i: self.onComboBCurrentTextChanged(text, row))
            self.comboC.currentTextChanged.connect(lambda text, row=i: self.onComboCCurrentTextChanged(text, row))
            self.comboD.currentTextChanged.connect(lambda text, row=i: self.onComboDCurrentTextChanged(text, row))
            self.comboE.currentTextChanged.connect(lambda text, row=i: self.onComboECurrentTextChanged(text, row))
            self.spinBox.valueChanged.connect(lambda text, row=i: self.onSpinBoxChanged(text, row))

    def addRow(self):
        for i in range(self.rowCount(),self.rowCount()+1):
            self.insertRow(self.rowCount())
        
            self.comboA = myCombo()
            self.comboB = myCombo()
            self.comboC = myCombo()
            self.comboD = myCombo()
            self.comboE = myCombo()
            self.spinBox = mySpinbox()

            self.comboA.addItems(sorted(self.offer1))
            self.comboB.addItems(sorted(self.offer2))
            self.comboC.addItems(sorted(self.offer3))
            self.comboD.addItems(self.offer4)
            self.comboE.addItems(sorted(self.offer5))

            self.setCellWidget(i, 0, self.comboA)
            self.setCellWidget(i, 1, self.comboB)
            self.setCellWidget(i, 2, self.comboC)
            self.setCellWidget(i, 3, self.comboD)
            self.setCellWidget(i, 4, self.spinBox)
            self.setCellWidget(i, 5, self.comboE)
            self.setItem(i, 6, QTableWidgetItem(' '))
            self.setItem(i, 7, QTableWidgetItem(' '))

            self.comboA.currentTextChanged.connect(lambda text, row=i: self.onComboACurrentTextChanged(text, row))
            self.comboB.currentTextChanged.connect(lambda text, row=i: self.onComboBCurrentTextChanged(text, row))
            self.comboC.currentTextChanged.connect(lambda text, row=i: self.onComboCCurrentTextChanged(text, row))
            self.comboD.currentTextChanged.connect(lambda text, row=i: self.onComboDCurrentTextChanged(text, row))
            self.comboE.currentTextChanged.connect(lambda text, row=i: self.onComboECurrentTextChanged(text, row))
            self.spinBox.valueChanged.connect(lambda text, row=i: self.onSpinBoxChanged(text, row))

    def slot(self):
        myList = []
        for i in range(self.rowCount()):
            output = []
            for j in range(self.columnCount()):
                if j == 3 or j == 5:
                    w = self.cellWidget(i,j)
                    a = w.currentText()
                    output.append(a)
                if j == 4:
                    w = self.cellWidget(i,j)
                    a = str(w.value())
                    output.append(a)
                if j > 5:
                    w = self.item(i,j)
                    a = w.text()
                    output.append(a)
            if output[0] != ' ' or output[1] != '0' or output[3] != ' ' or output[4] != ' ':
                myList.append(output)
        return myList
                            
    def updateCombox(self, row, combo1, combo2, combo3, combo4, combo5, SpinBox, offer1, offer2, offer3, offer5):

        font = QFont("Times", 8)
        
        text1 = combo1.currentText()
        text2 = combo2.currentText()
        text3 = combo3.currentText()
        text4 = combo4.currentText()
        text5 = combo5.currentText()
        value1 = SpinBox.value()
        combo1.blockSignals(True)
        combo2.blockSignals(True)
        combo3.blockSignals(True)
        combo5.blockSignals(True)
        combo1.clear()
        combo2.clear()
        combo3.clear()
        combo5.clear()

        if text1 == ' ': a = list(df[0].drop_duplicates())
        else: a = [text1]
        if text2 == ' ': b = list(df[1].drop_duplicates())
        else: b = [text2]
        if text3 == ' ': c = list(df[2].drop_duplicates())
        else: c = [text3]
        if text5 == ' ': d = list(df[3].drop_duplicates())
        else: d = [text5]

        offer1 = list(df.loc[df[0].isin(a) & df[1].isin(b) & df[2].isin(c) & df[3].isin(d)][0].drop_duplicates())
        offer1.insert(0, ' ')
        offer2 = list(df.loc[df[0].isin(a) & df[1].isin(b) & df[2].isin(c) & df[3].isin(d)][1].drop_duplicates())
        offer2.insert(0, ' ')
        offer3 = list(df.loc[df[0].isin(a) & df[1].isin(b) & df[2].isin(c) & df[3].isin(d)][2].drop_duplicates())
        offer3.insert(0, ' ')
        offer5 = list(df.loc[df[0].isin(a) & df[1].isin(b) & df[2].isin(c) & df[3].isin(d)][3].drop_duplicates())
        offer5.insert(0, ' ')

        combo1.addItems(sorted(offer1))
        combo1.setCurrentText(text1)
        combo2.addItems(sorted(offer2))
        combo2.setCurrentText(text2)
        combo3.addItems(sorted(offer3))
        combo3.setCurrentText(text3)
        combo5.addItems(sorted(offer5))
        combo5.setCurrentText(text5)

        zero = ' '

        if text1 != ' ' and text2 != ' ' and text3 != ' ' and text4 != ' ' and text5!= ' ':
            unit_price = list(df.loc[df[0].isin([text1]) & df[1].isin([text2]) & df[2].isin([text3]) & df[3].isin([text5])][4])[0]
            self.setItem(row,6,QTableWidgetItem(unit_price))
            self.item(row,6).setTextAlignment(Qt.AlignCenter)
            self.item(row,6).setFont(font)
        else:
            self.setItem(row,6,QTableWidgetItem(zero))
            self.item(row,6).setFont(font)

        if text1 != ' ' and text2 != ' ' and text3 != ' ' and text4 != ' ' and text5 != ' ' and value1 != 0:
            unit_price = list(df.loc[df[0].isin([text1]) & df[1].isin([text2]) & df[2].isin([text3]) & df[3].isin([text5])][4])[0]
            self.setItem(row,7,QTableWidgetItem(str(int(value1)*int(unit_price))))
            self.item(row,7).setTextAlignment(Qt.AlignCenter)
            self.item(row,7).setFont(font)
        else:
            self.setItem(row,7,QTableWidgetItem(zero))
            self.item(row,7).setFont(font)
        
        combo1.blockSignals(False)
        combo2.blockSignals(False)
        combo3.blockSignals(False)
        combo5.blockSignals(False)

    def onComboACurrentTextChanged(self, text, row):
        comboA = self.cellWidget(row, 0)
        comboB = self.cellWidget(row, 1)
        comboC = self.cellWidget(row, 2)
        comboD = self.cellWidget(row, 3)
        comboE = self.cellWidget(row, 5)
        SpinBox = self.cellWidget(row, 4)
        self.updateCombox(row, comboA, comboB, comboC, comboD, comboE, SpinBox, self.offer1, self.offer2, self.offer3, self.offer5)

    def onComboBCurrentTextChanged(self, text, row):
        comboA = self.cellWidget(row, 0)
        comboB = self.cellWidget(row, 1)
        comboC = self.cellWidget(row, 2)
        comboD = self.cellWidget(row, 3)
        comboE = self.cellWidget(row, 5)
        SpinBox = self.cellWidget(row, 4)
        self.updateCombox(row, comboA, comboB, comboC, comboD, comboE, SpinBox, self.offer1, self.offer2, self.offer3, self.offer5)

    def onComboCCurrentTextChanged(self, text, row):
        comboA = self.cellWidget(row, 0)
        comboB = self.cellWidget(row, 1)
        comboC = self.cellWidget(row, 2)
        comboD = self.cellWidget(row, 3)
        comboE = self.cellWidget(row, 5)
        SpinBox = self.cellWidget(row, 4)
        self.updateCombox(row, comboA, comboB, comboC, comboD, comboE, SpinBox, self.offer1, self.offer2, self.offer3, self.offer5)

    def onComboDCurrentTextChanged(self, text, row):
        comboA = self.cellWidget(row, 0)
        comboB = self.cellWidget(row, 1)
        comboC = self.cellWidget(row, 2)
        comboD = self.cellWidget(row, 3)
        comboE = self.cellWidget(row, 5)
        SpinBox = self.cellWidget(row, 4)
        self.updateCombox(row, comboA, comboB, comboC, comboD, comboE, SpinBox, self.offer1, self.offer2, self.offer3, self.offer5)

    def onComboECurrentTextChanged(self, text, row):
        comboA = self.cellWidget(row, 0)
        comboB = self.cellWidget(row, 1)
        comboC = self.cellWidget(row, 2)
        comboD = self.cellWidget(row, 3)
        comboE = self.cellWidget(row, 5)
        SpinBox = self.cellWidget(row, 4)
        self.updateCombox(row, comboA, comboB, comboC, comboD, comboE, SpinBox, self.offer1, self.offer2, self.offer3, self.offer5)

    def onSpinBoxChanged(self, text, row):
        comboA = self.cellWidget(row, 0)
        comboB = self.cellWidget(row, 1)
        comboC = self.cellWidget(row, 2)
        comboD = self.cellWidget(row, 3)
        comboE = self.cellWidget(row, 5)
        SpinBox = self.cellWidget(row, 4)
        self.updateCombox(row, comboA, comboB, comboC, comboD, comboE, SpinBox, self.offer1, self.offer2, self.offer3, self.offer5)

class TBWindow(QMainWindow):
    def __init__(self, sql, parent=None):
        super(TBWindow, self).__init__(parent)

        self.sql = sql

        if isinstance(self.sql, str):
            sql_query(sql_dict[sql]['sql'], sql_dict[sql]['args'])
            self.originalData = sql_obj
            data = DataFrame(sql_obj, columns = headers)
        elif isinstance(self.sql, list):
            sql_query(sql_dict['DisplayRes']['sql'], None)
            first5 = [list(x) for x in sql_obj]
            first5.append(['...', '...', '...', '...', '...', '...', '...', '...'])
            res = first5 + self.sql
            data = DataFrame(res, columns = headers)
        else:
            self.sql = None
            data = sql
            
        self.setWindowTitle('Aplikace Princezna Pampeliska')
        self.centralwidget  = QWidget(self)
        self.lineEdit       = QLineEdit(self.centralwidget)
        self.view           = QTableView(self.centralwidget)
        self.comboBox       = QComboBox(self.centralwidget)
        self.label          = QLabel(self.centralwidget)
        
        self.gridLayout = QGridLayout(self.centralwidget)

        self.label.setText("Filter")
        self.gridLayout.addWidget(self.comboBox, 0, 2, 1, 1)
        self.gridLayout.addWidget(self.label, 0, 0, 1, 1)
        self.gridLayout.addWidget(self.lineEdit, 0, 1, 1, 1)

        self.gridLayout.addWidget(self.view, 1, 0, 1, 3)
        self.setCentralWidget(self.centralwidget)

        if self.sql in ('DisplayModStock', 'DisplayZero', 'DisplayModPayers', 'DisplayModRents'):
            if len(data.index) == 0:#create empty table for new products
                emptyList = []
                for i in range(20):
                    emptyList.append(' ')
                    
                data = DataFrame({
                                'Category' : emptyList,
                                'Product_id' : emptyList,
                                'Product_desc' : emptyList,
                                'Label' : emptyList,
                                'Unit_price' : emptyList,
                                'Stock' : emptyList,
                                'Added_by' : emptyList,
                                'Usage_flag' : emptyList},
                                index=[0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19])
                data = data[headers]

            MenuBar = self.menuBar()
            saveChange = QAction("&Save Changes", self)
            saveChange.setShortcut("Ctrl+S")
            saveChange.setStatusTip('Save Changes')
            saveChange.triggered.connect(self.change_save)
            MenuBar.addAction(saveChange)
            MenuBar.addAction('Cancel Changes').triggered.connect(self.on_CancelModification)

            self.model = ModifiedModel(data)
            self.proxy = QSortFilterProxyModel(self)
            self.proxy.setSourceModel(self.model)
            self.view.setModel(self.proxy)
            
            for column in range(self.view.horizontalHeader().count()):
                self.view.resizeColumnToContents(column) 
                if column != 2:
                    self.view.horizontalHeader().setSectionResizeMode(column, QHeaderView.Stretch)
                else:
                    self.view.horizontalHeader().setDefaultSectionSize(400)
        else:
            MenuBar = self.menuBar()
            if isinstance(self.sql, list):
                MenuBar.addAction('Import Products').triggered.connect(self.importProducts)
                self.check = [x[1] for x in self.sql]
                sql_query(sql_dict['CheckStock']['sql'] %','.join('?'*len(self.check)), self.check)
                self.checkRes = sql_obj
                dups = [index + 6 for index, row in enumerate(self.sql) if tuple(row) in self.checkRes]
                row = 5
            else:
                MenuBar.addAction('Save File').triggered.connect(self.file_save)
                row = None
                dups = None
            MenuBar.addAction('Close Window').triggered.connect(self.on_Cancel)
            self.model = PandasModel(data, row, dups)
            self.proxy = QSortFilterProxyModel(self)
            self.proxy.setSourceModel(self.model)
            self.view.setModel(self.proxy)
            if self.sql == 'DisplayActStock' or isinstance(self.sql, list):
                for column in range(self.view.horizontalHeader().count()):
                    self.view.resizeColumnToContents(column) 
                    if column != 2:
                        self.view.horizontalHeader().setSectionResizeMode(column, QHeaderView.Stretch)
                    else:
                        self.view.horizontalHeader().setDefaultSectionSize(400)
            elif self.sql is None:
                for column in range(self.view.horizontalHeader().count()):
                    self.view.horizontalHeader().setSectionResizeMode(column, QHeaderView.Stretch)
            else:
                for column in range(self.view.horizontalHeader().count()):
                    self.view.resizeColumnToContents(column)
                    if column == self.view.horizontalHeader().count() - 1:
                        self.view.horizontalHeader().setSectionResizeMode(column, QHeaderView.Stretch)
            
        self.comboBox.addItems(list(data.columns.values))
        self.lineEdit.textChanged.connect(self.on_lineEdit_textChanged)
        self.comboBox.currentIndexChanged.connect(self.on_comboBox_currentIndexChanged)
        
        if self.sql in ('DisplayModPayers', 'DisplayModRents', 'DisplayZero'):
            self.comboBox.hide()
            self.label.hide()
            self.lineEdit.hide()

        elif isinstance(self.sql, list):
            self.comboBox.hide()
            self.label.hide()
            self.lineEdit.hide()

        elif self.sql is None:
            lastRow = data.append(data.sum(numeric_only=True), ignore_index=True).tail(1).fillna(0).astype(int).replace(0,'').values
            flat_Row = ['Total: ' + str(item) if item != '' else '' for sublist in lastRow for item in sublist]

            self.mWdg = QWidget()
            self.mWdg.setLayout(QHBoxLayout())
            for item in flat_Row:
                self.mWdg.layout().addWidget(QLabel(str(item)))
            self.statusBar().addPermanentWidget(self.mWdg, 1)

    def importProducts(self):
        if self.checkRes == []:
            for element in self.sql:
                sql_query(sql_dict['RegUpProd']['sql'], (element))
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The import of product(s) completed successfully.', QMessageBox.Ok)
            msgbox.exec()
            self.close()
        else:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The item(s) you want to insert into Database should be already there: %s.' %self.checkRes, QMessageBox.Ok)
            msgbox.exec()

    def on_CancelModification(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to proceed without saving the change(s)?', QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            self.close()

    def change_save(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to save the change(s)?', QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            newModel = self.view.model()
            data = []
            for row in range(newModel.rowCount()):
                rowRes = []
                for column in range(newModel.columnCount()):
                    index = newModel.index(row, column)
                    item = newModel.data(index)
                    rowRes.append(item)
                data.append(rowRes)
            cleanData = []
            #this needs to be finished: elif self.sql == ...
            if self.sql == 'DisplayModPayers':
                sql_query(sql_dict['DeleteModPayers']['sql'], (None,))
                for row in data:
                    if row[4] == 'None':
                        row[4] = None
                    if row[5] == 'None':
                        row[5] = None
                    sql_query(sql_dict['ChangeModPayers']['sql'], row)
                self.close()
            elif self.sql == 'DisplayModRents':
                check = [list(tup) for tup in sql_obj]
                iCheck = [(tup[2], tup[3]) for tup in sql_obj]
                res = []
                for row in data:
                    if row not in check:
                        res.append(row)
                InjectionCheck = [(x[2],x[3]) for x in res]
                fcheck = []
                for i in InjectionCheck:
                    if i in iCheck:
                        fcheck.append(i)
                if len(fcheck) == len(InjectionCheck) and len(fcheck) != 0:
                    if all(x[6] in ('Y', 'N') for x in res):
                        for row in res:
                            row.append(row[2])
                            row.append(row[3])
                            sql_query(sql_dict['UpdateRents']['sql'], row)
                        self.close()
                    else:
                        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The Outstanding column can only contain either N or Y values.', QMessageBox.Ok)
                        msgbox.exec()
                else:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'You cannot over-write Seller and/or Event_time columns.', QMessageBox.Ok)
                    msgbox.exec()
            else:
                for row in data:
                    if any([x for x in row if (x not in ' ' and x not in '')]):
                        cleanData.append(row)

                if len(cleanData) == 0 or any([(cleanData.index(x),x.index(y)) for x in cleanData for y in x if (y == ' ' or y == '')]):
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The template is incomplete, please check it.', QMessageBox.Ok)
                    msgbox.exec()
                else:
                    dataFrame = DataFrame(cleanData)
                    dataFrame[6] = uNm
                                    
                    if all(dataFrame[4].str.isnumeric()) and all(dataFrame[5].str.isnumeric()):
                        sql_query(sql_dict['Dups_check']['sql'], sql_dict['Dups_check']['args'])
                        df_check = DataFrame(sql_obj)
                        
                        if any(dataFrame[1].isin(df_check[3])):
                            dataFrame = dataFrame.values.tolist()
                            for element in dataFrame:
                                element.append(element[1])
                                sql_query(sql_dict['UpdateProd']['sql'], (element))
                            self.close()
                        else:
                            dataFrame = dataFrame.values.tolist()
                            for element in dataFrame:
                                sql_query(sql_dict['RegUpProd']['sql'], (element))
                            self.close()
                    else:
                        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'At least one input in Unit_price and/or Stock field is not of a numeric format.' , QMessageBox.Ok)
                        res = msgbox.exec()

    def on_Cancel(self):
        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Do you really want to close this window?' , QMessageBox.Yes|QMessageBox.No)
        res = msgbox.exec()
        if res == QMessageBox.Yes:
            self.close()

    def file_save(self):
        newModel = self.view.model()
        data = []
        for row in range(newModel.rowCount()):
            rowRes = []
            for column in range(newModel.columnCount()):
                index = newModel.index(row, column)
                item = newModel.data(index)
                if item != '':
                    rowRes.append(item)
            data.append(rowRes)
        dataFrame = DataFrame(data)
            
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        fileName, fEnd = QFileDialog.getSaveFileName(self, "Save File", "", ".csv")
        if fileName:
            address = fileName+fEnd
            dataFrame.to_csv(address, index=False, header=False, encoding='utf-16')

    @pyqtSlot(str)
    def on_lineEdit_textChanged(self, text):
        search = QRegExp(text, Qt.CaseSensitive, QRegExp.RegExp)
        self.proxy.setFilterRegExp(search)

        if self.sql is None:
            newModel = self.view.model()
            data = []
            for row in range(newModel.rowCount()):
                rowRes = []
                for column in range(newModel.columnCount()):
                    index = newModel.index(row, column)
                    item = newModel.data(index)
                    if item.isdigit():
                        rowRes.append(int(item))
                    else:
                        rowRes.append(item)
                data.append(rowRes)
            data = DataFrame(data)
            
            lastRow = data.append(data.sum(numeric_only=True), ignore_index=True).tail(1).fillna('').values
            flat_Row = ['Total: ' + str(int(float(str(item)))) if item != '' else '' for sublist in lastRow for item in sublist]

            self.statusBar().removeWidget(self.mWdg)

            self.mWdg = QWidget()
            self.mWdg.setLayout(QHBoxLayout())
            for item in flat_Row:
                self.mWdg.layout().addWidget(QLabel(item))
            self.statusBar().addPermanentWidget(self.mWdg, 1)
      
    @pyqtSlot(int)
    def on_comboBox_currentIndexChanged(self, index):
        self.proxy.setFilterKeyColumn(index)

class PandasModel(QAbstractTableModel):
    def __init__(self, data, row, dups, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data
        self.r = row
        self.dup = dups

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
            if role == Qt.BackgroundColorRole and self.r is not None and self.dup is not None:
                bgColor=QColor(230,255,230)
                rColor=QColor(255,240,240)
                if index.row() > self.r and index.row() not in self.dup: 
                    return QVariant(QColor(bgColor))
                elif index.row() in self.dup:
                    return QVariant(QColor(rColor))
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

class ToWspec(QTableWidget):
    def __init__(self, parent = None):
        super(ToWspec, self).__init__(parent)
        
        self.rowCount = 20
        self.columnCount = 6

        self.setColumnCount(self.columnCount)
        self.setRowCount(self.rowCount)
        self.setHorizontalHeaderLabels(['Item',
                                        'Unit',
                                        'Volume',
                                        'Payment Type',
                                        'Amount Paid',
                                        'Comment'])
        self.horizontalHeader().setFixedHeight(80)
        self.verticalHeader().hide()

        self.setColumnWidth(5,450)
        
        for i in range(self.columnCount - 1):
            self.horizontalHeader().setSectionResizeMode(i, QHeaderView.Stretch)
            
        self.setStyleSheet( "QHeaderView::section{"
                            "border-top:1px solid #D8D8D8;"
                            "border-left:1px solid #D8D8D8;"
                            "border-right:2px solid #D8D8D8;"
                            "border-bottom: 3px solid #D8D8D8;"
                            "background-color:white;"
                            "padding:4px;""}")
        
##        sqm = u"\u33A1"
        self.offer1 = cost_dict.keys()
        self.offer2 = [' ', 'm', 'CZK']
        self.offer3 = [' ', 'Cash', 'Noncash Payment', 'Bank Transfer']

        for i in range(self.rowCount):
            comboA = myCombo()
            comboB = myCombo()
            comboC = myCombo() 
            spinBox = QSpinBox()
            spinBox.lineEdit().setAlignment(Qt.AlignCenter)
            spinBox.lineEdit().setReadOnly(True)
            LE = QLineEdit()
            onlyInt = QIntValidator()
            LE.setValidator(onlyInt)
            
            comboA.addItems(self.offer1)
            comboB.addItems(sorted(self.offer2))
            comboC.addItems(sorted(self.offer3))
            self.setCellWidget(i, 0, comboA)
            self.setCellWidget(i, 1, comboB)
            self.setCellWidget(i, 3, comboC)
            self.setCellWidget(i, 2, spinBox)
            self.setCellWidget(i, 4, LE)
            self.setItem(i,5,QTableWidgetItem(' '))

    def slot(self):
        myList = []
        for i in range(self.rowCount):
            output = []
            for j in range(self.columnCount):
                if j < 2:
                    w = self.cellWidget(i,j)
                    a = w.currentText()
                    output.append(a)
                if j == 2:
                    w = self.cellWidget(i,j)
                    a = str(w.value())
                    output.append(a)
                if j == 3:
                    w = self.cellWidget(i,j)
                    a = w.currentText()
                    output.append(a)
                if j == 4:
                    w = self.cellWidget(i,j)
                    a = w.text()
                    output.append(a)
                if j > 4:
                    w = self.item(i,j)
                    a = w.text()
                    output.append(a)
            if output[0] != ' ' or output[1] != ' ' or output[2] != '0' or output[3] != ' ' or output[4] != '' or output[5] != ' ':
                myList.append(output)
        return myList

class ModifiedModel(QAbstractTableModel):
    def __init__(self, data, parent=None):
        QAbstractTableModel.__init__(self, parent)
        self._data = data

    def rowCount(self, parent=None):
        return self._data.shape[0]

    def columnCount(self, parent=None):
        return self._data.shape[1]

    def flags(self, index):
        fl = super(self.__class__,self).flags(index)
        fl |= Qt.ItemIsEditable
        fl |= Qt.ItemIsSelectable
        fl |= Qt.ItemIsEnabled
        fl |= Qt.ItemIsDragEnabled
        fl |= Qt.ItemIsDropEnabled
        return fl

    def setData(self, index, value, role=Qt.EditRole):
        if index.isValid():
            row = index.row()
            col = index.column()
            self._data.iloc[row][col] = value
            self.dataChanged.emit(index, index, (Qt.DisplayRole, ))
            return True
        return False

    def data(self, index, role=Qt.DisplayRole):
        if index.isValid():
            if role == Qt.DisplayRole:
                return str(self._data.iloc[index.row(), index.column()])
        return None

    def headerData(self, col, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self._data.columns[col]
        return None

class Example(QWidget):
    def __init__(self, glVar):
        super().__init__()
        self.glVar = glVar
        self.init_UI()

    def init_UI(self):
        self.years = [' ', '2017', '2018', '2019', '2020']
        self.months = [' ', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12']
        self.metrics = [' ', 'Seller', 'Section', 'Store']
        self.units = [' ', 'Number of Pieces Sold', 'Total Value']
        self.setWindowTitle('Dialog')

        self.sMetric = QLabel('Select metric(s):')
        self.sUnit = QLabel('Select unit(s):')
        self.okButton = QPushButton('Ok')
        self.cancelButton = QPushButton('Cancel')

        self.okButton.clicked.connect(self.okClicked)
        self.cancelButton.clicked.connect(self.cancelClicked)

        for i in (self.sMetric, self.sUnit, self.okButton, self.cancelButton):
            i.setFixedHeight(40)
            i.setFixedWidth(150)

        self.timeWise = QCheckBox('Time-wise View')
        self.timeWise.toggled.connect(self.on_checked)

        self.ownLabels = QCheckBox('Own Labels Only')

        self.metricCombo = CheckableComboBox()
        for index, element in enumerate(self.metrics):
            self.metricCombo.addItem(element)
            if index > 0:
                item = self.metricCombo.model().item(index, 0)
                item.setCheckState(Qt.Unchecked)

        self.unitCombo = CheckableComboBox()
        for index, element in enumerate(self.units):
            self.unitCombo.addItem(element)
            if index > 0:
                item = self.unitCombo.model().item(index, 0)
                item.setCheckState(Qt.Unchecked)

        self.grid = QGridLayout()
        self.grid.setSpacing(10)

        self.grid.addWidget(self.sMetric, 1, 0)
        self.grid.addWidget(self.metricCombo, 1, 1)
        self.grid.addWidget(self.sUnit, 2, 0)
        self.grid.addWidget(self.unitCombo, 2, 1)
        self.grid.addWidget(self.timeWise, 3, 0)
        if self.glVar is None:
            self.grid.addWidget(self.ownLabels, 3, 1)
        self.grid.addWidget(self.okButton, 6, 0)
        self.grid.addWidget(self.cancelButton, 6, 1)

        self.setLayout(self.grid)

        self.sYear = QLabel('Select year(s):')
        self.sMonth = QLabel('Select month(s):')

        for i in (self.sYear, self.sMonth):
            i.setFixedHeight(40)
            i.setFixedWidth(150)

        self.monthCombo = CheckableComboBox()
        for index, element in enumerate(self.months):
            self.monthCombo.addItem(element)
            if index > 0:
                item = self.monthCombo.model().item(index, 0)
                item.setCheckState(Qt.Unchecked)

        self.yearCombo = CheckableComboBox()
        for index, element in enumerate(self.years):
            self.yearCombo.addItem(element)
            if index > 0:
                item = self.yearCombo.model().item(index, 0)
                item.setCheckState(Qt.Unchecked)

        self.grid.addWidget(self.sYear, 4, 0)
        self.grid.addWidget(self.yearCombo, 4, 1)
        self.grid.addWidget(self.sMonth, 5, 0)
        self.grid.addWidget(self.monthCombo, 5, 1)
        self.grid.setSizeConstraint(QLayout.SetFixedSize)

        self.on_checked(False)

    def on_checked(self, checked):
        self.sYear.setVisible(checked)
        self.sMonth.setVisible(checked)
        self.monthCombo.setVisible(checked)
        self.yearCombo.setVisible(checked)

    def cancelClicked(self):
        self.close()

    def okClicked(self):
        yearSelection = []
        monthSelection = []
        metricSelection = []
        unitSelection = []
        if self.timeWise.isChecked():
            for index, element in enumerate(self.years):
                item = self.yearCombo.model().item(index, 0)
                if item.checkState() == Qt.Checked:
                    yearSelection.append(item.text())
            for index, element in enumerate(self.months):
                item = self.monthCombo.model().item(index, 0)
                if item.checkState() == Qt.Checked:
                    monthSelection.append(item.text())
            for index, element in enumerate(self.metrics):
                item = self.metricCombo.model().item(index, 0)
                if item.checkState() == Qt.Checked:
                    metricSelection.append(item.text())
            for index, element in enumerate(self.units):
                item = self.unitCombo.model().item(index, 0)
                if item.checkState() == Qt.Checked:
                    unitSelection.append(item.text())
        else:
            for index, element in enumerate(self.metrics):
                item = self.metricCombo.model().item(index, 0)
                if item.checkState() == Qt.Checked:
                    metricSelection.append(item.text())
            for index, element in enumerate(self.units):
                item = self.unitCombo.model().item(index, 0)
                if item.checkState() == Qt.Checked:
                    unitSelection.append(item.text())

        data = []
        emptyL = []
        concL = []
        dictionary = {'Number of Pieces Sold' : 'nSumUnits', 'Total Value' : 'nSumValue'}
        
        if metricSelection == [] or unitSelection == []:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The form is not complete, please proceed with your selections and make sure that desired boxes are checked.', QMessageBox.Ok)
            msgbox.exec()
        elif self.timeWise.isChecked() and monthSelection == [] and yearSelection == []:
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The form is not complete, please proceed with your selections and make sure that desired boxes are checked.', QMessageBox.Ok)
            msgbox.exec()
        else:
            for i in monthSelection:
                if len(i) == 1:
                    cons = '-' + '0' + i + '-'
                else:
                    cons = '-' + i + '-'
                emptyL.append(cons)

            if len(yearSelection) > 1:
                for i in yearSelection:
                    for j in emptyL:
                        fullStr = i + j
                        concL.append(fullStr)
                        
            if len(yearSelection) == 1:
                for j in emptyL:
                    fullStr = yearSelection[0] + j
                    concL.append(fullStr)

            if len(yearSelection) > 0 and len(monthSelection) > 0:
                enVar = concL
            if len(yearSelection) > 0 and len(monthSelection) == 0:
                enVar = yearSelection
            if len(monthSelection) > 0 and len(yearSelection) == 0:
                enVar = emptyL
            if len(monthSelection) == 0 and len(yearSelection) == 0:
                enVar = None

            if self.glVar is None:
                if enVar is None:
                    if self.ownLabels.isChecked():
                        sql_query(sql_dict['Aggregated_view_OL']['sql'], [sql_dict['Aggregated_view']['args']])
                        data.append(sql_obj)
                    else:
                        sql_query(sql_dict['Aggregated_view']['sql'], [sql_dict['Aggregated_view']['args']])
                        data.append(sql_obj)
                else:
                    if self.ownLabels.isChecked():
                        for index, element in enumerate(enVar):
                            sql_query(sql_dict['Dynamic_sql_OL']['sql'], [element, '%'+element+'%'])
                            data.append(sql_obj)
                    else:
                        for index, element in enumerate(enVar):
                            sql_query(sql_dict['Dynamic_sql']['sql'], [element, '%'+element+'%'])
                            data.append(sql_obj)
            else:
                if enVar is None:
                    sql_query(sql_dict['Aggregated_view_R']['sql'], [sql_dict['Aggregated_view']['args']])
                    data.append(sql_obj)
                else:
                    for index, element in enumerate(enVar):
                        sql_query(sql_dict['Dynamic_sql_R']['sql'], [element, '%'+element+'%'])
                        data.append(sql_obj)

            flat_data = [item for sublist in data for item in sublist]
            df = DataFrame(flat_data)

            if df.empty:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'There is no data within the database that would match your year and/or month selections.', QMessageBox.Ok)
                msgbox.exec()
            else:
                df.columns = ['Section', 'Seller', 'Store', 'Selected_period', 'Total_units', 'Total_value']
                if df['Selected_period'].isnull().all():
                    listOfCols  = metricSelection
                else:
                    metricSelection.append('Selected_period')
                    listOfCols  = metricSelection
                    df['Selected_period'] = df['Selected_period'].replace(month_dict, regex = True)
                
                df['nSumValue'] = df.groupby(listOfCols)['Total_value'].transform('sum')
                df['nSumUnits'] = df.groupby(listOfCols)['Total_units'].transform('sum')

                if len(unitSelection) > 1:
                    for index, element in enumerate(itemgetter(*unitSelection)(dictionary)):
                        listOfCols.append(element)
                else:
                    listOfCols.append(itemgetter(*unitSelection)(dictionary))

                eList = []
                for element in listOfCols:
                    eList.append(True)

                preFinOutput = df[listOfCols]
                FinOutput = preFinOutput.drop_duplicates()
                sql = FinOutput.sort_values(by=listOfCols, ascending=eList)

                if 'nSumValue' in sql.columns:
                    sql = sql.rename(columns={'nSumValue': 'Total Value'})
                if 'nSumUnits' in sql.columns:
                    sql = sql.rename(columns={'nSumUnits': 'Total Pieces'})

                uCols = [x for x in listOfCols if x not in ('nSumValue', 'nSumUnits')]
                
                if 'Selected_period' in sql.columns:
                    sql = sql.set_index(uCols).unstack(level = -1).fillna(0).astype(int)
                    sql.columns = sql.columns.get_level_values(0) + ': ' + sql.columns.get_level_values(1)
                    sql = sql.reset_index()

                if any(sql.columns.str.contains('2017|2018|2019|2020')) or any(sql.columns.str.contains('Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec')):
                    
                    firstPart = [item for item in metricSelection if item not in ('nSumValue', 'nSumUnits', 'Selected_period')]
                    secondPart = [item for item in sql.columns.values if item not in firstPart]

                    renamed = DataFrame(enVar)[0].replace(month_dict, regex = True).values.tolist()
                    fullList = ['Total Value: ' + e for e in renamed] + ['Total Pieces: ' + e for e in renamed]
                    
                    reOrder = []
                    for item in fullList:
                        if item in secondPart:
                            reOrder.append(item)

                    sql = sql[firstPart + reOrder]
                    
                self.TBView = TBWindow(sql)
                self.TBView.showMaximized()
                self.close()
                
    def cancelClicked(self):
        self.close()

class PDF_gen(QObject):
    finished = pyqtSignal()
    intReady = pyqtSignal(int)

    if True:
        global sellerDF, dataR, sellers
        # Rents due for current month
        sql_query(sql_dict['Rent_addressee']['sql'], None)
        sellerDF = DataFrame(sql_obj, columns = headers)

        # Sold Items from previous month
        sql_query(sql_dict['PMSoldItems']['sql'], sql_dict['PMSoldItems']['args'])
        dataR = DataFrame(sql_obj, columns = headers)

        # All sellers
        sql_query(sql_dict['All_sellers']['sql'], None)
        sellers = DataFrame(sql_obj, columns = headers)

    global soldItems_Rent, soldItems_only, restList, AdList, backupList
    soldItems_Rent = [x for x in dataR[dataR.columns[0]].unique() if x in sellerDF[sellerDF.columns[0]].unique()]
    soldItems_only = [x for x in dataR[dataR.columns[0]].unique() if x not in sellerDF[sellerDF.columns[0]].unique()]
    restList = [x for x in sellerDF[sellerDF.columns[0]].unique() if x not in dataR[dataR.columns[0]].unique()]

    AdList = {}
    for index, item in enumerate(sellerDF[sellerDF.columns[0]]):
        AdList[item] = sellerDF[sellerDF.columns[1]].values.tolist()[index]

    backupList = {}
    for index, item in enumerate(sellers[sellers.columns[0]]):
        backupList[item] = sellers[sellers.columns[1]].values.tolist()[index]

    def docGeneration(self):
        
        # Doc files (sold items)
        for index, item in enumerate(dataR[dataR.columns[0]].unique()):
            label = item
            item = dataR.query('Label == "%s"' % item)
            lines = item[['Product_desc', 'Number_of_pieces', 'Total_value']].applymap(str).values.tolist()
            totals = item.sum(numeric_only=True).apply(str).tolist()    

            keys = ['ProdDesc', 'NPieces', 'TotalPrice']
            dictList = []

            for i in lines:
                i[2] = "%.2f" % round(float(i[2]),2)
                dictList.append(dict(zip(keys, i)))

            DocTemplate = inputs + 'SalesTemplate.docx'
            document = MailMerge(DocTemplate)
            document.merge(
                sPeriod = str(pmonthStart)[:10],
                ePeriod = str(ipmonthEnd)[:10],
                totalAmount = "%.2f" % round(float(totals[1]),2),
                dueDate = dDate)
            document.merge_rows('ProdDesc', dictList)
            document.write(tempfolder + 'S-' + unidecode.unidecode(label) + str((date.today()).strftime('%Y-%m-%d')) + '.docx')

        # Doc files (rents)
        for index, item in enumerate(sellerDF[sellerDF.columns[0]].unique()):
            if sellerDF.values.tolist()[index][7] is None:
                InvoiceN = sellerDF.values.tolist()[index][9][4:]
            else:
                InvoiceN = sellerDF.values.tolist()[index][7]
            
            DocTemplate = inputs + 'RentsTemplate.docx'
            document = MailMerge(DocTemplate)
            document.merge(
                InvoiceNo = str((date.today()).strftime('%Y')) + '/' + InvoiceN + '/' + str((date.today()).strftime('%m')),
                FullName = sellerDF.values.tolist()[index][2],
                Street = sellerDF.values.tolist()[index][3],
                Postcode = sellerDF.values.tolist()[index][5],
                Town = sellerDF.values.tolist()[index][4],
                Country = sellerDF.values.tolist()[index][6],
                ID_number = sellerDF.values.tolist()[index][7],
                SubStr = InvoiceN,
                IssueDate = str((date.today()).strftime('%Y-%m-%d')),
                DueDate = str((date.today()).strftime('%Y-%m-%d'))[:8] + '10',
                TaxDueDate = str((date.today()).strftime('%Y-%m-%d'))[:8] + '10',
                TotalValue = str(sellerDF.values.tolist()[index][10]),
                ContractedDate = sellerDF.values.tolist()[index][8])
            document.write(tempfolder + 'R-' + unidecode.unidecode(item) + str((date.today()).strftime('%Y-%m-%d')) + '.docx') 
    
    @pyqtSlot()
    def procCounter(self):
        pythoncom.CoInitialize()
        word = client.DispatchEx("Word.Application")
        for index, item in enumerate(os.listdir(tempfolder)):
            self.intReady.emit((index + 1)*100/folderLen)
            if 'R-' in item or 'S-' in item:
                PDF_creator(item, word)
##            else:
##                try:
##                    os.remove(item)
##                except OSError:
##                    pass
        word.Quit()
        self.finished.emit()

class EM_gen(QObject):
    finished = pyqtSignal()
    intReady = pyqtSignal(int)

    @pyqtSlot()
    def procCounter(self):
        for index, item in enumerate(soldItems_Rent + soldItems_only + restList):
            self.intReady.emit(index * len(soldItems_Rent + soldItems_only + restList))
            send_to = AdList.get(item)
            if send_to is None:
                send_to = backupList.get(item)
                if send_to is None:
                    send_to = 'chroustovskyjiri@gmail.com'#'info@princezna-pampeliska.cz'
            if item in soldItems_Rent:
                out_files = [solditems + 'S-' + unidecode.unidecode(item) + str((date.today()).strftime('%Y-%m-%d')) + '.pdf', rents + 'R-' + unidecode.unidecode(item) + str((date.today()).strftime('%Y-%m-%d')) + '.pdf']
                subject = 'Prehled prodeju a faktura za souvisejici marketingove sluzby'
                message = 'Toto je automaticky generovany email. V priloze naleznete prehled prodeju Vaseho zbozi v kamenne prodejne Princezny Pampelisky v Holesovicich a fakturu za marketingove sluzby s timto prodejem souvisejici.'
            elif item in soldItems_only:
                out_files = [solditems + 'S-' + unidecode.unidecode(item) + str((date.today()).strftime('%Y-%m-%d')) + '.pdf']
                subject = 'Prehled prodeju'
                message = 'Toto je automaticky generovany email. V priloze naleznete prehled prodeju Vaseho zbozi v kamenne prodejne Princezny Pampelisky v Holesovicich.'
            elif item in restList:
                out_files = [rents + 'R-' + unidecode.unidecode(item) + str((date.today()).strftime('%Y-%m-%d')) + '.pdf']
                subject = 'Faktura za poskytnute marketingove sluzby'
                message = 'Toto je automaticky generovany email. V priloze naleznete fakturu za marketingove sluzby souvisejici s prodejem Vaseho zbozi v kamenne prodejne Princezny Pampelisky v Holesovicich.'

            server = 'smtp.gmail.com'
            port = 587
            username = 'jana.klapetkova@gmail.com'#'tmp.testem@gmail.com'
            password = 'Bpv15v5blc'#'ocepo123'

            msg = MIMEMultipart()
            msg['From'] = username
            msg['To'] = send_to
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = subject
            msg.attach(MIMEText(message))

            for attachment in (out_files):
                part = MIMEBase('application', "octet-stream")
                with open(attachment, 'rb') as file:
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="{}"'.format(os.path.basename(attachment)))
                msg.attach(part)

            smtp = smtplib.SMTP(server, port)
            smtp.ehlo()
            smtp.starttls()
            smtp.login(username, password)
            smtp.sendmail(username, send_to, msg.as_string())
            smtp.quit()
        self.finished.emit()

class SRreport(QWidget):
    def __init__(self):
        QWidget.__init__(self)

        self.setWindowTitle('Progress')
        self.label1 = QLabel('Progress in PDF files creation:')
        self.label2 = QLabel('Progress in sending emails with attachments:')

        layout = QVBoxLayout(self)
        self.progressBar1 = QProgressBar()
        self.progressBar2 = QProgressBar()

        StartButton = QPushButton('Start')
        CancelButton = QPushButton('Cancel')
        StatusBar = QStatusBar()
        StatusBar.addPermanentWidget(StartButton, 1)
        StatusBar.addPermanentWidget(CancelButton, 1)

        layout.addWidget(self.label1)
        layout.addWidget(self.progressBar1)
        layout.addWidget(self.label2)
        layout.addWidget(self.progressBar2)
        layout.addWidget(StatusBar)

        StartButton.clicked.connect(self.StartButtonClicked)
        CancelButton.clicked.connect(self.CancelButtonClicked)

    def StartButtonClicked(self):
        self.writeDB()
        self.thread = QThread()
        self.f = PDF_gen()
        self.f.docGeneration()
        global folderLen
        folderLen = len(os.listdir(tempfolder))
        self.s = EM_gen()
        self.f.moveToThread(self.thread)
        self.s.moveToThread(self.thread)
        self.thread.started.connect(self.f.procCounter)
        self.thread.started.connect(self.s.procCounter)
        self.thread.start()
        self.f.finished.connect(self.thread.start)
        self.f.intReady.connect(self.progressBar1.setValue, Qt.QueuedConnection)
        self.s.intReady.connect(self.progressBar2.setValue, Qt.QueuedConnection)
        self.s.finished.connect(self.closeW)

    def CancelButtonClicked(self):
        self.close()

    def writeDB(self):
        sql_query(sql_dict['rActData']['sql'], sql_dict['rActData']['args'])
        data = DataFrame(sql_obj, columns = headers)
        data['Event_time'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data = data[['Rent_ID', 'Rent_Amount', 'Seller', 'Event_time', 'Rent_value', 'Payment_type', 'Outstanding']]
        finList = data.values.tolist()
        for row in finList:
            sql_query(sql_dict['NewRents']['sql'], row)

        sql_query(sql_dict['sPayOff']['sql'], sql_dict['sPayOff']['args'])
        data = DataFrame(sql_obj, columns = headers)
        data['pfC'] = 'vyplata za prodane zbozi v obdobi %s - %s, prodejce: ' % ((pmonthStart), (pmonthEnd))
        data['Comment'] = data['pfC'] + data['Label']
        data = data[['Item', 'Unit', 'Volume', 'Payment_type', 'Amount_paid', 'Comment', 'Type', 'Cost_ID']]
        data_f = data.values.tolist()
        Cost_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        Outstanding = 'Y'
        for i, j in enumerate(data_f):
            letter = chr(i+97)
            identifier = Cost_time.replace(' ',letter)
            j.extend([Cost_time, identifier])
            j[6], j[7], j[8], j[9] = j[8], j[9], j[6], j[7]
            j[4] = int(j[4])#it rounds down like a floor function from math library
            j.append(Outstanding)
            sql_query(sql_dict['NewCosts']['sql'], j)

    def closeW(self):
        self.close()

class Form(QMainWindow):
    def __init__(self, glVar):
        QMainWindow.__init__(self)
        self.glVar = glVar
        self.initUI()

    def initUI(self):

        self.setWindowTitle(' ')
        self.runButton = QPushButton('Save')
        self.CButton = QPushButton('Cancel')
        self.CButton.clicked.connect(self.onCancelClicked)
        self.toolBar = self.addToolBar('oneAct')
        self.toolBar.addWidget(self.runButton)
        self.toolBar.addWidget(self.CButton)

        if self.glVar == None:
            self.runButton.clicked.connect(self.cInfo)
        elif self.glVar == 'AddModPayer':
            self.runButton.clicked.connect(self.AddcInfo)

        self.createFormGroupBox()

        self.cWidget = QWidget()
        self.mainLayout = QVBoxLayout()

        self.mainLayout.addWidget(self.formGroupBox)
        self.cWidget.setLayout(self.mainLayout)

        self.setCentralWidget(self.cWidget)

        self.infoWidget = QLabel('   Info: all fields with * are mandatory.')
        self.statusBar().addPermanentWidget(self.infoWidget, 1)

    def createFormGroupBox(self):

        self.rentItemsDict = {
                        '' : None,
                        'Malá polička' : 'hrnt01',
                        'Dlouhá polička' : 'hrnt02',
                        'Polovina skříňky' : 'hrnt03',
                        'Celá skříňka' : 'hrnt04',
                        'Místo nahoře na skříni' : 'hrnt05',
                        '7 háčků' : 'hrnt06',
                        '10 háčků' : 'hrnt07',
                        '14 háčků' : 'hrnt08',
                        '28 háčků' : 'hrnt09',
                        'Jiné služby' : 'hrnt10',
                        'Mix nájemních ploch' : 'hrnt11'
                        }
        
        self.loop0 = QLabel('Label:*')
        self.loop1 = QLabel('Email address:*')
        self.loop2 = QLabel('Full name:*')
        self.loop3 = QLabel('Phone number:*')
        self.loop4 = QLabel('Street:*')
        self.loop5 = QLabel('Town:*')
        self.loop6 = QLabel('Postcode:*')
        self.loop7 = QLabel('Country:*')
        self.loop8 = QLabel('ID number:')
        self.loop9 = QLabel('BA prefix:')
        self.loop10 = QLabel('BA number:')
        self.loop11 = QLabel('BA code:')
        self.loop12 = QLabel('Rent type:*')   
        self.loop13 = QLabel('Rent amount:*')
        self.loop14 = QLabel('Active payer:*')
        self.loop15 = QLabel('Comment:')
        self.loop16 = QLabel('Contract since:*')
        
        sql_query(sql_dict['DistinctPayers']['sql'], None)
        sellers = [''] + [x[0] for x in sql_obj]

        self.newLine = QComboBox()
        self.newLine.addItems(sellers)
        self.Line0 = QLineEdit()
        self.Line1 = QLineEdit()
        self.Line2 = QLineEdit()
        self.Line3 = QLineEdit()
        self.Line4 = QLineEdit()
        self.Line5 = QLineEdit()
        self.Line6 = QLineEdit()
        self.Line7 = QLineEdit()
        self.Line8 = QLineEdit()
        self.Line9 = QLineEdit()
        self.Line10 = QLineEdit()
        self.Line11 = QLineEdit()
        self.Line12 = QComboBox()
        self.Line12.addItems(self.rentItemsDict.keys())
        self.Line13 = QLineEdit()

        self.Line13.setText(str(1))
        onlyInteger = QIntValidator()
        self.Line13.setValidator(onlyInteger)
        
        self.Line14 = QComboBox()
        self.Line14.addItems([' ', 'Y', 'N'])
        self.Line15 = QComboBox()
        self.Line15.addItems([' ', 'Commission'])
        self.Line16 = QHBoxLayout()

        self.yearEdit = QDateTimeEdit()
        self.yearEdit.setDisplayFormat('yyyy')
        self.yearEdit.setDateRange(QDate(1753, 1, 1), QDate(8000, 1, 1))
        self.yearEdit.setDate(QDate.currentDate())

        self.monthEdit = QComboBox()
        monthList = month_dict.values()
        monthList = [x[1:] for x in monthList]
        self.monthEdit.addItems(monthList)

        self.dayEdit = QDateTimeEdit()
        self.dayEdit.setDisplayFormat('dd')
        self.dayEdit.setDate(QDate.currentDate())

        self.Line16.addWidget(self.yearEdit)
        self.Line16.addWidget(self.monthEdit)
        self.Line16.addWidget(self.dayEdit)
        
        self.formGroupBox = QGroupBox('Seller details')
        layout = QFormLayout()
        
        if self.glVar == None:
            layout.addRow(self.loop0, self.Line0)
            layout.addRow(self.loop1, self.Line1)
            layout.addRow(self.loop2, self.Line2)
            layout.addRow(self.loop3, self.Line3)
            layout.addRow(self.loop4, self.Line4)
            layout.addRow(self.loop5, self.Line5)
            layout.addRow(self.loop6, self.Line6)
            layout.addRow(self.loop7, self.Line7)
            layout.addRow(self.loop8, self.Line8)
            layout.addRow(self.loop9, self.Line9)
            layout.addRow(self.loop10, self.Line10)
            layout.addRow(self.loop11, self.Line11)
            layout.addRow(self.loop12, self.Line12)
            layout.addRow(self.loop13, self.Line13)
            layout.addRow(self.loop14, self.Line14)
            layout.addRow(self.loop15, self.Line15)
            layout.addRow(self.loop16, self.Line16)
        elif self.glVar == 'AddModPayer':
            layout.addRow(self.loop0, self.newLine)
            layout.addRow(self.loop12, self.Line12)
            layout.addRow(self.loop13, self.Line13)
            layout.addRow(self.loop14, self.Line14)
            layout.addRow(self.loop15, self.Line15)
            layout.addRow(self.loop16, self.Line16)
        
        self.formGroupBox.setLayout(layout)

    def onCancelClicked(self):
        self.close()

    def AddcInfo(self):
        try:
            emptyL = []
            for item in (self.newLine, self.Line12, self.Line13, self.Line14, self.Line15):
                if item == self.Line13:
                    if item.text() == '':
                        element = None
                    else:
                        element = item.text()
                    emptyL.append(element)
                elif item in (self.newLine, self.Line12, self.Line14, self.Line15):
                    if str(item.currentText()) == '' or str(item.currentText()) == ' ':
                        element = None
                    else:
                        element = str(item.currentText())
                    emptyL.append(element)
            emptyL.append(str(self.yearEdit.text()) + dict((v,k) for k,v in (month_dict.items())).get(' ' + str(self.monthEdit.currentText())) + str(self.dayEdit.text()))
            emptyL[1] = self.rentItemsDict.get(emptyL[1])
            if None in (emptyL[0], emptyL[1], emptyL[2], emptyL[3], emptyL[5]):
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'You need to complete the form before proceeding.', QMessageBox.Ok)
                    msgbox.exec()
            else:
                sql_query(sql_dict['DisplayModPayers']['sql'], None)
                res = [(row[0],row[1]) for row in sql_obj]
                if (emptyL[0],emptyL[1]) in res:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'This combination of rent payer and rent type already exists within the DataBase.', QMessageBox.Ok)
                    msgbox.exec()
                else:
                    sql_query(sql_dict['InsertNewPayer']['sql'], emptyL)
                    self.close()
        except Exception as e:
            print(e)

    def cInfo(self):
        try:
            emptyL = []
            for item in (self.Line0, self.Line1, self.Line2, self.Line3, self.Line4, self.Line5, self.Line6, self.Line7, self.Line8, self.Line9, self.Line10, self.Line11, self.Line12, self.Line13, self.Line14, self.Line15, self.Line16):
                if item in (self.Line0, self.Line1, self.Line2, self.Line3, self.Line4, self.Line5, self.Line6, self.Line7, self.Line8, self.Line9, self.Line10, self.Line11, self.Line13):
                    if item.text() == '':
                        element = None
                    else:
                        element = item.text()
                    emptyL.append(element)
                elif item in (self.Line12, self.Line14, self.Line15):
                    if str(item.currentText()) == '' or str(item.currentText()) == ' ':
                        element = None
                    else:
                        element = str(item.currentText())
                    emptyL.append(element)
            emptyL.append(str(self.yearEdit.text()) + dict((v,k) for k,v in (month_dict.items())).get(' ' + str(self.monthEdit.currentText())) + str(self.dayEdit.text()))
            if None in (emptyL[0], emptyL[1], emptyL[2], emptyL[3], emptyL[4], emptyL[5], emptyL[6], emptyL[7], emptyL[12], emptyL[13], emptyL[14], emptyL[16]):
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'You need to complete the form before proceeding.', QMessageBox.Ok)
                msgbox.exec()
            else:
                sql_query(sql_dict['DisplayActStock']['sql'], None)
                pdf = DataFrame(sql_obj)
                if pdf[3].str.contains(emptyL[0]).any():
                    sql_query(sql_dict['ValidateSeller']['sql'], None)
                    anotherPDF = DataFrame(sql_obj)
                    if anotherPDF[0].str.contains(emptyL[0]).any():
                        msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'This seller already exists within Sellers table, it is not possible to add any seller more than one time.', QMessageBox.Ok)
                        msgbox.exec()
                    else:
                        RentPayers = [emptyL[0], self.rentItemsDict.get(emptyL[12]), emptyL[13], emptyL[14], emptyL[15], emptyL[16]]
                        Sellers = [emptyL[0], emptyL[1], emptyL[2], emptyL[3], emptyL[4], emptyL[5], emptyL[6], emptyL[7], emptyL[8], emptyL[9], emptyL[10], emptyL[11]]
                        sql_query(sql_dict['InsertRentPayers']['sql'], RentPayers)
                        sql_query(sql_dict['InsertSellers']['sql'], Sellers)
                        self.close()
                else:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'No products of this seller found within the DB, please add the products first, then you can proceed with adding the seller.', QMessageBox.Ok)
                    msgbox.exec()
        except Exception as e:
            print(e)

class UserForm(QMainWindow):
    def __init__(self, sql):
        QMainWindow.__init__(self)
        self.sql = sql
        self.initUI()

    def initUI(self):
        
        self.createFormGroupBox()
        self.cWidget = QWidget()
        self.mainLayout = QVBoxLayout()
        self.mainLayout.addWidget(self.formGroupBox)
        self.cWidget.setLayout(self.mainLayout)
        self.setCentralWidget(self.cWidget)

        sql_query(sql_dict['LogIn']['sql'], sql_dict['LogIn']['args'])
        self.logDict = dict(sql_obj)

    def createFormGroupBox(self):
        
        if self.sql == 'Change':
            self.setWindowTitle(' ')
            self.runButton = QPushButton('Save')
            self.runButton.clicked.connect(self.SaveC)
            self.CButton = QPushButton('Cancel')
            self.CButton.clicked.connect(self.onCancelClicked)
            self.statusBar().addPermanentWidget(self.runButton)
            self.statusBar().addPermanentWidget(self.CButton)

            self.loop0 = QLabel('User:')
            self.loop1 = QLabel('Old Password:')
            self.loop2 = QLabel('New Password:')
            self.Line0 = QLineEdit()
            self.Line0.setText(uNm)
            self.Line1 = QLineEdit()
            self.Line2 = QLineEdit()
            self.formGroupBox = QGroupBox('User details')
            layout = QFormLayout()
            layout.addRow(self.loop0, self.Line0)
            layout.addRow(self.loop1, self.Line1)
            layout.addRow(self.loop2, self.Line2)
            self.formGroupBox.setLayout(layout)
            
        if self.sql == 'Remove':
            self.setWindowTitle(' ')
            self.runButton = QPushButton('Remove')
            self.runButton.clicked.connect(self.onDel)
            self.CButton = QPushButton('Cancel')
            self.CButton.clicked.connect(self.onCancelClicked)
            self.statusBar().addPermanentWidget(self.runButton)
            self.statusBar().addPermanentWidget(self.CButton)

            self.loop0 = QLabel('User:')
            self.loop1 = QLabel('Password:')
            self.Line0 = QLineEdit()
            self.Line1 = QLineEdit()
            self.formGroupBox = QGroupBox('User details')
            layout = QFormLayout()
            layout.addRow(self.loop0, self.Line0)
            layout.addRow(self.loop1, self.Line1)
            self.formGroupBox.setLayout(layout)

        if self.sql == 'Add':
            self.setWindowTitle(' ')
            self.runButton = QPushButton('Add')
            self.runButton.clicked.connect(self.onAdd)
            self.CButton = QPushButton('Cancel')
            self.CButton.clicked.connect(self.onCancelClicked)
            self.statusBar().addPermanentWidget(self.runButton)
            self.statusBar().addPermanentWidget(self.CButton)

            self.loop0 = QLabel('User:')
            self.loop1 = QLabel('Password:')
            self.Line0 = QLineEdit()
            self.Line1 = QLineEdit()
            self.formGroupBox = QGroupBox('User details')
            layout = QFormLayout()
            layout.addRow(self.loop0, self.Line0)
            layout.addRow(self.loop1, self.Line1)
            self.formGroupBox.setLayout(layout)

    def onCancelClicked(self):
        self.close()

    def SaveC(self):
        if self.Line1.text() == '' or self.Line1.text() == ' ' or self.Line0.text() == '' or self.Line0.text() == ' ' or self.Line2.text() == '' or self.Line2.text() == ' ':
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Please fill in the form before proceeding!', QMessageBox.Ok)
            msgbox.exec()
        else:
            if self.Line1.text() == self.logDict[self.Line0.text()]:
                sql_query(sql_dict['ChangeP']['sql'], (self.Line2.text(),self.Line0.text()))
                self.close()
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The password has been successfully changed.', QMessageBox.Ok)
                msgbox.exec()
            else:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Please re-enter correct old password.', QMessageBox.Ok)
                msgbox.exec()

    def onDel(self):
        if self.Line1.text() == '' or self.Line1.text() == ' ' or self.Line0.text() == '' or self.Line0.text() == ' ':
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Please fill in the form before proceeding!', QMessageBox.Ok)
            msgbox.exec()
        else:
            if self.Line1.text() == self.logDict[self.Line0.text()] and uNm in ('Jana', 'Admin'):
                sql_query(sql_dict['RemoveU']['sql'], (self.Line0.text(),))
                self.close()
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The user has been permanently deleted from the DataBase.', QMessageBox.Ok)
                msgbox.exec()
            else:
                if uNm in ('Jana', 'Admin'):
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The password you entered is incorrect.', QMessageBox.Ok)
                    msgbox.exec()
                else:
                    msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'You do not have an authorization for this action.', QMessageBox.Ok)
                    msgbox.exec()                

    def onAdd(self):
        if self.Line1.text() == '' or self.Line1.text() == ' ' or self.Line0.text() == '' or self.Line0.text() == ' ':
            msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'Please fill in the form before proceeding!', QMessageBox.Ok)
            msgbox.exec()
        else:
            if uNm in ('Jana', 'Admin'):
                sql_query(sql_dict['AddU']['sql'], (self.Line0.text(),self.Line1.text()))
                self.close()
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'The new user has been successfully added into the Database.', QMessageBox.Ok)
                msgbox.exec()
            else:
                msgbox = QMessageBox(QMessageBox.Information, 'Dialog', 'You do not have an authorization for this action.', QMessageBox.Ok)
                msgbox.exec()
