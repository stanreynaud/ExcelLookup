from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
import xlrd
import xlwt
import sys
import os.path


class Window(QMainWindow):
    height=600
    width=600

    base=""
    search=""

    basecolchoice= -1
    searchcolchoice= -1

    def __init__(self):
        QWidget.__init__(self)
        super().__init__()

        self.setFont(QFont('Calibri', 10))
        
        
        self.head = QLabel(self)
        self.head.setGeometry(20,10,600,70)
        self.head.setText("-Check terminal for progress if program not responding                               -A=row 0   B=row 1   ...\n-Case insensitive comparison used\n-Excel date are not processed correctly (converted to float)\n-Automatically copies the first line (header) in result file")
        self.head.setFont(QFont('Calibri', 8))

        self.progress = QProgressBar(self)
        self.progress.setGeometry(10,80,580,20)
        self.explainlabel = QLabel(self)
        self.explainlabel.setGeometry(170,100,350,20)
        self.explainlabel.setText("I use data from base file to search in search file")

        self.baselabel = QLabel(self)
        self.baselabel.setGeometry(20,130,200,20)
        self.baselabel.setText("Name of base file  (.xls , .xlsx)")
        self.textboxbase = QLineEdit(self)
        self.textboxbase.setGeometry(20,150,200,20)
        self.textboxbase.returnPressed.connect(self.basecheckrow)
        self.basebutton = QPushButton("Update base file",self)
        self.basebutton.setGeometry(20,180,150,25)
        self.basebutton.clicked.connect(self.basecheckrow)
        self.baselabel = QLabel(self)
        self.baselabel.setGeometry(20,205,200,20)
        self.basecol = QLabel(self)
        self.basecol.setGeometry(20,230,200,20)
        self.basedrop = QComboBox(self)
        self.basedrop.setGeometry(140,230,40,20)
        self.basedrop.hide()
        self.basedrop.activated[str].connect(self.basechoice)


        self.searchlabel = QLabel(self)
        self.searchlabel.setGeometry(self.width-220,130,200,20)
        self.searchlabel.setText("Name of search file (.xls , .xlsx)")
        self.textboxsearch = QLineEdit(self)
        self.textboxsearch.setGeometry(self.width-220,150,200,20)
        self.textboxsearch.returnPressed.connect(self.searchcheckrow)
        self.searchbutton = QPushButton("Update search file",self)
        self.searchbutton.setGeometry(self.width-220,180,150,25)
        self.searchbutton.clicked.connect(self.searchcheckrow)
        self.searchlabel = QLabel(self)
        self.searchlabel.setGeometry(380,205,200,20)
        self.searchcol = QLabel(self)
        self.searchcol.setGeometry(self.width-220,230,200,20)

        self.searchdrop = QComboBox(self)
        self.searchdrop.setGeometry(500,230,40,20)
        self.searchdrop.hide()
        self.searchdrop.activated[str].connect(self.searchchoice)



        self.happenlabel = QLabel(self)
        self.happenlabel.setGeometry(20,320,500,60)
        self.happenlabel.setText("--- RESULT FILE --- \n\nI want everything that is in base file and that                                in search file")
        self.happenlabel.hide()

        self.isdrop = QComboBox(self)
        self.isdrop.setGeometry(273,357,80,20)
        self.isdrop.addItem("")
        self.isdrop.addItem("is")
        self.isdrop.addItem("is not")
        self.isdrop.activated[str].connect(self.isisnot)
        self.isdrop.hide()

        
        self.resultlabel = QLabel(self)
        self.resultlabel.setGeometry(20,380,200,50)
        self.resultlabel.setText("Name of result file: \n (no extension)")
        self.resultlabel.hide()

        self.resultbox = QLineEdit(self)
        self.resultbox.setGeometry(135,390,200,20)
        self.resultbox.returnPressed.connect(self.go)
        self.resultbox.hide()


        self.lineslabel = QLabel(self)
        self.lineslabel.setGeometry(20,420,220,50)
        self.lineslabel.setText("I want  a copy of match lines from ")
        self.lineslabel.hide()

        self.linesdrop = QComboBox(self)
        self.linesdrop.setGeometry(215,434,100,20)
        self.linesdrop.addItem("")
        self.linesdrop.addItem("base file")
        self.linesdrop.addItem("search file")
        self.linesdrop.addItem("both files")
        #self.linesdrop.activated[str].connect(self.lineschoice)
        self.linesdrop.hide()
        
        self.setWindowTitle("2880")
        self.setWindowIcon(QIcon("a.PNG"))
        self.setGeometry(0,0,self.height,self.width)
        #pop = QMessageBox.question(self,"PoP","close?", QMessageBox.Yes | QMessageBox.No,QMessageBox.No)



        self.cross = QLabel(self)
        self.cross.setGeometry(433,460,220,100)
        self.cross.setText("______________________")
        self.cross3 = QLabel(self)
        self.cross3.setGeometry(410,460,220,100)

        self.cross2 = QLabel(self)
        self.cross2.setGeometry(497,420,20,200)
        self.cross2.setText("|\n|\n|\n|\n|\n|\n|\n|\n|")
        self.c = QPushButton("Go",self)
        self.c.setGeometry(475,500,50,40)
        self.c.clicked.connect(self.go)


        
    def change(self):
        if self.d.checkState() == Qt.Checked:
            print("True")
        else:
            print("False")
            
    def fdef(self):
        self.progressnum=0
        while self.progressnum<1000:
            self.progressnum += 0.0001
            self.progress.setValue(self.progressnum)


    def basechoice(self, text):
        if text != "":
            self.basecolchoice = int(text)
        if(self.basecolchoice is not -1) and (self.searchcolchoice is not -1):
            self.happenlabel.show()
            self.isdrop.show()
            self.happenlabel.show()
            self.resultbox.show()
            self.resultlabel.show()
            
        return text
        
    def searchchoice(self, text):
        if text != "":
            self.searchcolchoice = int(text)
        if(self.basecolchoice is not -1) and (self.searchcolchoice is not -1):
            self.happenlabel.show()
            self.isdrop.show()
            self.happenlabel.show()
            self.resultbox.show()
            self.resultlabel.show()
        return text

    def fdefr(self,text):
        print(text)
        
    def go(self):
        if(self.isdrop.currentText() == "is"):
            if(self.linesdrop.currentText() == "base file"):
                sep = xlrd.open_workbook(self.base)
                flow = xlrd.open_workbook(self.search)


                flow_computername_position = self.searchcolchoice
                sep_computername_position = self.basecolchoice

                result = xlwt.Workbook()
                result_page = result.add_sheet("Inventory")

                next=1
                #Recopie du titre des colonnes
                for i in range(sep.sheet_by_index(0).ncols):
                    result_page.write(0,i,sep.sheet_by_index(0).cell(0,i).value)
                    
                self.progress.setMaximum(sep.sheet_by_index(0).nrows)
                print("base:",sep.sheet_by_index(0).nrows)
                print("search:",flow.sheet_by_index(0).nrows)
                for seprow in range(1,sep.sheet_by_index(0).nrows):
                    sepname = sep.sheet_by_index(0).cell(seprow,sep_computername_position).value
                    for flowrow in range(1,flow.sheet_by_index(0).nrows): 
                        if str(sepname).lower() == str(flow.sheet_by_index(0).cell(flowrow,flow_computername_position).value).lower():
                            for i in range(sep.sheet_by_index(0).ncols):
                                result_page.write(next,i,str(sep.sheet_by_index(0).cell(seprow,i).value))
                            next+=1
                            break
                    self.progress.setValue(seprow)
                    if ((seprow*100)/sep.sheet_by_index(0).nrows)%1<=0.01:
                        print('%.0f'%((seprow*100)/sep.sheet_by_index(0).nrows),"%")
                self.progress.setValue(sep.sheet_by_index(0).nrows)
                print("100 %")
                result.save(self.resultbox.text()+".xls")
                
            elif(self.linesdrop.currentText() == "search file"):
                sep = xlrd.open_workbook(self.base)
                flow = xlrd.open_workbook(self.search)


                flow_computername_position = self.searchcolchoice
                sep_computername_position = self.basecolchoice

                result = xlwt.Workbook()
                result_page = result.add_sheet("Inventory")

                next=1
                #Recopie du titre des colonnes
                for i in range(flow.sheet_by_index(0).ncols):
                    result_page.write(0,i,flow.sheet_by_index(0).cell(0,i).value)
                    
                self.progress.setMaximum(sep.sheet_by_index(0).nrows)
                print("base:",sep.sheet_by_index(0).nrows)
                print("search:",flow.sheet_by_index(0).nrows)
                for seprow in range(1,sep.sheet_by_index(0).nrows):
                    sepname = sep.sheet_by_index(0).cell(seprow,sep_computername_position).value
                    for flowrow in range(1,flow.sheet_by_index(0).nrows): 
                        if str(sepname).lower() == str(flow.sheet_by_index(0).cell(flowrow,flow_computername_position).value).lower():
                            for i in range(flow.sheet_by_index(0).ncols):
                                result_page.write(next,i,str(flow.sheet_by_index(0).cell(flowrow,i).value))
                            next+=1
                            break
                    self.progress.setValue(seprow)
                    if ((seprow*100)/sep.sheet_by_index(0).nrows)%1<=0.01:
                        print('%.0f'%((seprow*100)/sep.sheet_by_index(0).nrows),"%")
                self.progress.setValue(sep.sheet_by_index(0).nrows)
                print("100 %")
                result.save(self.resultbox.text()+".xls")
                

            elif(self.linesdrop.currentText() == "both files"):
                sep = xlrd.open_workbook(self.base)
                flow = xlrd.open_workbook(self.search)


                flow_computername_position = self.searchcolchoice
                sep_computername_position = self.basecolchoice

                result = xlwt.Workbook()
                result_page = result.add_sheet("Inventory")

                next=1
                #Recopie du titre des colonnes
                for i in range(sep.sheet_by_index(0).ncols):
                    result_page.write(0,i,sep.sheet_by_index(0).cell(0,i).value)
                for j in range(flow.sheet_by_index(0).ncols):
                    result_page.write(0,sep.sheet_by_index(0).ncols+j+1,flow.sheet_by_index(0).cell(0,j).value)
                    
                self.progress.setMaximum(sep.sheet_by_index(0).nrows)
                print("base:",sep.sheet_by_index(0).nrows)
                print("search:",flow.sheet_by_index(0).nrows)
                for seprow in range(1,sep.sheet_by_index(0).nrows):
                    sepname = sep.sheet_by_index(0).cell(seprow,sep_computername_position).value
                    for flowrow in range(1,flow.sheet_by_index(0).nrows): 
                        if str(sepname).lower() == str(flow.sheet_by_index(0).cell(flowrow,flow_computername_position).value).lower():
                            for i in range(sep.sheet_by_index(0).ncols):
                                result_page.write(next,i,sep.sheet_by_index(0).cell(seprow,i).value)
                            for j in range(flow.sheet_by_index(0).ncols):
                                result_page.write(next,sep.sheet_by_index(0).ncols+j+1,flow.sheet_by_index(0).cell(flowrow,j).value)
                            next+=1
                            break
                    self.progress.setValue(seprow)
                    if ((seprow*100)/sep.sheet_by_index(0).nrows)%1<=0.01:
                        print('%.0f'%((seprow*100)/sep.sheet_by_index(0).nrows),"%")
                self.progress.setValue(sep.sheet_by_index(0).nrows)
                print("100 %")
                result.save(self.resultbox.text()+".xls")
                
        elif(self.isdrop.currentText() == "is not"):
            base = xlrd.open_workbook(self.base)
            authorized = xlrd.open_workbook(self.search)

            authorized_computername_position = self.searchcolchoice
            base_computername_position = self.basecolchoice

            next=1
            result = xlwt.Workbook()
            result_page = result.add_sheet("Inventory")
            
            for i in range(base.sheet_by_index(0).ncols):
                result_page.write(0,i,base.sheet_by_index(0).cell(0,i).value)

            self.progress.setMaximum(base.sheet_by_index(0).nrows)
            print("base:",base.sheet_by_index(0).nrows)
            print("search:",authorized.sheet_by_index(0).nrows)
            for baserow in range(1,base.sheet_by_index(0).nrows):
                found=False
                basename = base.sheet_by_index(0).cell(baserow,base_computername_position).value
                for authorizedrow in range(1,authorized.sheet_by_index(0).nrows):     
                    if str(basename).lower() == str(authorized.sheet_by_index(0).cell(authorizedrow,authorized_computername_position).value).lower():
                        found=True
                        break      
                if not found:
                    for i in range(base.sheet_by_index(0).ncols):
                        result_page.write(next,i,str(base.sheet_by_index(0).cell(baserow,i).value))
                    next+=1
                self.progress.setValue(baserow)
                if ((baserow*100)/base.sheet_by_index(0).nrows)%1<=0.01:
                    print('%.0f'%((baserow*100)/base.sheet_by_index(0).nrows),"%")
            self.progress.setValue(base.sheet_by_index(0).nrows)
            print("100 %")
            result.save(self.resultbox.text()+".xls")
                

    def basecheckrow(self):
        if os.path.isfile(self.textboxbase.text()) and (".xls" in self.textboxbase.text() or ".xlsx" in self.textboxbase.text()):
            self.base = self.textboxbase.text()
            file = xlrd.open_workbook(self.textboxbase.text())
            self.baselabel.setText("Rows: "+str(file.sheet_by_index(0).nrows))
            self.basedrop.addItem("")
            for i in range(file.sheet_by_index(0).ncols):
                self.basedrop.addItem(str(i))
            self.basecol.setText("Select pivot column:")
            self.basedrop.show()

    def searchcheckrow(self):
        if os.path.isfile(self.textboxsearch.text()) and (".xls" in self.textboxsearch.text() or ".xlsx" in self.textboxsearch.text()):
            self.search = self.textboxsearch.text()
            file = xlrd.open_workbook(self.textboxsearch.text())
            self.searchlabel.setText("Rows: "+str(file.sheet_by_index(0).nrows))
            self.searchdrop.addItem("")
            for i in range(file.sheet_by_index(0).ncols):
                self.searchdrop.addItem(str(i))
            self.searchcol.setText("Select pivot column:")
            self.searchdrop.show()

    def isisnot(self,text):
        if str(text) == "is":
            self.lineslabel.show()
            self.linesdrop.show()
            
        else:
            self.lineslabel.hide()
            self.linesdrop.hide()
            

app = QApplication(sys.argv)
my = Window()
my.show()
app.setStyle('Fusion') 
app.exec_()
