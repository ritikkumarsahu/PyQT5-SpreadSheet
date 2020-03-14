import sys
import os
import csv
import xlrd
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QTableWidget, QTabWidget, QVBoxLayout, QWidget, QApplication, QMainWindow, QTableWidgetItem, QFileDialog
from PyQt5.QtWidgets import qApp, QAction, QMessageBox
from json import dumps
from os import remove
class MyTable(QTableWidget):
    ''' This class creates a QTableWidget Object'''
    def __init__(self, r, c):
        ''' initialise the table with rows and columns

        :param r: number of rows
        :param c: number of columns

        :return: None
        '''
        super().__init__(r, c)
        self.check_change = True
        self.ncols = 26
        self.nrows = 26
        self.init_ui()

    def init_ui(self):
        ''' initialise User Interface

        :return: None
        '''
        self.cellChanged.connect(self.c_current)
        self.show()

    def c_current(self):
        ''' Called whenever cell value changes

        Functionality: checks the bad value for a cell and pops dialog box

        :return: None
        '''
        if self.check_change:
            row = self.currentRow()
            col = self.currentColumn()
            self.ncols = max(col,self.ncols)
            self.nrows = max(row,self.nrows)
            value = self.item(row, col)
            Value = value.text()
            if Value.replace('.','').replace('-','').replace('+','').isnumeric() == False:
                self.messageBox([row,col],'cell')

    def messageBox(self,bad_values,flag):
        ''' initialise the QMessageBox object for warning message

        :param bad_values: list of (rows,cols) containing bad values
        :param flag: Flags the condition for bad values

        :return: None
        '''
        msg = QMessageBox()
        msg.setWindowTitle("Warning!")
        if flag == 'cell':
            msg.setText('cell ({0:d}, {1:d}) Contains bad value'.format(bad_values[0]+1,bad_values[1]+1))
        else:
            temp_str = ''
            for row,col in bad_values:
                temp_str += '({0:d}, {1:d}), '.format(row+1,col+1)
            temp_str = 'cell {0:s} Contain bad values'.format(temp_str)
            if flag == 'DuplicateID':
                if len(bad_values) > 0:
                    temp_str += '\nColumn(ID) contains Duplicate Values'
                else:
                    temp_str = 'Column(ID) contains Duplicate Values'
            msg.setText(temp_str)
        x = msg.exec_()

class Sheet():
    ''' Class for creating table '''
    def __init__(self):
        ''' Constructor for instantiating the table object

        :return: None
        '''
        super().__init__()
        self.form_widget = MyTable(26, 26)
        self.col_headers = ['ID', 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']
    def setColumnHeaders(self):
        ''' Sets the Headers for the table Object

        :return: None
        '''
        self.form_widget.setHorizontalHeaderLabels(self.col_headers)
class Main(QMainWindow):
    ''' Main class inherits from QMainWindow'''
    def __init__(self):
        ''' initialises the class with essential variables and objects

        :return: None
        '''
        super().__init__()
        self.setWindowTitle('FOSSEE SpreadSheet')
        self.resize(800, 600)
        self.work_books = 1
        self.Tab_widget = QTabWidget()
        self.vBox = QVBoxLayout()
        Central_widget = QWidget()
        Central_widget.setLayout(self.vBox)
        self.work_booksList = []
        self.work_booksList.append(Sheet())
        self.work_booksList[0].setColumnHeaders()
        self.vBox.addWidget(self.Tab_widget)
        self.setTab()
        self.setCentralWidget(Central_widget)
        self.show()
        self.setMenu()
    def setTab(self,index=0,name="Untitled"):
        ''' Add Tabs to QTabWidget for number

        :param index: current index of Tab in QTabWidget
        :param name: sets the title of tab according to csv file name

        :return: None
        '''
        self.Tab_widget.addTab(self.work_booksList[index].form_widget,name)

    def setMenu(self):
        ''' Sets the menu, toolbar, user interface

        :return: None
        '''
        #setup actions
        quit_action = QAction('Quit', self)
        load_action  = QAction('Load inputs', self)
        validate_action = QAction('Validate', self)
        submit_action = QAction('Submit', self)
        #setup triggered action
        load_action.triggered.connect(self.open_sheet)
        validate_action.triggered.connect(self.validate_sheet)
        submit_action.triggered.connect(self.submit_sheet)
        quit_action.triggered.connect(self.quit_app)
        #setup Toolbar
        toolbar = QMainWindow.addToolBar(self,'Toolbar')
        toolbar.addAction(load_action)
        toolbar.addAction(validate_action)
        toolbar.addAction(submit_action)
        toolbar.addAction(quit_action)

    def xlsxToCsv(self,path):
        '''Converts the xlsx file to csv file format for each workbook

        :param path: Directory path for the file

        :return: None
        '''
        wb = xlrd.open_workbook(path,on_demand=True)
        self.work_books = wb.nsheets
        for i in range(self.work_books):
            sh = wb.sheet_by_index(i)
            your_csv_file = open('your_csv_file'+str(i)+'.csv', 'w')
            wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL,lineterminator='\n')

            for rownum in range(sh.nrows):
                wr.writerow(sh.row_values(rownum))                    
            your_csv_file.close()
        return [['your_csv_file'+str(i)+'.csv', wb.sheet_names()[i] ] for i in range(self.work_books)]

    def open_sheet(self):
        ''' calls file browser to open the csv,xlsx file

        :return: workbooklist - contains class sheet object list
                  paths - list of path for each csv file
        '''
        paths = ['']
        self.work_booksList[0].form_widget.check_change = False
        path = QFileDialog.getOpenFileName(self, 'Open CSV/XLS', os.getenv('HOME'), 'CSV(*.csv *.xlsx)')
        if path[0] == '':
            return
        paths[0] = path[0]
        sheet_names = [path[0].split('/')[-1].split('.')[0]]
        if path[0].split('/')[-1].split('.')[-1] == 'xlsx':
            paths.clear()
            paths,sheet_names = list(zip(*self.xlsxToCsv(path[0])))
        else:
            self.work_books = 1
                #changes needed to open multiple WB
        self.work_booksList[0].form_widget.deleteLater()
        self.Tab_widget.removeTab(0)
        self.work_booksList.clear()
        for g in range(self.work_books):
            self.work_booksList.append(Sheet())
            self.setTab(g,sheet_names[g])
            self.work_booksList[g].form_widget.check_change = False
        def Sheets(Workbooks,paths):
            ''' repopulates the each table with each csv file 

            :param Workbooks: list of class sheet objects
            :param paths: list of paths to the each csv files
                        
            :return: None
            '''
            for Workbook,path in zip(Workbooks,paths):
                with open(path, newline='',encoding='utf-8',errors='ignore') as csv_file:
                    Workbook.form_widget.setRowCount(0)
                    #Workbook.form_widget.setColumnCount(10)
                    my_file = csv.reader(csv_file, dialect='excel')
                    fields = next(my_file)
                    Workbook.col_headers = list(filter(lambda x: x != "", fields))
                    if len(fields) < 26:
                        Workbook.form_widget.ncols = len(fields)
                        Workbook.form_widget.setColumnCount(26)
                    Workbook.setColumnHeaders()
                    for row_data in my_file:
                        row = Workbook.form_widget.rowCount()
                        Workbook.form_widget.insertRow(row)
                        for column, stuff in enumerate(row_data):
                            item = QTableWidgetItem(stuff)
                            Workbook.form_widget.setItem(row, column, item)
                    if Workbook.form_widget.rowCount() < 26:
                        Workbook.form_widget.nrows = Workbook.form_widget.rowCount()
                        Workbook.form_widget.setRowCount(26)
                Workbook.form_widget.check_change = True
            for w in range(self.work_books):
                try:
                   remove('your_csv_file{0:d}.csv'.format(w))
                except:
                      pass

        return Sheets(self.work_booksList,paths)
    def validate_sheet(self):
        ''' validates the workbook table if contains bad values calls message box

        :return: None
        '''
        Bad_val = []
        ID_col = []
        flag='NoDuplicateID'
        cur_workbook = self.Tab_widget.currentIndex()
        ncols = self.work_booksList[cur_workbook].form_widget.ncols 
        nrows = self.work_booksList[cur_workbook].form_widget.nrows 
        for i in range(0,ncols):
            for j in range(0,nrows):
                if self.work_booksList[cur_workbook].form_widget.item(j, i) is not None:
                    if (self.work_booksList[cur_workbook].form_widget.item(j, i).text()).replace('.','').isnumeric() == False:
                        Bad_val.append([i,j])
                    elif i == 0:
                        ID_col.append(self.work_booksList[cur_workbook].form_widget.item(j, i).text())
        #checking duplicate IDs
        if len(ID_col) > len(set(ID_col)):
            flag = 'DuplicateID'
        if len(Bad_val) > 0 or flag == 'DuplicateID':
            self.work_booksList[cur_workbook].form_widget.messageBox(Bad_val,flag)
       
    def submit_sheet(self):
        ''' creates the text file of each row of workbook containg the data of that row

        :return: None
        '''
        cur_workbookIndex = self.Tab_widget.currentIndex()
        cur_workbookTitle = self.Tab_widget.tabText(cur_workbookIndex)
        ncols = self.work_booksList[cur_workbookIndex].form_widget.ncols 
        nrows = self.work_booksList[cur_workbookIndex].form_widget.nrows
        Dictonary = {}
        for i in range(0,nrows):
            for j in range(0,ncols):
                if self.work_booksList[cur_workbookIndex].form_widget.item(i, j) is not None:
                   Dictonary[self.work_booksList[cur_workbookIndex].form_widget.horizontalHeaderItem(j).text()] = self.work_booksList[cur_workbookIndex].form_widget.item(i, j).text()
            if self.work_booksList[cur_workbookIndex].form_widget.item(i, 0) is not None:
                with open('{0:s}_{1:s}.txt'.format(cur_workbookTitle,self.work_booksList[cur_workbookIndex].form_widget.item(i, 0).text()), 'w') as file:
                     file.write(dumps(Dictonary).replace('{','').replace('}',''))
                Dictonary.clear()

    def quit_app(self):
        ''' Quits the window and closes the app

        :return: None
        '''
        qApp.quit()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = Main()
    sys.exit(app.exec_())

