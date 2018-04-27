# Software development2
# Natakorn Thongdee 5901012620037
# Saranporn Thitakasikorn 5901012620169
# Last Edit: 27/04/2018

import pandas as pd
from PyQt5 import QtCore, QtGui, QtWidgets
import numpy as np
import resource_rc
import sys
from threading import Thread
import dill
import pyqtgraph as pg
import random

class FileManagement:
    '''This class is used to manage file corresponding to this program.'''
    def __init__(self):
        self.workSheet = {'currentWorkSheet':None, 'excel':None, 'sheets':None}

    def readExcel(self, filePath):
        '''Read excel file.'''
        self.workSheet['excel'] = pd.ExcelFile(filePath)

    def getSheet(self):
        '''Get sheets list.'''
        self.workSheet['sheets'] = self.workSheet['excel'].sheet_names

    def readSheet(self, sheetName):
        '''Read a table from sheetname.'''
        self.workSheet['currentWorkSheet'] = sheetName.text()
        if self.workSheet['currentWorkSheet'] not in self.workSheet:
            self.workSheet[self.workSheet['currentWorkSheet']] = {'df':self.workSheet['excel'].parse(self.workSheet['currentWorkSheet']),
                              'selectedColumns':list(),
                              'selectedRows':list(),
                              'columnsValue':dict(),
                              'currentSelectedFilter':None,
                              'previousCurrentRow':None,
                              'filteredColumns':dict(),
                              'dimensions':list(),
                              'measurements':list(),
                              'dTypes':dict(),
                              'filteredDF':pd.DataFrame(),
                              'groupedDF':pd.DataFrame()
                             }

        else:
            ui.getState()

    def saveFile(self, filePath, data):
        '''Save file.'''
        with open(filePath, 'wb') as file:
            dill.dump(data, file)

    def loadFile(self, filePath):
        '''Load file.'''
        with open(filePath, 'rb') as file:
            self.workSheet = dill.load(file)

class DataOrganization(FileManagement):
    '''This class is used to mange the data that read from excel file.'''
    def __init__(self):
        super(DataOrganization, self).__init__()

    def addDictToFiltered(self):
        '''Add key of set to attribute worksheet['filteredColumns'].'''
        columns = self.workSheet[self.workSheet['currentWorkSheet']]['df'].columns
        for eachColumn in columns:
            self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][eachColumn] = set()

    def classifyDimensionMeasurement(self, df):
        '''Classify dimension and measurement.'''
        columnsType = {'dimensions':[], 'measurements':[]}

        for eachColumn, eachDataType in zip(df.columns, df.dtypes):
            if eachDataType != 'float64':
                columnsType['dimensions'].append(eachColumn)
            else:
                columnsType['measurements'].append(eachColumn)

        self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'] = columnsType['dimensions']
        self.workSheet[self.workSheet['currentWorkSheet']]['measurements'] = columnsType['measurements']
        self.getDType(df)

        if 'datetime64[ns]' in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']:
            for eachDateColum in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_tmp'] = self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum]
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_dt'] = pd.to_datetime(self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum])
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_month'] = \
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_dt'].dt.month
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_date'] = \
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_dt'].dt.day
                self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum + '_year'] = self.workSheet[self.workSheet['currentWorkSheet']]['df'][eachDateColum+ '_dt'].dt.year
                self.getColumnValue(eachDateColum + '_month')
                self.getColumnValue(eachDateColum + '_year')
                self.getColumnValue(eachDateColum + '_date')
                self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][eachDateColum + '_tmp'] = self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][eachDateColum]
                self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][eachDateColum + '_month'] = set()
                self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][eachDateColum+ '_year'] = set()
                self.workSheet[self.workSheet['currentWorkSheet']]['filteredDF'] = self.workSheet[self.workSheet['currentWorkSheet']]['df']



    def getDType(self, df):
        '''Grop data type of all column.'''
        group = df.columns.to_series().groupby(df.dtypes).groups
        self.workSheet[self.workSheet['currentWorkSheet']]['dTypes'] = {k.name : list(v) for k,v in group.items()}

    def getColumnValue(self, column):
        '''Get all rows data from each column in excel columns.'''
        values = np.array(self.workSheet[self.workSheet['currentWorkSheet']]['df'][column])
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][column] = list(np.unique(values))

    def filterByColumns(self, df, filter):
        '''Filter data frame by column name.'''
        for eachColumn in filter:
            if len(filter[eachColumn]) != 0:
                df = df[~df[eachColumn].isin(filter[eachColumn])]

        self.workSheet[self.workSheet['currentWorkSheet']]['filteredDF'] = df

    def groupData(self, df, dimensions, measurements):
        '''Group the data accroding to selected dimension and measurement.'''
        self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'] = pd.pivot_table(df, index=dimensions, values=measurements, aggfunc=np.sum)

    def multiThread(self, worker, args):
        '''Make UI responsive.'''
        for eacharg in args:
            p = Thread(target=worker, args=(eacharg,))
            p.start()

class Ui_MainWindow(DataOrganization):
    '''This class is used to manage graphical user interface.'''
    def __init__(self):
        DataOrganization.__init__(self)
        self.legend = None

    def setupUi(self, MainWindow):
        '''Construst the GUI.'''
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(1261, 838)
        MainWindow.setWindowTitle('Analytic Tool')

        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")

        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")


        # Set scroll area widget
        self.scrollArea = QtWidgets.QScrollArea(self.centralwidget)
        self.scrollArea.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea.setWidgetResizable(True)
        self.scrollArea.setObjectName("scrollArea")

        self.scrollAreaWidgetContents = QtWidgets.QWidget()
        self.scrollAreaWidgetContents.setGeometry(QtCore.QRect(0, 0, 541, 765))
        self.scrollAreaWidgetContents.setObjectName("scrollAreaWidgetContents")

        self.gridLayout_2 = QtWidgets.QGridLayout(self.scrollAreaWidgetContents)
        self.gridLayout_2.setObjectName("gridLayout_2")
        self.gridLayout_2.setColumnMinimumWidth(4, 1)

        self.rowLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.rowLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.rowLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(18, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.rowLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.rowLabel.setScaledContents(False)
        self.rowLabel.setObjectName("rowLabel")
        self.rowLabel.setText("Rows")
        self.gridLayout_2.addWidget(self.rowLabel, 4, 1)

        # Set row list widget
        self.rowListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.rowListWidget.setObjectName("rowListWidget")
        self.rowListWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.rowListWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.rowListWidget.currentItemChanged.connect(self.displayRowsFilter)
        self.rowListWidget.itemClicked.connect(self.displayRowsFilter)
        self.rowListWidget.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.gridLayout_2.addWidget(self.rowListWidget, 5, 1)

        self.columnLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.columnLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.columnLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(18, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.columnLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.columnLabel.setScaledContents(False)
        self.columnLabel.setObjectName("columnLabel")
        self.columnLabel.setText("Columns")
        self.gridLayout_2.addWidget(self.columnLabel, 2, 1, 1, 1)

        # Set column list widget
        self.columnListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.columnListWidget.setObjectName("columnListWidget")
        self.columnListWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.columnListWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.columnListWidget.currentItemChanged.connect(self.displayColumnFilter)
        self.columnListWidget.itemClicked.connect(self.displayColumnFilter)
        self.columnListWidget.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.gridLayout_2.addWidget(self.columnListWidget, 3, 1)

        self.filterLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.filterLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.filterLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(18, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.filterLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.filterLabel.setScaledContents(False)
        self.filterLabel.setObjectName("filterLabel")
        self.filterLabel.setText("Filter")
        self.gridLayout_2.addWidget(self.filterLabel, 0, 1)

        # Set filter list widget
        self.filterListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.filterListWidget.setObjectName("filterListWidget")
        self.filterListWidget.itemActivated.connect(self.toggleCheckBoxesState)
        self.gridLayout_2.addWidget(self.filterListWidget, 1, 1)

        self.sheetLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.sheetLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.sheetLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(18, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.sheetLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.sheetLabel.setScaledContents(False)
        self.sheetLabel.setObjectName("sheetLabel")
        self.sheetLabel.setText("Sheets")
        self.gridLayout_2.addWidget(self.sheetLabel, 0, 0)

        # Set sheet list widget
        self.sheetListWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.sheetListWidget.setObjectName("sheetList")
        self.sheetListWidget.itemActivated.connect(self.displayDimensionsMeasurements)
        self.gridLayout_2.addWidget(self.sheetListWidget, 1, 0)

        self.measurementLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.measurementLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.measurementLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(18, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.measurementLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.measurementLabel.setScaledContents(False)
        self.measurementLabel.setObjectName("measurementLabel")
        self.measurementLabel.setText("Measurements")
        self.gridLayout_2.addWidget(self.measurementLabel, 4, 0)

        # Set measurement list widget
        self.measurementWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.measurementWidget.setObjectName("measurementsList")
        self.measurementWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.measurementWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.gridLayout_2.addWidget(self.measurementWidget, 5, 0)

        self.dimensionLabel = QtWidgets.QLabel(self.scrollAreaWidgetContents)
        self.dimensionLabel.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.dimensionLabel.setStyleSheet("background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:0, stop:0 rgba(18, 181, 181, 255), stop:1 rgba(255, 255, 255, 255));")
        self.dimensionLabel.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.dimensionLabel.setScaledContents(False)
        self.dimensionLabel.setObjectName("dimensionLabel")
        self.dimensionLabel.setText("Dimensions")
        self.gridLayout_2.addWidget(self.dimensionLabel, 2, 0)

        # Set dimension list widget
        self.dimensionWidget = QtWidgets.QListWidget(self.scrollAreaWidgetContents)
        self.dimensionWidget.setObjectName("dimensionList")
        self.dimensionWidget.setDefaultDropAction(QtCore.Qt.MoveAction)
        self.dimensionWidget.setDragDropMode(QtWidgets.QAbstractItemView.DragDrop)
        self.gridLayout_2.addWidget(self.dimensionWidget, 3, 0)

        self.scrollArea.setWidget(self.scrollAreaWidgetContents)
        self.gridLayout.addWidget(self.scrollArea, 0, 0)

        # Set scroll area2
        self.scrollArea_2 = QtWidgets.QScrollArea(self.centralwidget)
        self.scrollArea_2.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.scrollArea_2.setWidgetResizable(True)
        self.scrollArea_2.setObjectName("scrollArea_2")
        self.scrollAreaWidgetContents_2 = QtWidgets.QWidget()
        self.scrollAreaWidgetContents_2.setGeometry(QtCore.QRect(0, 0, 342, 765))
        self.scrollAreaWidgetContents_2.setObjectName("scrollAreaWidgetContents_2")

        self.gridLayout_3 = QtWidgets.QGridLayout(self.scrollAreaWidgetContents_2)
        self.gridLayout_3.setObjectName("gridLayout_3")

        # Set Icon of continuous chart button
        self.continuousChart = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.continuousChart.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.continuousChart.setToolButtonStyle(3)  # Set text position under icon
        self.continuousChart.setText("Continuous")
        continuousChartIcon = QtGui.QIcon()
        continuousChartIcon.addPixmap(QtGui.QPixmap(":/resource/continuousChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.continuousChart.setIcon(continuousChartIcon)
        self.continuousChart.setIconSize(QtCore.QSize(170, 170))
        self.continuousChart.setObjectName("continuousChart")
        self.continuousChart.clicked.connect(self.lineChart)
        self.gridLayout_3.addWidget(self.continuousChart, 0, 0)

        # Set Icon of bar chart button
        self.barButton = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        self.barButton.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.barButton.setToolButtonStyle(3)  # Set text position under icon
        self.barButton.setText("Stack Bar")
        histrogramIcon = QtGui.QIcon()
        histrogramIcon.addPixmap(QtGui.QPixmap(":/resource/stackBarChart.svg"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.barButton.setIcon(histrogramIcon)
        self.barButton.setIconSize(QtCore.QSize(170, 170))
        self.barButton.setObjectName("stackBar")
        self.barButton.clicked.connect(self.barChart)
        self.gridLayout_3.addWidget(self.barButton, 1, 0)

        # Set Icon of scatter chart button
        self.scatterButton = QtWidgets.QToolButton(self.scrollAreaWidgetContents_2)
        scatterIcon = QtGui.QIcon()
        self.scatterButton.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.scatterButton.setToolButtonStyle(3)  # Set text position under icon
        self.scatterButton.setText("Scatter")
        scatterIcon.addPixmap(QtGui.QPixmap(":/resource/scatterChart.png"), QtGui.QIcon.Normal, QtGui.QIcon.Off)
        self.scatterButton.setIcon(scatterIcon)
        self.scatterButton.setIconSize(QtCore.QSize(170, 170))
        self.scatterButton.setObjectName("scatterButton")
        self.scatterButton.clicked.connect(self.scatterChart)
        self.gridLayout_3.addWidget(self.scatterButton, 2, 0)

        self.scrollArea_2.setWidget(self.scrollAreaWidgetContents_2)
        self.gridLayout.addWidget(self.scrollArea_2, 0, 2)
        self.tabWidget = QtWidgets.QTabWidget(self.centralwidget)

        # Set tab widget
        self.tabWidget.setBaseSize(QtCore.QSize(2000, 2000))
        self.tabWidget.setObjectName("tabWidget")

        # Set chart tab in tab widget
        self.chartTab = QtWidgets.QWidget()
        self.chartTab.setObjectName("chartTab")
        self.plotLayout = QtWidgets.QVBoxLayout()
        self.plotWidget = pg.PlotWidget(name='Plot')
        self.plotWidget.setBackground('w')
        self.plotLayout.addWidget(self.plotWidget)
        self.chartTab.setLayout(self.plotLayout)
        self.tabWidget.addTab(self.chartTab, "")
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.chartTab), "Chart")

        # Set tabel tab in tab widget
        self.tableTab = QtWidgets.QWidget()
        self.tableTab.setObjectName("tableTab")
        self.tableLayout = QtWidgets.QVBoxLayout()
        self.tableWidget = pg.TableWidget()
        self.tableLayout.addWidget(self.tableWidget)
        self.tableTab.setLayout(self.tableLayout)
        self.tabWidget.addTab(self.tableTab, "")

        self.gridLayout.addWidget(self.tabWidget, 0, 1)
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.chartTab), "Chart")
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tableTab), "Table")

        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1261, 26))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)

        # Set status bar
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        # Add file to statusbar
        self.menuFile = QtWidgets.QMenu(self.menubar)
        self.menuFile.setObjectName("menuFile")
        self.menuFile.setTitle("File")

        # Add open action to file
        self.actionOpen = QtWidgets.QAction(MainWindow)
        self.actionOpen.setObjectName("actionImport")
        self.actionOpen.setText("Open")
        self.actionOpen.setShortcut("Ctrl+O")
        self.actionOpen.triggered.connect(self.openFileDialog)

        # Add import action to file
        self.actionImport = QtWidgets.QAction(MainWindow)
        self.actionImport.setObjectName("actionImport")
        self.actionImport.setText("Import")
        self.actionImport.setShortcut("Ctrl+I")
        self.actionImport.triggered.connect(self.openFileDialog)

        # Add save action to file
        self.actionSave = QtWidgets.QAction(MainWindow)
        self.actionSave.setObjectName("actionSave")
        self.actionSave.setText("Save")
        self.actionSave.setShortcut("Ctrl+S")
        self.actionSave.triggered.connect(self.saveFileDialog)

        self.menuFile.addAction(self.actionOpen)
        self.menuFile.addAction(self.actionImport)
        self.menuFile.addAction(self.actionSave)
        self.menubar.addAction(self.menuFile.menuAction())

        self.tabWidget.setCurrentIndex(0)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def addListObject(self, iterableObject, widget):
        '''Add list item to QListWidget.'''
        for eachObject in iterableObject:
            self.item = QtWidgets.QListWidgetItem(eachObject)
            widget.addItem(self.item)

    def openFileDialog(self):
        '''Get loaded file name.'''
        directory = QtWidgets.QFileDialog.getOpenFileName()
        directory = str(directory[0])
        if directory != '':
            fileName = directory.split('/')[-1]
            extension = fileName.split('.')[-1]
            self.sheetListWidget.clear()
            self.dimensionWidget.clear()
            self.measurementWidget.clear()
            self.rowListWidget.clear()
            self.columnListWidget.clear()
            self.filterListWidget.clear()

            if extension == 'xlsx':
                self.readExcel(fileName)
                self.getSheet()
                self.addListObject(self.workSheet['sheets'], self.sheetListWidget)

            elif extension == 'pkl':
                self.loadFile(fileName)
                self.addListObject(self.workSheet['sheets'], self.sheetListWidget)
                self.addListObject(self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'], self.dimensionWidget)
                self.addListObject(self.workSheet[self.workSheet['currentWorkSheet']]['measurements'], self.measurementWidget)
                self.addListObject(self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns'], self.columnListWidget)
                self.addListObject(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'], self.rowListWidget)

                if self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] is not None:
                    checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'], self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'])
                    for eachCheckBox in checkBoxes:
                        self.filterListWidget.addItem(eachCheckBox)

    def saveFileDialog(self):
        '''Get saved file name from dialog window.'''
        fileName = QtWidgets.QFileDialog.getSaveFileName()
        fileName = fileName[0]
        self.getState()

        if fileName != '':
            fileName = fileName.split('/')[-1] + '.pkl'
            self.saveFile(fileName, self.workSheet)

    def displayDimensionsMeasurements(self, sheet):
        '''Display dimensions and measurements.'''
        self.readSheet(sheet)
        self.dimensionWidget.clear()
        self.measurementWidget.clear()
        self.columnListWidget.clear()
        self.rowListWidget.clear()
        self.filterListWidget.clear()
        self.multiThread(self.getColumnValue, list(self.workSheet[self.workSheet['currentWorkSheet']]['df'].columns))
        self.classifyDimensionMeasurement(self.workSheet[self.workSheet['currentWorkSheet']]["df"])
        self.addListObject(self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'], self.dimensionWidget)
        self.addListObject(self.workSheet[self.workSheet['currentWorkSheet']]['measurements'], self.measurementWidget)
        self.addDictToFiltered()

    def getCheckBoxes(self, filters, column):
        '''Instantiate checkboxes.'''
        items = list()
        filters = list(map(str, filters))

        if 'datetime64[ns]' in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']:
            if '_month' in column:
                filters.sort(key=int)

        for eachFilter in filters:
            item = QtWidgets.QListWidgetItem()
            item.setText(eachFilter)
            item.setFlags(item.flags() | QtCore.Qt.ItemIsUserCheckable)

            if item.text() not in self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][column]:
                item.setCheckState(QtCore.Qt.Checked)
            else:
                item.setCheckState(QtCore.Qt.Unchecked)
            items.append(item)
        return items

    def toggleCheckBoxesState(self):
        '''Toggle all check boxes state.'''
        selectedItem = self.filterListWidget.currentItem().text()

        if selectedItem == 'Uncheck all filters':
            for eachRow in range(self.filterListWidget.count()):
                item = self.filterListWidget.item(eachRow)
                if item.text() != 'Uncheck all filters' and item.text() != 'Check all filters':
                    if item.checkState() == 2:
                        item.setCheckState(QtCore.Qt.Unchecked)

        elif selectedItem == 'Check all filters':
            for eachRow in range(self.filterListWidget.count()):
                item = self.filterListWidget.item(eachRow)
                if item.text() != 'Uncheck all filters' and item.text() != 'Check all filters':
                    if item.checkState() == 0:
                        item.setCheckState(QtCore.Qt.Checked)

    def getCheckBoxesState(self, column):
        '''Get checkboxes state from filter box.'''
        for eachRow in range(self.filterListWidget.count()):
            item = self.filterListWidget.item(eachRow)
            if item.text() != 'Uncheck all filters' and item.text() != 'Check all filters':
                if item.checkState() != 2:
                    self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][column].add(item.text())
                elif item.text() in self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][column]:
                    self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][column].remove(item.text())

    def deleteFilteredColumns(self, column):
        '''Delete filter.'''
        for eachFilter in list(self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][column]):
            if eachFilter in self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][column]:
                self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'][column].remove(eachFilter)


    def defaultFormat(self):
        '''Add checkboxes as default format of date time type.'''
        self.filterListWidget.clear()
        self.workSheet[self.workSheet['currentWorkSheet']]['df'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = \
        self.workSheet[self.workSheet['currentWorkSheet']]['df'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_tmp']
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = \
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_tmp']
        checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']],
                                        self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'])
        for eachCheckBox in checkBoxes:
            self.filterListWidget.addItem(eachCheckBox)
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Uncheck all filters'))
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Check all filters'))


    def dailyFormat(self):
        '''Add checkboxes as default format of daily.'''
        self.filterListWidget.clear()
        self.workSheet[self.workSheet['currentWorkSheet']]['df'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = self.workSheet[self.workSheet['currentWorkSheet']]['df'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_date']
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = \
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_date']
        checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_date'], self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_date')
        for eachCheckBox in checkBoxes:
            self.filterListWidget.addItem(eachCheckBox)
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Uncheck all filters'))
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Check all filters'))


    def monthlyFormat(self):
        '''Add checkboxes as default format of monthly.'''
        self.filterListWidget.clear()
        self.workSheet[self.workSheet['currentWorkSheet']]['df'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = self.workSheet[self.workSheet['currentWorkSheet']]['df'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_month']
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = \
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_month']
        checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_month'], self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_month')
        for eachCheckBox in checkBoxes:
            self.filterListWidget.addItem(eachCheckBox)
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Uncheck all filters'))
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Check all filters'))

    def yearlyFormat(self):
        '''Add checkboxes as default format of yearly.'''
        self.filterListWidget.clear()
        self.workSheet[self.workSheet['currentWorkSheet']]['df'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = \
        self.workSheet[self.workSheet['currentWorkSheet']]['df'][
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_year']
        self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter']] = self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] + '_year']
        checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][
                                            self.workSheet[self.workSheet['currentWorkSheet']][
                                                'currentSelectedFilter'] + '_year'],
                                        self.workSheet[self.workSheet['currentWorkSheet']][
                                            'currentSelectedFilter'] + '_year')
        for eachCheckBox in checkBoxes:
            self.filterListWidget.addItem(eachCheckBox)
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Uncheck all filters'))
        self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Check all filters'))

    def addAction(self, listWidget):
        '''Add action to date time column.'''
        self.actionDefault = QtWidgets.QAction("As Default", listWidget)
        self.actionDefault.triggered.connect(self.defaultFormat)
        self.actionMonthly = QtWidgets.QAction("Monthly", listWidget)
        self.actionMonthly.triggered.connect(self.monthlyFormat)
        self.actionYearly = QtWidgets.QAction("Yearly", listWidget)
        self.actionYearly.triggered.connect(self.yearlyFormat)
        self.actionDaily = QtWidgets.QAction("Daily", listWidget)
        self.actionDaily.triggered.connect(self.dailyFormat)
        listWidget.addAction(self.actionDefault)
        listWidget.addAction(self.actionDaily)
        listWidget.addAction(self.actionMonthly)
        listWidget.addAction(self.actionYearly)


    def deleteAction(self, listWidget):
        '''Delete action to date time column.'''
        listWidget.removeAction(self.actionDefault)
        listWidget.removeAction(self.actionMonthly)
        listWidget.removeAction(self.actionYearly)
        listWidget.removeAction(self.actionDaily)

    def displayColumnFilter(self):
        '''Display filter box when current item in columns box is changed.'''
        if self.columnListWidget.currentRow() != -1:
            column = self.columnListWidget.item(self.columnListWidget.currentRow()).text()
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] = column

            if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in \
                    self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                self.deleteAction(self.columnListWidget)
            if column in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                self.addAction(self.columnListWidget)

            if column != self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow']:
                self.getState()
                if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] is not None:
                    if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'] or \
                                    self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in self.workSheet[self.workSheet['currentWorkSheet']]['measurements']:
                        self.deleteFilteredColumns(self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'])
                    else:
                        self.getCheckBoxesState(self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'])

            self.filterListWidget.clear()
            self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] = column
            checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][column], column)

            for eachCheckBox in checkBoxes:
                self.filterListWidget.addItem(eachCheckBox)
            self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Uncheck all filters'))
            self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Check all filters'))
        else:

            if 'datetime64[ns]' in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']:
                if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in \
                        self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                    self.deleteAction(self.columnListWidget)

            self.filterListWidget.clear()
            self.addDictToFiltered()
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] = None
            self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] = None


    def displayRowsFilter(self):
        '''Display filter box when current item in rows box is changed.'''
        if self.rowListWidget.currentRow() != -1:
            column = self.rowListWidget.item(self.rowListWidget.currentRow()).text()
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] = column

            if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in \
                    self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                self.deleteAction(self.rowListWidget)
            if column in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                self.addAction(self.rowListWidget)


            if column != self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow']:
                self.getState()
                if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] is not None:
                    if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'] or \
                                    self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in self.workSheet[self.workSheet['currentWorkSheet']]['measurements']:

                        self.deleteFilteredColumns(self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'])
                    else:
                        self.getCheckBoxesState(self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'])

            self.filterListWidget.clear()
            self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] = column
            checkBoxes = self.getCheckBoxes(self.workSheet[self.workSheet['currentWorkSheet']]['columnsValue'][column], column)

            for eachCheckBox in checkBoxes:
                self.filterListWidget.addItem(eachCheckBox)
            self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Uncheck all filters'))
            self.filterListWidget.addItem(QtWidgets.QListWidgetItem('Check all filters'))

        else:
            if 'datetime64[ns]' in self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']:
                if self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] in \
                        self.workSheet[self.workSheet['currentWorkSheet']]['dTypes']['datetime64[ns]']:
                    self.deleteAction(self.rowListWidget)

            self.filterListWidget.clear()
            self.addDictToFiltered()
            self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'] = None
            self.workSheet[self.workSheet['currentWorkSheet']]['previousCurrentRow'] = None

    def getState(self):
        '''Save item from in dimensions measurements columns and rows in workSheet attribute. '''
        self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'] = list()
        for eachRow in range(self.dimensionWidget.count()):
            self.workSheet[self.workSheet['currentWorkSheet']]['dimensions'].append(self.dimensionWidget.item(eachRow).text())

        self.workSheet[self.workSheet['currentWorkSheet']]['measurements'] = list()
        for eachRow in range(self.measurementWidget.count()):
            self.workSheet[self.workSheet['currentWorkSheet']]['measurements'].append(self.measurementWidget.item(eachRow).text())

        self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns'] = list()
        for eachRow in range(self.columnListWidget.count()):
            self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns'].append(self.columnListWidget.item(eachRow).text())

        self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'] = list()
        for eachRow in range(self.rowListWidget.count()):
            self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'].append(self.rowListWidget.item(eachRow).text())

    def plotTable(self, dimensionAxis):
        '''Create table in table tab.'''
        data = {eachIndex:list(self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].loc[eachIndex, :]) for eachIndex in dimensionAxis}
        self.tableWidget.setData(data)
        self.tableWidget.setHorizontalHeaderLabels(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'])


    def randomDistinctColor(self, numbers):
        '''Random n distinct color.'''
        colors = []
        red = int(random.random() * 256)
        green = int(random.random() * 256)
        blue = int(random.random() * 256)
        step = 256 / numbers
        for i in range(numbers):
            red += step
            green += step
            blue += step
            red = int(red) % 256
            green = int(green) % 256
            blue = int(blue) % 256
            colors.append((red, green, blue))
        return colors

    def barChart(self):
        '''Plot bar chart in chart tab.'''
        self.getState()
        if (len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']) >= 1 and len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns']) != 0):
            self.getCheckBoxesState(self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'])
            self.filterByColumns(self.workSheet[self.workSheet['currentWorkSheet']]['df'], self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'])
            self.groupData(self.workSheet[self.workSheet['currentWorkSheet']]['filteredDF'], self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns'], self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'])
            dimensionsAxis = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].index.values
            self.plotTable(dimensionsAxis)
            self.plotWidget.showGrid(x=False, y=False)
            measurementAxis = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].loc[:, self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']]
            dimensionsAxis = list(map(str, self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].index.values))
            dimensionsAxis = list(enumerate(dimensionsAxis))
            self.plotWidget.clear()
            self.plotWidget.clearPlots()
            colors = self.randomDistinctColor(len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']))
            if self.legend != None:
                self.legend.scene().removeItem(self.legend)
            self.legend = self.plotWidget.addLegend()
            yMin = 0
            yMax = 0
            Xmin = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].shape[0]
            if Xmin > 20:
                Xmin = 20
            for eachMeasurement, color in zip(measurementAxis, colors):
                eachX = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].shape[0]
                eachY = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'][eachMeasurement]
                maxY = np.max(eachY)
                minY = np.min(eachY)
                if maxY >  yMax:
                    yMax = maxY
                if minY < yMin:
                    yMin = minY
                barChart = pg.BarGraphItem(x=np.arange(eachX), height=eachY, width=0.5, brush=color)
                self.plotWidget.addItem(barChart)
                self.plotWidget.plot(name=eachMeasurement, pen=color)

            self.plotWidget.setLimits(yMin= yMin, yMax=yMax)
            self.plotWidget.setXRange(0, Xmin)
            self.plotWidget.getAxis('bottom').setTicks([dimensionsAxis])

    def lineChart(self):
        '''Plot line chart in chart tab.'''
        self.getState()
        if (len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']) >= 1 and len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns']) != 0):
            self.getCheckBoxesState(self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'])
            self.filterByColumns(self.workSheet[self.workSheet['currentWorkSheet']]['df'], self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'])
            self.groupData(self.workSheet[self.workSheet['currentWorkSheet']]['filteredDF'], self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns'],
                           self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'])
            dimensionsAxis = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].index.values
            self.plotTable(dimensionsAxis)
            self.plotWidget.showGrid(x=True, y=True)
            measurementAxis = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].loc[:, self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']]
            dimensionsAxis = list(map(str, self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].index.values))
            dimensionsAxis = list(enumerate(dimensionsAxis))
            self.plotWidget.clear()
            self.plotWidget.clearPlots()
            if self.legend != None:
                self.legend.scene().removeItem(self.legend)
            self.legend = self.plotWidget.addLegend()
            colors = self.randomDistinctColor(len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']))
            yMin = 0
            yMax = 0
            Xmin = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].shape[0]
            if Xmin > 20:
                Xmin = 20
            for eachMeasurement, color in zip(measurementAxis, colors):
                x = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].shape[0]
                y = list(self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'][eachMeasurement])
                maxY = np.max(y)
                minY = np.min(y)
                if maxY > yMax:
                    yMax = maxY
                if minY < yMin:
                    yMin = minY
                self.plotWidget.plot(x=np.arange(x), y=y, pen=color, symbol='x', symbolPen=color, name=eachMeasurement)

            self.plotWidget.setLimits(yMin=yMin, yMax=yMax)
            self.plotWidget.setXRange(0, Xmin)

            self.plotWidget.getAxis('bottom').setTicks([dimensionsAxis])

    def scatterChart(self):
        '''Scatter chart in chart tab.'''
        self.getState()
        if (len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']) >= 1 and len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns']) != 0):
            self.getCheckBoxesState(self.workSheet[self.workSheet['currentWorkSheet']]['currentSelectedFilter'])
            self.filterByColumns(self.workSheet[self.workSheet['currentWorkSheet']]['df'], self.workSheet[self.workSheet['currentWorkSheet']]['filteredColumns'])
            self.groupData(self.workSheet[self.workSheet['currentWorkSheet']]['filteredDF'], self.workSheet[self.workSheet['currentWorkSheet']]['selectedColumns'], self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows'])
            dimensionsAxis = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].index.values
            self.plotTable(dimensionsAxis)
            self.plotWidget.showGrid(x=False, y=False)
            measurementAxis = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].loc[:, self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']]
            dimensionsAxis = list(map(str, self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].index.values))
            dimensionsAxis = list(enumerate(dimensionsAxis))
            self.plotWidget.clear()
            self.plotWidget.clearPlots()
            if self.legend != None:
                self.legend.scene().removeItem(self.legend)
            self.legend = self.plotWidget.addLegend()
            colors = self.randomDistinctColor(len(self.workSheet[self.workSheet['currentWorkSheet']]['selectedRows']))
            yMin = 0
            yMax = 0
            Xmin = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].shape[0]
            if Xmin > 20:
                Xmin = 20

            for eachMeasurement, color in zip(measurementAxis, colors):
                x = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'].shape[0]
                y = self.workSheet[self.workSheet['currentWorkSheet']]['groupedDF'][eachMeasurement]

                maxY = np.max(y)
                minY = np.min(y)
                if maxY > yMax:
                    yMax = maxY
                if minY < yMin:
                    yMin = minY
                scatter = pg.ScatterPlotItem(size=10, pen=color, name=eachMeasurement)
                scatter.addPoints(x=np.arange(x), y=y)
                self.plotWidget.addItem(scatter)
                self.plotWidget.plot(name=eachMeasurement, pen=color)

            self.plotWidget.setLimits(yMin=yMin, yMax=yMax)
            self.plotWidget.setXRange(0, Xmin)
            self.plotWidget.getAxis('bottom').setTicks([dimensionsAxis])


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = Ui_MainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())
