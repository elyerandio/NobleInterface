#***********************************
# Program Name : NobleBrowser.py
# Date Written : August 29, 2013
# Description  : a program to browse time-in/time-out data in Docomo's Noble System
# Author       : Eleazer L. Erandio
#************************************

import sys
import os
from PySide.QtCore import *
from PySide.QtGui import *
import pyodbc
from datetime import *
import re
import ConfigParser


class NobleBrowserMainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(NobleBrowserMainWindow, self).__init__(parent)
        form = NobleBrowserForm()
        self.setCentralWidget(form)
        self.setWindowTitle('Noble/Orisoft Database Browser')
        #form.show()

class NobleBrowserForm(QDialog):

    def __init__(self, parent=None):
        super(NobleBrowserForm, self).__init__(parent)

        fromLabel = QLabel('From')
        fromLabel.setAlignment(Qt.AlignHCenter)
        toLabel = QLabel('To')
        toLabel.setAlignment(Qt.AlignHCenter)
        workgroupLabel = QLabel('WorkGroup')
        employeeLabel = QLabel('Employee')
        dateLabel = QLabel('Date')
        self.workgroupComboFrom = QComboBox()
        self.workgroupComboTo = QComboBox()
        self.employeeComboFrom = QComboBox()
        self.employeeComboTo = QComboBox()
        self.dateEditFrom = QDateTimeEdit(QDate.currentDate())
        self.dateEditFrom.setCalendarPopup(True)
        self.dateEditFrom.setMaximumDate(QDate.currentDate())
        self.dateEditTo = QDateTimeEdit(QDate.currentDate())
        self.dateEditTo.setCalendarPopup(True)
        self.rdoBrowseNoble = QRadioButton('Browse Noble')
        self.rdoBrowseNoble.setChecked(True)
        self.rdoBrowseOrisoft = QRadioButton('Browse Orisoft')
        self.labelStatus = QLabel('')
        #self.dateEditTo.setDateRange(self.dateEditFrom.date(), QDate.currentDate())
        #self.dateEditTo.setMinimumDate(self.dateEditFrom.date())

        self.viewLogButton = QPushButton('View Log')
        self.processButton = QPushButton('Process')
        self.cancelButton = QPushButton('Cancel')
        buttonLayout = QHBoxLayout()
        buttonLayout.addSpacing(2)
        buttonLayout.addWidget(self.viewLogButton)
        buttonLayout.addWidget(self.processButton)
        buttonLayout.addWidget(self.cancelButton)

        layout = QGridLayout()
        layout.addWidget(fromLabel, 0, 1)
        layout.addWidget(toLabel, 0, 2)
        layout.addWidget(workgroupLabel, 1, 0)
        layout.addWidget(self.workgroupComboFrom, 1, 1)
        layout.addWidget(self.workgroupComboTo, 1, 2)
        layout.addWidget(employeeLabel, 2, 0)
        layout.addWidget(self.employeeComboFrom, 2, 1)
        layout.addWidget(self.employeeComboTo, 2, 2)
        layout.addWidget(dateLabel, 3, 0)
        layout.addWidget(self.dateEditFrom, 3, 1)
        layout.addWidget(self.dateEditTo, 3, 2)
        #layout.addWidget(QLabel(''), 5,0)
        layout.addWidget(self.rdoBrowseNoble, 4, 1)
        layout.addWidget(self.rdoBrowseOrisoft, 4, 2)
        layout.addLayout(buttonLayout, 6, 1, 1, 2)
        layout.addWidget(self.labelStatus, 7, 0, 1, 3)
        self.initGui()

        #self.connect(self.workgroupComboFrom, SIGNAL('currentIndexChanged(int)'), self.populateBadge)
        #self.connect(self.workgroupComboTo, SIGNAL('currentIndexChanged(int)'), self.populateBadge)
        self.workgroupComboFrom.currentIndexChanged.connect(self.setWorkGroupTo)
        self.workgroupComboTo.currentIndexChanged.connect(self.setWorkGroupFrom)
        self.employeeComboTo.currentIndexChanged.connect(self.setEmployeeFrom)
        self.dateEditFrom.dateChanged.connect(self.setDateTo)
        self.dateEditTo.dateChanged.connect(self.setDateFrom)
        self.connect(self.employeeComboFrom, SIGNAL('currentIndexChanged(int)'), self.setEmployeeTo)
        self.connect(self.processButton, SIGNAL('clicked()'), self.process)
        self.connect(self.cancelButton, SIGNAL('clicked()'), self.canceled)
        self.viewLogButton.clicked.connect(self.viewLog)

        self.setLayout(layout)
        self.setWindowTitle('Noble/Orisoft Database Browser')


    def initGui(self):
        cur = connOriTMS.cursor()
        cur.execute('Select WORK_GROUP_CODE, WORK_GROUP_DESCRIPTION from WORK_GROUP order by WORK_GROUP_CODE')
        i = 0
        for rec in cur:
            #self.workgroupComboFrom.addItem('%6s -- %s' % (rec[0], rec[1]))
            #self.workgroupComboTo.addItem('%6s -- %s' % (rec[0], rec[1]))
            self.workgroupComboFrom.addItem(rec[0])
            self.workgroupComboTo.addItem(rec[0])
            i += 1

        self.workgroupComboTo.setCurrentIndex(i-1)
        cur.execute("Select BADGE_NO, EMPLOYEE_NAME from EMPLOYEE_BADGE where CATEGORY_CODE = 'CAGT' order by EMPLOYEE_NAME")
        i = 0
        for rec in cur:
            self.employeeComboFrom.addItem(rec[1])
            self.employeeComboFrom.setItemData(i, rec[0])
            self.employeeComboTo.addItem(rec[1])
            self.employeeComboTo.setItemData(i, rec[0])
            i += 1

        cur.close()

    def setWorkGroupTo(self):
        if self.workgroupComboTo.currentIndex() < self.workgroupComboFrom.currentIndex():
            self.workgroupComboTo.setCurrentIndex(self.workgroupComboFrom.currentIndex())

        self.populateBadge()

    def setWorkGroupFrom(self):
        if self.workgroupComboTo.currentIndex() < self.workgroupComboFrom.currentIndex():
            self.workgroupComboFrom.setCurrentIndex(self.workgroupComboTo.currentIndex())

        self.populateBadge()

    def populateBadge(self):
        self.employeeComboFrom.clear()
        self.employeeComboTo.clear()
        sqlStr = "Select BADGE_NO, EMPLOYEE_NAME from EMPLOYEE_BADGE where CATEGORY_CODE = 'CAGT' and " \
                "WORK_GROUP_CODE in " \
                "(Select WORK_GROUP_CODE from WORK_GROUP where WORK_GROUP_CODE >= '%s' and WORK_GROUP_CODE <= '%s') order by EMPLOYEE_NAME" % \
                (self.workgroupComboFrom.currentText(), self.workgroupComboTo.currentText())

        print sqlStr
        cursor = connOriTMS.cursor()
        cursor.execute(sqlStr)
        i = 0
        for rec in cursor:
            self.employeeComboFrom.addItem(rec[1])
            self.employeeComboFrom.setItemData(i, rec[0])
            self.employeeComboTo.addItem(rec[1])
            self.employeeComboTo.setItemData(i, rec[0])
            i += 1

        cursor.close()

    def setEmployeeFrom(self):
        # don't allow employeeComboTo to be less than employeeComboFrom
        # set employeeComboFrom to employeeComboTo
        if self.employeeComboTo.currentIndex() < self.employeeComboFrom.currentIndex():
            self.employeeComboFrom.setCurrentIndex(self.employeeComboTo.currentIndex())

    def setEmployeeTo(self):
        # set employeeComboTo to employeeComboFrom
        self.employeeComboTo.setCurrentIndex(self.employeeComboFrom.currentIndex())

    def setDateFrom(self):
        # don't allow dateEditTo.date to be less than dateEditFrom.date
        # set dateEditFrom.date to dateEditTo.date
        if self.dateEditTo.date() < self.dateEditFrom.date():
            self.dateEditFrom.setDate(self.dateEditTo.date())

        if self.dateEditTo.date() > QDate.currentDate():
            self.dateEditTo.setDate(QDate.currentDate())

    def setDateTo(self):
        # set dateEditTo.date to dateEditFrom.date
        self.dateEditTo.setDate(self.dateEditFrom.date())

        if self.dateEditFrom.date() > QDate.currentDate():
            self.dateEditFrom.setDate(QDate.currentDate())

    def process(self):
        global pauseTbl, badgeTbl, insertedCount

        # employees to process; process only employees whose CATEGORY_CODE = 'CAGT'
        #sqlStr = "Select NOBLEID, EMPLOYEE_NO from BADGE_CONTROL where EMPLOYEE_NO in " \
        #        "(Select BADGE_NO from employee_badge where CATEGORY_CODE = 'CAGT' and WORK_GROUP_CODE >= '%s' and " \
        #        "WORK_GROUP_CODE <= '%s' and BADGE_NO >= '%s' and BADGE_NO <= '%s')" % \
        #        (self.workgroupComboFrom.currentText(), self.workgroupComboTo.currentText(),
        #        self.employeeComboFrom.itemData(self.employeeComboFrom.currentIndex()), self.employeeComboTo.itemData(self.employeeComboTo.currentIndex()))

        # show the confirmation message
        flags = QMessageBox.StandardButton.Yes
        flags |= QMessageBox.StandardButton.No
        question = "Do you really want to process right now?"
        response = QMessageBox.question(self, "Confirm Process", question, flags, QMessageBox.No)
        if response == QMessageBox.No:
            return

        cur = connOriTMS.cursor()
        # truncate the Noble Exception table
        cur.execute('Truncate table USER_NOBLE_EXCEPTION'
                    '')
        sqlStr = "Select bc.NOBLEID, bc.EMPLOYEE_NO, eb.EMPLOYEE_NAME from BADGE_CONTROL bc, EMPLOYEE_BADGE eb where " \
                 "bc.EMPLOYEE_NO = eb.EMPLOYEE_NO and EMPLOYEE_STATUS = 'A' and CATEGORY_CODE = 'CAGT' and WORK_GROUP_CODE >= '%s' and " \
                "WORK_GROUP_CODE <= '%s' and EMPLOYEE_NAME >= '%s' and EMPLOYEE_NAME <= '%s'" % \
                (self.workgroupComboFrom.currentText(), self.workgroupComboTo.currentText(),
                self.employeeComboFrom.currentText(), self.employeeComboTo.currentText())


        cur.execute(sqlStr)
        badgeListStr = []

        badgeTbl = {}                           # table for NobleID to Orisoft Employee Number conversion
        for rec in cur:
            nobleID = rec[0]
            employee_no = rec[1]
            employee_name = rec[2]
            badgeTbl[nobleID] = { 'number' : employee_no, 'name' : employee_name}
            badgeListStr.append("'%s'" % nobleID)

        badgeListStr = ','.join(badgeListStr)
        totalEmployees = len(badgeTbl)
        self.labelStatus.setText('Number of employees to process : %d' % totalEmployees)
        self.repaint()

        # get pause codes and pause limits from table
        cur.execute('Select pause_code, limit_secs from user_noble_pause')
        pauseTbl = {}
        for rec in cur:
            pauseTbl[rec[0]] = rec[1]

        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()

        sqlStr = "Select employee_no, actual_date, schedule_type, time_in1, time_out1"
        cur.execute()
        now = datetime.now()
        now = datetime(now.year, now.month, now.day, now.hour, now.minute, now.second)
        tblNameDate = dateFrom

        # if selected dates are not equal to current date, add 1 day to them to get the correct table name in
        # history server.
        #if tblNameDate != QDate.currentDate():
        #    tblNameDate = tblNameDate.addDays(1)
        if dateTo != QDate.currentDate():
            dateTo = dateTo.addDays(1)          #add 1 day to get the correct table name in database
            dateTo = dateTo.addDays(1)          #add 1 day to get the correct logout time of the cross-over shifts
                                                #in the dateEditTo widget
        insertedCount = 0
        while tblNameDate <= dateTo:
            if tblNameDate != QDate.currentDate():
                strDate = "%02d%02d%s" % (tblNameDate.month(), tblNameDate.day(), str(tblNameDate.year())[2:])
                tskTblName = 'tsktsrhst' + strDate
                pauTblName = 'tskpauhst' + strDate
                curTsk = connNobleHist.cursor()
                curPau = connNobleHist.cursor()
            else:
                tskTblName = 'tsktsrday'
                pauTblName = 'tskpauday'
                curTsk = connNobleCurr.cursor()
                curPau = connNobleCurr.cursor()

            strSqlTsr = "Select tsr, call_date, appl, logon_time, logoff_time, time_connect, time_paused, time_waiting, time_deassigned, time_acw from %s where tsr in (%s) " \
                    "order by tsr, logon_time, logoff_time" % (tskTblName, badgeListStr)
            strSqlPau = "Select tsr, call_date, end_time, pause_code, pause_time, appl, elem_type, state_num, asn from %s where tsr in (%s) " \
                    "order by tsr, end_time, pause_code" % (pauTblName, badgeListStr)

            if tblNameDate >= self.dateEditFrom.date() and tblNameDate <= self.dateEditTo.date():
                self.labelStatus.setText("Reading %s attendance from Noble" % tblNameDate.toString())
                self.repaint()

            self.saveEmployeeTime(strSqlTsr, strSqlPau, curTsk, curPau)
            tblNameDate = tblNameDate.addDays(1)
            self.repaint()

        self.labelStatus.setText('Process finished')
        self.viewLog()


    def saveEmployeeTime(self, strSqlTsk, strSqlPau, curTsk, curPau):
        #curTsk = connNobleHist.cursor()
        #curPau = connNobleHist.cursor()

        try:
            curTsk.execute(strSqlTsk)
            for rec in curTsk:
                call_date = "%d-%02d-%02d" % (rec[1].year, rec[1].month, rec[1].day)
                tsr = rec[0]

                if tsr not in tsktsr:
                    tsktsr[tsr] = {}

                if call_date not in tsktsr[tsr]:
                    tsktsr[tsr][call_date] = []

                tsktsr[tsr][call_date].append([rec[0], rec[1], rec[2], rec[3], rec[4], rec[5], rec[6], rec[7], rec[8], rec[9]])

            curTsk.close()

            curPau.execute(strSqlPau)
            for rec in curPau:
                currDate = date(rec[1].year, rec[1].month, rec[1].day)
                call_date = "%d-%02d-%02d" % (currDate.year, currDate.month, currDate.day)
                tsr = rec[0]
                if tsr not in tskpau:
                    tskpau[tsr] = {}
                if call_date not in tskpau[tsr]:
                    tskpau[tsr][call_date] = []

                tskpau[tsr][call_date].append(rec)

        except pyodbc.Error, e:
            QMessageBox.critical(None,'Noble Browser Error', str(e))
            self.abort()

        curPau.close()

    def viewLog(self):

        tsrlist = []
        for tsr in sorted(tsktsr):
            name = badgeTbl[tsr]['name']
            emp = badgeTbl[tsr]['number']

            for call_date in sorted(tsktsr[tsr]):
                for rec in tsktsr[tsr][call_date]:
                    line = []
                    line.append(emp)
                    line.append(name)
                    for field in rec:
                        line.append(field)

                    tsrlist.append(line)

        rept = NobleBrowse(tsrlist,['EMPLOYEE_NO', 'NAME', 'TSR', 'CALL_DATE', 'APPL', 'LOGON_TIME', 'LOGOFF_TIME', 'TIME_CONNECT', 'TIME_PAUSED', 'TIME_WAITING',
                                    'TIME_DEASSIGNED', 'TIME_ACW'], self)
        rept.resize(800,600)
        rept.setWindowTitle("TSKTSRHST")
        rept.exec_()

    def canceled(self):
        # show the confirmation message
        flags = QMessageBox.StandardButton.Yes
        flags |= QMessageBox.StandardButton.No
        question = "Do you really want to cancel?"
        response = QMessageBox.question(self, "Confirm Cancel", question, flags, QMessageBox.No)
        if response == QMessageBox.No:
            return

        self.abort()
        #self.parent().reject()

    def abort(self):
        connNobleCurr.close()
        connNobleHist.close()
        connOriTMS.close()
        self.reject()
        app.exit(1)

class NobleBrowse(QDialog):
    def __init__(self, data_list, header, parent=None):
        super(NobleBrowse, self).__init__(parent)
        self.tabBar = QTabWidget()
        self.data_list = data_list
        self.header = header
        self.table_view = QTableView(self.tabBar)
        self.table_model = QStandardItemModel(len(data_list), len(header), self.table_view)

        # set header
        for i in range(len(header)):
            self.table_model.setHeaderData(i, Qt.Horizontal, header[i])

        for row in range(len(data_list)):
            for col in range(len(header)):
                item = QStandardItem('%s' % (data_list[row][col]))
                item.setTextAlignment(Qt.AlignCenter)
                self.table_model.setItem(row, col, item)

        self.table_view.setModel(self.table_model)
        self.table_view.resizeColumnsToContents()

        self.table_view2 = QTableView(self.tabBar)
        self.table_model2 = QStandardItemModel(len(data_list), len(header), self.table_view2)

        # set header
        for i in range(len(header)):
            self.table_model2.setHeaderData(i, Qt.Horizontal, header[i])

        for row in range(len(data_list)):
            for col in range(len(header)):
                item = QStandardItem('%s' % (data_list[row][col]))
                item.setTextAlignment(Qt.AlignCenter)
                self.table_model2.setItem(row, col, item)

        self.table_view2.setModel(self.table_model2)
        self.table_view2.resizeColumnsToContents()

        self.tabBar.addTab(self.table_view,'Noble')
        self.tabBar.addTab(self.table_view2,'Orisoft')
        layout = QVBoxLayout()
        layout.addWidget(self.tabBar)
        self.setLayout(layout)


def readIni():
    global nobleCurrDsn, nobleCurrUser, nobleCurrPwd
    global nobleHistDsn, nobleHistUser, nobleHistPwd
    global orisoftDsn, orisoftUser, orisoftPwd

    # name of configuration file
    iniFile = 'NobleInterface.ini'
    # test if config file exists
    if not os.path.exists(iniFile):
        QMessageBox.critical(None,'Config File Missing', "The configuration file '%s' in '%s' not found!" % \
            (iniFile, os.getcwd()))
        app.exit(1)

    try:
        config = ConfigParser.ConfigParser()
        config.read(iniFile)

        # read Noble Current Table settings
        nobleCurrDsn = config.get('NobleCurrentDSN', 'dsn')
        nobleCurrUser = config.get('NobleCurrentDSN', 'uid')
        nobleCurrPwd = config.get('NobleCurrentDSN', 'pwd')

        # read Noble History Table settings
        nobleHistDsn = config.get('NobleHistoryDSN', 'dsn')
        nobleHistUser = config.get('NobleHistoryDSN', 'uid')
        nobleHistPwd = config.get('NobleHistoryDSN', 'pwd')

        # read Orisoft TMS settings
        orisoftDsn = config.get('OrisoftTMSDSN', 'dsn')
        orisoftUser = config.get('OrisoftTMSDSN', 'uid')
        orisoftPwd = config.get('OrisoftTMSDSN', 'pwd')

    except ConfigParser.NoSectionError, e:
        QMessageBox.critical(None,'Config File Error', str(e))
        app.exit(1)

    except ConfigParser.NoOptionError, e:
        QMessageBox.critical(None, 'Config File Error', str(e))
        app.exit(1)


app = QApplication(sys.argv)
try:
    readIni()
    #connNobleCurr = pyodbc.connect('DSN=mnlcc-apps; UID=root; PWD=VuT9A5hA', autocommit=True)
    #connNobleHist = pyodbc.connect('DSN=mnlcc-ras; UID=root; PWD=VuT9A5hA', autocommit=True)
    #connOriTMS = pyodbc.connect('DSN=ORITMS_Docomo; UID=tmsadm; PWD=gawinmo')
    #connNobleCurr = pyodbc.connect('DSN=Atomix_Local; UID=sa; PWD=justdoit')
    #connNobleHist = pyodbc.connect('DxSN=Atomix_Local; UID=sa; PWD=justdoit')
    #connOriTMS = pyodbc.conitect('DSN=ORITMS_Docomo; UID=sa; PWD=justdoit')

    # connection for Noble Current Database
    connNobleCurr = pyodbc.connect('DSN=%s; UID=%s; PWD=%s' % (nobleCurrDsn, nobleCurrUser, nobleCurrPwd), autocommit=True)
    # connection for Noble History Database
    connNobleHist = pyodbc.connect('DSN=%s; UID=%s; PWD=%s' % (nobleHistDsn, nobleHistUser, nobleHistPwd), autocommit=True)
    # connection for Orisoft TMS Database
    connOriTMS = pyodbc.connect('DSN=%s; UID=%s; PWD=%s' % (orisoftDsn, orisoftUser, orisoftPwd))

except pyodbc.Error, e:
    QMessageBox.critical(None,'Noble Interface Connection Error', str(e))
    app.exit(1)

employeeAttendance = {}                          # employees daily attendance
tskpau = {}
tsktsr = {}
badgeTbl = {}

form = NobleBrowserMainWindow()
form.show()
sys.exit(app.exec_())
