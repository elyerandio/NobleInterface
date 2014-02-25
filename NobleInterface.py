#***********************************
# Program Name : NobleInterface.py
# Date Written : April 25, 2013
# Description  : a program to interface time-in/time-out data from Docomo's Noble System
#                into Orisoft Employee Attendance.
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
from time import sleep
from Attendance import Attendance
from NobleReport import *


class NobleInterfaceMainWindow(QMainWindow):
    def __init__(self, parent=None):
        super(NobleInterfaceMainWindow, self).__init__(parent)
        form = NobleInterfaceForm()
        self.setCentralWidget(form)
        self.setWindowTitle('Noble-Orisoft Integration Clock Upload')
        #form.show()

class NobleInterfaceForm(QDialog):

    def __init__(self, parent=None):
        super(NobleInterfaceForm, self).__init__(parent)

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
        #self.chOverwrite = QCheckBox('Overwrite existing record')
        #self.chOverwrite.setChecked(True)
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
        #layout.addWidget(self.chOverwrite, 4, 1)
        layout.addLayout(buttonLayout, 5, 1, 1, 2)
        layout.addWidget(self.labelStatus, 6, 0, 1, 3)
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
        self.setWindowTitle('Noble Integration Clock Upload')


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
        cur.execute("Select BADGE_NO, EMPLOYEE_NAME from EMPLOYEE_BADGE where CATEGORY_CODE in ('CAGT','CC') order by EMPLOYEE_NAME")
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
        sqlStr = "Select BADGE_NO, EMPLOYEE_NAME from EMPLOYEE_BADGE where CATEGORY_CODE in ('CAGT','CC') and " \
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

    def nightShiftPrev(self, badge, login):

        prevDay = login - timedelta(1)
        prevDay = date(prevDay.year, prevDay.month, prevDay.day).isoformat()
        if prevDay not in empDailySched[badge]:
            return False

        schedIn = empDailySched[badge][prevDay].get("schedIn", time(0,0,0))
        schedOut = empDailySched[badge][prevDay].get("schedOut", time(0,0,0))

        if schedOut < schedIn:
            return True

    def saveEmployeeAttendance(self, attendance_id, employee, badge, name, currDate, schedule, login, logout, doc6ss_breaktime, \
                               doc6ss_break_from, doc6ss_break_to, totalOverBreak):
        global insertedCount

        seq = 1
        now = datetime.now()
        (totalWorkHour, totalOtHour, elapsedTimeIn, elapsedTimeOut) = self.computeTotalHour(schedule, login, logout)

        print "SaveEmployeeAttendance", doc6ss_break_from, doc6ss_break_to, doc6ss_breaktime
        if not doc6ss_break_from:
            doc6ss_break_from = ''
        if not doc6ss_break_to:
            doc6ss_break_to = ''

        if schedule.startswith('R'):
            if self.nightShiftPrev(badge, login):
                login = ''
                logout = ''
            else:
                #replace 00:00:00 login/logout with NULL
                if login.hour == 0 and login.minute == 0 and login.second == 0:
                    login = ''
                if logout.hour == 0 and logout.minute == 0 and logout.second == 0:
                    logout = ''
        else:
            #replace 00:00:00 login/logout with NULL
            if login.hour == 0 and login.minute == 0 and login.second == 0:
                login = ''
            if logout.hour == 0 and logout.minute == 0 and logout.second == 0:
                logout = ''

        if login == '' and logout <> '':
            logout = ''


        cursor = connOriTMS.cursor()

        # if employee has no schedule for current date, save record to user_noble_exception table for reporting purpose
        if schedule == 'NA':
            cursor.execute("Insert into user_noble_exception(EMPLOYEE_NO, EMPLOYEE_NAME, CALL_DATE, LOGIN, LOGOUT, REMARKS, CREATED_BY, CREATED_DATE) " \
                "VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % (employee, name, currDate, login, logout, 'No Schedule', 'Noble', now.strftime('%Y-%m-%d %H:%M:%S')))

        #check if employee is on leave for current date
        remarks = ''
        if employee in leaveInfo:
            if currDate.isoformat() in leaveInfo[employee]:
                remarks = leaveInfo[employee][currDate.isoformat()]

        # repeatCount variable is used to avert an infinite loop, the while loop will exit if the value is more than 1
        repeatCount = 0
        while 1:
            repeatCount += 1
            try:
                cursor.execute("Insert into employee_attendance(ID, BADGE_NO, EMPLOYEE_NO, SEQ_NO, ACTUAL_DATE, SCHEDULE_TYPE, MACHINE_IN1, TIME_IN1, MACHINE_OUT1, TIME_OUT1, \
                TOTAL_WORK_HOUR, ELAPSED_TIMEIN, ELAPSED_TIMEOUT, TOTAL_OT_HOUR, REMARKS, CREATED_BY, CREATED_DATE, DOC6SS_BREAKTIME, DOC6SS_BREAK_FROM, DOC6SS_BREAK_TO, \
                DOC6SS_TOTAL_OVERBREAK) VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % \
                    (attendance_id, employee, employee, seq, currDate, schedule, login, login, logout, logout, totalWorkHour, elapsedTimeIn, elapsedTimeOut, totalOtHour, remarks, 'Noble',
                    now.strftime('%Y-%m-%d %H:%M:%S'), doc6ss_breaktime, doc6ss_break_from, doc6ss_break_to, totalOverBreak))
                attendance_id = attendance_id + 1
                insertedCount = insertedCount + 1
                return attendance_id

            except pyodbc.IntegrityError, e:
                # check the overwrite data checkbox
                #if self.chOverwrite.isChecked():
                # Record exists, delete record
                cursor.execute("Delete from EMPLOYEE_ATTENDANCE where BADGE_NO = '%s' and ACTUAL_DATE = '%s' and SEQ_NO = %d" % (employee, currDate, 1))

                # check if repeated already
                if repeatCount > 1:
                    return attendance_id


    def saveDataToOrisoft(self):
        cursor = connOriTMS.cursor()

        self.crossOverShift(employees)
        interface_id = 0
        #cursor.execute("select ctrlctr from ofcctrlid where ctrlcol = 'interface_time'")
        cursor.execute("select max(ID) from interface_time")
        for rec in cursor:
            interface_id = int(rec[0]) + 1

        # get the next available id for employee_attendance table
        attendance_id = 0
        #cursor.execute("select ctrlctr from ofcctrlid where ctrlcol = 'employee_attendance'")
        cursor.execute("select max(ID) from employee_attendance")
        attendance_id = int(cursor.fetchone()[0]) + 1

        #totalEmployees = len(employees)
        i = 0
        for call_date in sorted(employees):
            i = i + 1
            #print 'employee = [%s]' % employee, 'recCnt = ', len(employees[employee])
            #restrict date to the selection date
            if call_date < self.dateEditFrom.date().toPython().isoformat() or call_date > self.dateEditTo.date().toPython().isoformat():
                continue
            self.labelStatus.setText("Saving %s Attendance" % call_date)
            for badge in sorted(employees[call_date]):
                strLoginTime = employees[call_date][badge].getTgenStr()
                strLogoutTime = employees[call_date][badge].getLogoutTimeStr()
                strLogoutTimePrev = employees[call_date][badge].getLogoutTimePrevStr()
                currDate = employees[call_date][badge].loginDate
                schedule = employees[call_date][badge].schedule

                #skip absent employee i.e. no timein/timeout
                #if strLoginTime == '00:00:00' and strLogoutTime == '00:00:00':
                #    continue
                #elif (strLoginTime == '00:00:00' or strLogoutTime == '00:00:00'): # and schedule.startswith('R'):
                #    continue

                #restrict date to the selection date
                if currDate < self.dateEditFrom.date().toPython() or currDate > self.dateEditTo.date().toPython():
                    continue

                self.labelStatus.setText("Saving %s Attendance For %s" % (badge, currDate.isoformat()))
                self.repaint()

                employee = badgeTbl[badge]['number']
                name = badgeTbl[badge]['name']
                loginDate = employees[call_date][badge].loginDate
                loginTime = employees[call_date][badge].tgen
                login = datetime(loginDate.year, loginDate.month, loginDate.day, loginTime.hour, loginTime.minute, loginTime.second)
                logoutDate = employees[call_date][badge].logoutDate
                logoutTime = employees[call_date][badge].logoutTime
                logout = datetime(logoutDate.year, logoutDate.month, logoutDate.day, logoutTime.hour, logoutTime.minute, logoutTime.second)
                doc6ss_breaktime = employees[call_date][badge].doc6ss_breaktime
                doc6ss_break_from = employees[call_date][badge].doc6ss_break_from
                doc6ss_break_to = employees[call_date][badge].doc6ss_break_to
                if '1HR ' in employees[call_date][badge].pauses:
                    over1HR = employees[call_date][badge].pauses['1HR ']
                else:
                    over1HR = 0

                if '30M ' in employees[call_date][badge].pauses:
                    over30M = employees[call_date][badge].pauses['30M ']
                else:
                    over30M = 0

                if '15M ' in employees[call_date][badge].pauses:
                    over15M = employees[call_date][badge].pauses['15M ']
                else:
                    over15M = 0

                if 'BATH' in employees[call_date][badge].pauses:
                    overBath = employees[call_date][badge].pauses['BATH']
                else:
                    overBath = 0

                # if employee has no SCHEDULE for current date, assign 'NA' to SCHEDULE_TYPE just to be able to insert it to EMPLOYEE_ATTENDANCE table
                if schedule == '':
                    schedule = 'NA'

                totalOverBreak = over1HR + over30M + over15M + overBath
                #cursor.execute("Insert into user_noble_interface(CALL_DATE, EMPLOYEE_BADGE, TSR, SCHEDULE_TYPE,LOGIN, LOGOUT, [1HR], [30M], [15M], BATH, MODIFIED_DATE) " \
                #    "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",  currDate.isoformat(), badge, employee, schedule, login.toPython() , logout.toPython(), over1HR, over30M, over15M, overBath,
                #               QDateTime().currentDateTime().toPython())
                cursor.execute("Insert into user_noble_interface(CALL_DATE, EMPLOYEE_BADGE, TSR, NAME, SCHEDULE_TYPE,LOGIN, LOGOUT, [1HR], [30M], [15M], BATH, MODIFIED_DATE) " \
                    "VALUES(?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",  currDate.isoformat(), employee, badge, name, schedule, login , logout, over1HR, over30M, over15M, overBath,
                               datetime.now())

                attendance_id = self.saveEmployeeAttendance(attendance_id, employee, badge, name, currDate, schedule, login, logout, doc6ss_breaktime, doc6ss_break_from, \
                                                            doc6ss_break_to, totalOverBreak)
                connOriTMS.commit()

        # update the next available id for employee_attendance table
        msg = "update ofcctrlid set ctrlctr = %d where ctrlcol = 'employee_attendance'" % (attendance_id)
        cursor.execute(msg)

        # update EMPLOYEE_ATTENDANCE to replace the 'NA' SCHEDULE_TYPE with NULL
        dateFrom = self.dateEditFrom.date().toPython()
        dateTo = self.dateEditTo.date().toPython()
        cursor.execute("Update EMPLOYEE_ATTENDANCE set SCHEDULE_TYPE = NULL where ACTUAL_DATE between '%s' and '%s' and SCHEDULE_TYPE ='NA'" % (dateFrom, dateTo))

        # update EMPLOYEE_ATTENDANCE to replace default login/logout of 1900-01-01 00:00:00.0 with NULL
        cursor.execute("Update EMPLOYEE_ATTENDANCE set machine_in1 = NULL, time_in1 = NULL where ACTUAL_DATE between '%s' and '%s' and machine_in1 = '1900-01-01 00:00:00.0'" %
                       (dateFrom, dateTo))
        cursor.execute("Update EMPLOYEE_ATTENDANCE set machine_out1 = NULL, time_out1 = NULL where ACTUAL_DATE between '%s' and '%s' and machine_out1 = '1900-01-01 00:00:00.0'" %
                       (dateFrom, dateTo))
        connOriTMS.commit()

    def createDailyAttendance(self, tblNameDate, employees, badgeTbl):

        # to get current date, subtract 1 from tblNameDate
        currDate = tblNameDate.addDays(-1)
        call_date = "%d-%02d-%02d" % (currDate.year(), currDate.month(), currDate.day())
        if call_date not in employees:
            employees[call_date] = {}

        for emp in badgeTbl:
            if emp not in employees[call_date]:
                employees[call_date][emp] = Attendance(call_date)

                if emp not in empDailySched:
                    empDailySched[emp] = {}
                    self.getEmployeeDailySched(emp, empDailySched)


            if call_date in empDailySched[emp]:
                sched = empDailySched[emp][call_date].get("schedType",'')
                employees[call_date][emp].schedule = sched

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
        cur.execute('Truncate table dbo.user_noble_exception')
        sqlStr = "Select control_no, EMPLOYEE_NO, EMPLOYEE_NAME from EMPLOYEE_BADGE where " \
                 "EMPLOYEE_STATUS = 'A' and CATEGORY_CODE IN ('CAGT','CC') and WORK_GROUP_CODE >= '%s' and " \
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

        # save run datetime to user_noble_interface_run table
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

            strSqlTsk = "Select call_date, tsr, logon_time, logoff_time from %s where tsr in (%s) " \
                    "order by tsr, logon_time, logoff_time" % (tskTblName, badgeListStr)
            strSqlPau = "Select call_date, tsr, end_time, pause_code, pause_time from %s where tsr in (%s) " \
                    "order by tsr, pause_code, end_time" % (pauTblName, badgeListStr)

            if tblNameDate >= self.dateEditFrom.date() and tblNameDate <= self.dateEditTo.date():
                self.labelStatus.setText("Reading %s attendance from Noble" % tblNameDate.toString())
                self.repaint()

            self.createDailyAttendance(tblNameDate,employees, badgeTbl)
            self.saveEmployeeTime(strSqlTsk, strSqlPau, curTsk, curPau,employees, empDailySched)
            tblNameDate = tblNameDate.addDays(1)
            self.repaint()

        self.getLeaveInfo()
        self.saveDataToOrisoft()
        self.labelStatus.setText('Process finished')

        # save history log
        cur.execute('Insert into USER_NOBLE_INTERFACE_RUN( RUN_DATE, WORKGROUP_FROM, WORKGROUP_TO, EMPLOYEE_FROM, EMPLOYEE_TO, DATE_FROM, DATE_TO, NUMBER_OF_EMPLOYEES, INSERTED_RECORDS)'  \
            " VALUES( '%s', '%s', '%s', '%s', '%s', '%s', '%s', %d, %d)" % (now, self.workgroupComboFrom.currentText(), self.workgroupComboTo.currentText(), \
            self.employeeComboFrom.currentText(), self.employeeComboTo.currentText(), dateFrom.toPython() , dateTo.toPython(), totalEmployees, insertedCount))
        cur.close()
        connOriTMS.commit()

        while 1:
            msg = "Finished!\n%d employee(s) processed.\nAttendance from %s to %s uploaded to Orisoft.\n" \
                  "%d records inserted into Employee Attendance table" % (totalEmployees,
                self.dateEditFrom.text(), self.dateEditTo.text(), insertedCount)
            msgBox = QMessageBox()
            msgBox.setText(msg)
            msgBox.setIcon(QMessageBox.Information)
            msgBox.setWindowTitle('Noble-Orisoft Interface')
            btnViewExcReport = msgBox.addButton('View Exception Report', QMessageBox.AcceptRole)
            btnOk = msgBox.addButton('Exit', QMessageBox.RejectRole)
            msgBox.exec_()
            if msgBox.clickedButton() == btnViewExcReport:
                self.viewExceptionReport()
                self.abort()
            else:
                try:
                    connNobleCurr.close()
                    connNobleHist.close()
                    connOriTMS.close()
                except pyodbc.Error:
                    pass

                sys.exit(0)
                self.abort()


    def saveEmployeeTime(self, strSqlTsk, strSqlPau, curTsk, curPau, employees,empDailySched):
        #curTsk = connNobleHist.cursor()
        #curPau = connNobleHist.cursor()

        print 'Start saveEmployeeTime'
        print 'strSqlTsk', strSqlTsk
        print 'strSqlPau', strSqlPau
        try:
            curTsk.execute(strSqlTsk)
            #empDailySched = {}
            for rec in curTsk:
                #self.labelStatus.setText('Processing employee #%d' % len(employees))
                call_date = "%d-%02d-%02d" % (rec[0].year, rec[0].month, rec[0].day)
                tsr = rec[1]

                #parse and convert login from int to datetime.time
                loginStr = "%06d" % rec[2]                                          #convert login from int to 6-char string
                (hour, min, sec) = (loginStr[:2], loginStr[2:4], loginStr[4:])      #extract hour, min, sec
                login = time(int(hour), int(min), int(sec))

                logoutStr = "%06d" % rec[3]
                (hour, min, sec) = (logoutStr[:2], logoutStr[2:4], logoutStr[4:])
                logout = time(int(hour), int(min), int(sec))
                att = Attendance(call_date)

                if call_date not in employees:
                    employees[call_date] = {}

                if tsr not in employees[call_date]:
                    employees[call_date][tsr] = att

                if tsr not in empDailySched:
                    empDailySched[tsr] = {}
                    self.getEmployeeDailySched(tsr, empDailySched)

                if employees[call_date][tsr].getLoginTimeStr() == '00:00:00' or employees[call_date][tsr].loginTime > login:
                    employees[call_date][tsr].loginTime = login
                if employees[call_date][tsr].logoutTime < logout:
                    employees[call_date][tsr].logoutTime = logout

                # get the logout time for previous day if schedule is a graveyard shift (cross-over shift)
                if tsr in empDailySched and call_date in empDailySched[tsr]:
                    sched = empDailySched[tsr][call_date].get("schedType",'')
                    employees[call_date][tsr].schedule = sched
                    # if current day is Rest Day, get the previous day's schedule
                    if sched.startswith('R'):
                        prevDay = (employees[call_date][tsr].loginDate - timedelta(1)).isoformat()
                        if prevDay not in empDailySched[tsr]:
                            continue

                        sched = empDailySched[tsr][prevDay].get("schedType",'')
                        schedIn = empDailySched[tsr][prevDay].get("schedIn", time(0,0,0))
                        schedOut = empDailySched[tsr][prevDay].get("schedOut", time(0,0,0))
                    else:
                        schedIn = empDailySched[tsr][call_date].get("schedIn", time(0,0,0))
                        schedOut = empDailySched[tsr][call_date].get("schedOut", time(0,0,0))

                    if schedOut < schedIn:
                        if login < time(12,0) and logout < schedIn and logout < time(12,0):
                            employees[call_date][tsr].logoutTimePrev = logout
                            #print "logoutPrev = ", logout

                # if logout time is 23:59:59 then the employee is in night shift
                # the official logout date/time is in the next-day table
                if logoutStr == '23:59:59':
                    employees[call_date][tsr].isCrossOver = True

                # get the logout time for previous day login, i.e. the employee is night shift
                if login.isoformat() == '00:00:00' and employees[call_date][tsr].logoutTimePrev < logout:
                    employees[call_date][tsr].logoutTimePrev = logout

                #print "Call_Date : ", call_date, "\tDaily Time:", employees[tsr][call_date]

            curTsk.close()
            #print employees

            curPau.execute(strSqlPau)
            for rec in curPau:
                currDate = date(rec[0].year, rec[0].month, rec[0].day)
                call_date = "%d-%02d-%02d" % (currDate.year, currDate.month, currDate.day)
                prevDay = currDate - timedelta(1)                              #subtract 1 day from call_date to get previous day
                prevDay = prevDay.isoformat()
                tsr = rec[1]
                endTimeStr = "%06d" % rec[2]
                (hour, min, sec) = (endTimeStr[:2], endTimeStr[2:4], endTimeStr[4:])
                end_time = time(int(hour), int(min), int(sec))
                pause_code = rec[3]
                pause_time = int(rec[4])
                att = Attendance(call_date)

                if call_date not in employees:
                    employees[call_date] = {}

                if tsr not in employees[call_date]:
                    employees[call_date][tsr] = att

                if tsr not in empDailySched:
                    empDailySched[tsr] = {}
                    self.getEmployeeDailySched(tsr, empDailySched)

                if tsr in empDailySched and call_date in empDailySched[tsr]:
                    sched = empDailySched[tsr][call_date].get("schedType",'')
                    schedIn = empDailySched[tsr][call_date].get("schedIn",time(0,0,0))
                    schedOut = empDailySched[tsr][call_date].get("schedOut",time(0,0,0))

                    if sched.startswith('R'):
                        if prevDay in empDailySched[tsr]:
                            schedIn = empDailySched[tsr][prevDay].get("schedIn", time(0,0,0))
                            schedOut = empDailySched[tsr][prevDay].get("schedOut", time(0,0,0))
                        else:
                            schedIn = time(0,0,0)
                            schedOut = time(0,0,0)
                else:
                    schedIn = time(0,0,0)
                    schedOut = time(0,0,0)

                if pause_code == 'TGEN':
                    if schedOut < schedIn:                                  #cross-over shift
                        #if end_time > schedOut:                             #login time for cross-over shift
                        if end_time > time(12,0,0):                          #login time for cross-over shift
                            if employees[call_date][tsr].getTgenStr() == "00:00:00":
                                #or end_time < employees[call_date][tsr].tgen:
                                employees[call_date][tsr].tgen = end_time
                    else:                                                   #day shift
                        if employees[call_date][tsr].getTgenStr() == "00:00:00" or end_time < employees[call_date][tsr].tgen:
                            employees[call_date][tsr].tgen = end_time
                else:
                    # compute for total break in seconds
                    if pause_code in pauseTbl:
                        if pause_code not in employees[call_date][tsr].pauses:
                            #employees[call_date][tsr].addPauseTime(pause_code,0)
                            employees[call_date][tsr].pauses[pause_code] = 0

                        pauseLimit = pauseTbl.get(pause_code, 0)
                        if schedOut < schedIn:
                            #cross-over shift
                            if end_time <= schedOut:
                                #previous day shift
                                if prevDay not in employees:
                                    employees[prevDay] = {}
                                if tsr not in employees[prevDay]:
                                    employees[prevDay][tsr] = Attendance(prevDay)
                                if pause_code not in employees[prevDay][tsr].pauses:
                                    employees[prevDay][tsr].pauses[pause_code] = 0
                                if pause_time > pauseLimit:
                                    employees[prevDay][tsr].pauses[pause_code] = employees[prevDay][tsr].pauses[pause_code] + (pause_time - pauseLimit)

                                employees[prevDay][tsr].doc6ss_breaktime += pause_time

                                # get doc6ss_break fields
                                if pause_code == '1HR' or pause_code == '1HR ':
                                    break_to = datetime(currDate.year, currDate.month, currDate.day, end_time.hour, end_time.minute, end_time.second)
                                    break_from = break_to - timedelta(seconds = pause_time)
                                    employees[prevDay][tsr].doc6ss_break_to = break_to
                                    employees[prevDay][tsr].doc6ss_break_from = break_from
                                    print "Break :" , break_from, break_to, pause_time
                            elif end_time > schedOut:
                                #current day shift
                                if pause_time > pauseLimit:
                                    #employees[call_date][tsr].addPauseTime(pause_code, pause_time - pauseLimit)
                                    employees[call_date][tsr].pauses[pause_code] = employees[call_date][tsr].pauses[pause_code] + (pause_time - pauseLimit)

                                employees[call_date][tsr].doc6ss_breaktime += pause_time

                                # get doc6ss_break fields
                                if pause_code == '1HR' or pause_code == '1HR ':
                                    break_to = datetime(currDate.year, currDate.month, currDate.day, end_time.hour, end_time.minute, end_time.second)
                                    break_from = break_to - timedelta(seconds = pause_time)
                                    employees[call_date][tsr].doc6ss_break_to = break_to
                                    employees[call_date][tsr].doc6ss_break_from = break_from
                                    print "Break :" , break_from, break_to, pause_time
                        else:
                            #regular day shift
                            employees[call_date][tsr].doc6ss_breaktime += pause_time
                            if pause_time > pauseLimit:
                                #employees[call_date][tsr].addPauseTime(pause_code, pause_time - pauseLimit)
                                employees[call_date][tsr].pauses[pause_code] = employees[call_date][tsr].pauses[pause_code] + (pause_time - pauseLimit)

                            # get doc6ss_break fields
                            if pause_code == '1HR' or pause_code == '1HR ':
                                break_to = datetime(currDate.year, currDate.month, currDate.day, end_time.hour, end_time.minute, end_time.second)
                                break_from = break_to - timedelta(seconds = pause_time)
                                employees[call_date][tsr].doc6ss_break_to = break_to
                                employees[call_date][tsr].doc6ss_break_from = break_from
                                print "Break :" , break_from, break_to, pause_time

        except pyodbc.Error, e:
            QMessageBox.critical(None,'Noble Interface Error', str(e) + '\n\nNo data was saved to Orisoft!')
            self.abort()

        curPau.close()
        #print "end saveEmployeeTime"

    def crossOverShift(self, employees):
        """
        Adjust the logout date/time and the doc6ss_break_to/doc6ss_break_from for cross-over shifts.
        """
        for call_date in employees:
            if call_date < self.dateEditFrom.date().toPython().isoformat():
                continue

            for emp in employees[call_date]:
                currDate = employees[call_date][emp].loginDate
                strLoginTime = employees[call_date][emp].getTgenStr()
                strLogoutTime = employees[call_date][emp].getLogoutTimeStr()
                strLogoutTimePrev = employees[call_date][emp].getLogoutTimePrevStr()

                #skip absent employee i.e. no timein/timeout
                if strLoginTime == '00:00:00' and strLogoutTime == '00:00:00':
                    continue

                #check for the logout time of the crossover shift for previous day
                if strLogoutTimePrev != '00:00:00':
                    # subtract 1 day from loginDate to get the previous day
                    prevDayStr = (employees[call_date][emp].loginDate - timedelta(1)).isoformat()
                    #print "call_date = ", call_date, "prevDayStr = ", prevDayStr

                    if emp in employees[prevDayStr]:
                        #if employees[prevDayStr][emp].getLogoutTimeStr() == '23:59:59':
                        employees[prevDayStr][emp].logoutDate = currDate
                        employees[prevDayStr][emp].logoutTime = employees[call_date][emp].logoutTimePrev
                        if employees[prevDayStr][emp].doc6ss_break_from == None:
                            employees[prevDayStr][emp].doc6ss_break_from = employees[call_date][emp].doc6ss_break_from
                            employees[prevDayStr][emp].doc6ss_break_to = employees[call_date][emp].doc6ss_break_to

                        #print 'call date =', call_date, 'call logout =', employees[emp][call_date].logoutTime, \
                            #'call login =', employees[call_date][emp].loginTime, 'prevLogout = ', employees[call_date][emp].logoutTimePrev
                        #print 'prev day =', prevDayStr, 'prev logout =', employees[emp][prevDayStr].logoutTime, \
                            #'prev login =', employees[prevDayStr][emp].loginTime, 'prevLogout = ', employees[prevDayStr][emp].logoutTimePrev

    def getLeaveInfo(self):
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()
        cur = connOriTMS.cursor()
        cur.execute("Select li.employee_no, ld.leave_date, ld.status, ld.approve_date, li.leave_code from employee_leave_day ld, employee_leave_info li" \
            " where ld.employee_id = li.employee_id and ld.reference_id = li.reference_no and ld.status in ('A') and " \
            "approve_date between '%s' and '%s'" % (dateFrom.toPython(), dateTo.toPython()))

        for rec in cur:
            employee = rec[0]
            leave_date = rec[1].date().isoformat()
            status = rec[2]
            approve_date = rec[3].date().isoformat()
            leave_code = rec[4]

            if employee not in leaveInfo:
                leaveInfo[employee] = {}

            leaveInfo[employee][approve_date] = leave_code



    def getEmployeeDailySched(self, tsr, dailySched):
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date().addDays(2)
        badge = badgeTbl[tsr]['number']
        cur = connOriTMS.cursor()
        cur.execute("Select schedule_date, schedule_type, schedule_type_description from employee_schedule a, schedule_type b" \
            #' where a.schedule_type = b.schedule_type_code and badge_no = ? and schedule_date between ? and ?', badge, dateFrom.toPython(), dateTo.toPython())
            " where a.schedule_type = b.schedule_type_code and badge_no = '%s' and schedule_date between '%s' and '%s'" % (badge, dateFrom.toPython(), dateTo.toPython()))
        if tsr not in dailySched:
            dailySched[tsr] = {}
        print 'Employee Daily Schedule', tsr, badge
        pattern = re.compile(r"(\d\d:\d\d)\s*-\s*(\d\d:\d\d)")
        for rec in cur:
            schedDate = rec[0].date().isoformat()
            schedType = rec[1]
            schedDesc = rec[2]

            # check schedule type code if numeric
            match = re.search(r"^\d\d\d\d", schedType)
            if match:
                match = re.search(pattern, schedDesc)
                if match:
                    schedIn = match.group(1)
                    schedOut = match.group(2)

                    (hour, min) = schedIn.split(':')
                    schedIn = time(int(hour), int(min))
                    (hour, min) = schedOut.split(':')
                    schedOut = time(int(hour), int(min))
            else:
                if schedType.upper() == 'REST':
                    schedTime = 'Rest Day'
                    schedIn = 'Rest Day'
                    schedOut = 'Rest Day'
                else:
                    schedTime = schedType
                    schedIn = schedType
                    schedOut = schedType

            dailySched[tsr][schedDate] = { 'schedType' : schedType, 'schedIn' : schedIn, 'schedOut' : schedOut}
            #print tsr, badge, schedDate, schedType, schedIn, schedOut
            #print 'dailySched', dailySched[tsr]

    def computeOTHour(self, schedule, login, logout):

        #if no login or no logout, return with default value
        if (login.hour == 0 and login.minute == 0 and login.second == 0) or \
                (logout.hour == 0 and logout.minute ==0 and logout.second == 0):
            return datetime(2000, 1, 1, 0, 0, 0)

        totalOT = 0
        if schedule.isdigit():
            schedIn = datetime(login.year, login.month, login.day,int(schedule[:2]),0)
            schedOut = datetime(logout.year, logout.month, logout.day, int(schedule[2:]),0)
            if schedIn > login:
                totalOT = (schedIn - login).seconds
            if logout > schedOut:
                totalOT = totalOT + (logout - schedOut).seconds
        else:
            # if no schedule or schedule is Rest Day, use only login & logout for computation
            totalOT = (logout - login).seconds

        #get the hour and remainder by dividing the total seconds with 60*60
        (hour, totalOT) = divmod(totalOT, 60 * 60)
        if hour < 0:
            hour = 0
        elif hour > 23:
            hour = 23

        #get the minute and second by dividing the remaining seconds with 60
        (minute, second) = divmod(totalOT, 60)
        if minute > 59:
            minute = 59
        if second > 59:
            second = 59

        totalOTHour = datetime(2000,01,01,hour,minute,second)
        return totalOTHour

    def computeTotalHour(self, schedule, login, logout):
        default =  datetime(2000, 1, 1, 0, 0, 0)

        elapsedTimeIn = 0
        elapsedTimeOut = 0
        if not schedule or schedule == 'NA' or schedule.strip() == '' or schedule.startswith('R'):
            totalWorkHour = datetime(default.year, default.month, default.day, default.hour, default.minute, default.second)
            elapsedTimeIn = 0
            elapsedTimeOut = 0
            #compute Total_Ot_Hour
            totalOTHour = self.computeOTHour(schedule, login, logout)
        else:
            # compute total_work_hour (in seconds), subtract 1 hour for 1-hr break
            totalWorkHour = (logout - login).seconds - 3600
            if totalWorkHour < 0:
                totalWorkHour = totalWorkHour * -1

            hours = totalWorkHour / (60 * 60)
            if hours < 0:
                hours = 0
            elif hours > 23:
                hours = 23

            totalWorkHour = totalWorkHour % (60 * 60)           # get the remaining seconds after the hour
            minutes = totalWorkHour / 60
            if minutes > 59:
                minutes = 59

            seconds = totalWorkHour % 60
            if seconds > 59:
                seconds = 59

            # convert totalWorkHour to datetime, w/ default date of '2000-01-01'
            totalWorkHour = datetime(default.year, default.month, default.day, hours, minutes, seconds)

            #compute Total_Ot_Hour
            totalOTHour = self.computeOTHour(schedule, login, logout)

            #compute for elapsed_timein
            schedIn = datetime(login.year, login.month, login.day, int(schedule[:2]), 0, 0)
            if login > schedIn:
                elapsedTimeIn = (login - schedIn).seconds
            elif login < schedIn:
                elapsedTimeIn = (schedIn - login).seconds * -1
            else:
                elapsedTimeIn = 0

            #compute for elapsed_timeout
            schedOut = datetime(logout.year, logout.month, logout.day, int(schedule[2:]), 0, 0)
            if logout > schedOut:
                elapsedTimeOut = (logout - schedOut).seconds
            elif logout < schedOut:
                elapsedTimeOut = (schedOut - logout).seconds * -1
            else:
                elapsedTimeOut = 0

        return (totalWorkHour, totalOTHour, elapsedTimeIn, elapsedTimeOut)

    def viewLog(self):
        cursor = connOriTMS.cursor()
        cursor.execute("Select * from USER_NOBLE_INTERFACE_RUN order by ID desc")
        runHistory = []
        for rec in cursor:
            line = []
            if rec[1]:                                    # Run_Date
                dt = rec[1].isoformat()
                dt = dt.replace("T", " ")                 # replace 'T' in datetime with a space
                line.append(dt)
            else:
                line.append('')
            line.append(rec[2])                             # Workgroup_From
            line.append(rec[3])                             # Workgroup_To
            line.append(rec[4])                             # Employee_From
            line.append(rec[5])                             # Employee_To
            if rec[6]:                                      # Date_From
                line.append(rec[6])
            else:
                line.append('')
            if rec[7] <> None:                              # Date_To
                line.append(rec[7])
            else:
                line.append('')
            line.append(rec[8])                             # Number of employees
            line.append(rec[9])                             # Inserted records

            runHistory.append(line)
            #runHistory.append([rec[0], rec[1].isoformat(), rec[2], rec[3], rec[4], rec[5], rec[6].isoformat(), rec[7].isoformat(), rec[8], rec[9]])
        cursor.close()
        rept = NobleReport(runHistory,['RUN DATE', 'WORKGROUP FROM', 'WORKGROUP TO', 'EMPLOYEE FROM', 'EMPLOYEE TO', 'DATE FROM', 'DATE TO', 'NUMBER OF EMPLOYEES', 'INSERTED RECORDS'], self)
        rept.resize(800,600)
        rept.setWindowTitle("Run History")
        rept.exec_()

    def viewExceptionReport(self):
        cursor = connOriTMS.cursor()
        cursor.execute("Select * from USER_NOBLE_EXCEPTION order by EMPLOYEE_NAME, CALL_DATE")
        exceptionReport = []
        for rec in cursor:
            line = []
            line.append(rec[1])                 # employee no
            line.append(rec[2])                 # employee name
            line.append(rec[3])                 # call date
            line.append(rec[4])                 # schedule type
            #login
            if rec[5]:
                dt = rec[5]
                if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
                    dt = ''
                else:
                    dt = dt.isoformat()

                dt = dt.replace("T", " ")                 # replace 'T' in datetime with a space
                line.append(dt)
            else:
                line.append('')
            #logout
            if rec[6]:
                dt = rec[6]
                if dt.hour == 0 and dt.minute == 0 and dt.second == 0:
                    dt = ''
                else:
                    dt = dt.isoformat()
                dt = dt.replace("T", " ")                 # replace 'T' in datetime with a space
                line.append(dt)
            else:
                line.append('')

            line.append(rec[7])                             # remarks
            line.append(rec[8])                             # created by
            line.append(rec[9])                             # created date

            exceptionReport.append(line)
        cursor.close()
        rept = NobleReport(exceptionReport,['EMPLOYEE NO', 'EMPLOYEE NAME', 'CALL DATE', 'SCHEDULE_TYPE', 'LOGIN', 'LOGOUT', 'REMARKS', 'CREATED BY', 'CREATED DATE'], self)
        rept.resize(800,600)
        rept.setWindowTitle("Exception Report")
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
        try:
            connNobleCurr.close()
            connNobleHist.close()
            connOriTMS.close()
        except pyodbc.Error, e:
            pass

        self.reject()
        app.exit(1)

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


nobleCurrDsn = ''
NobleCurrUser = ''
nobleCurrPwd = ''
nobleHistDsn = ''
nobleHistUser = ''
nobleHistPwd = ''
orisoftDsn = ''
orisoftUser = ''
orisoftPwd = ''

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

employees = {}                          # employees daily attendance
empDailySched = {}                      # employees daily schedule
leaveInfo = {}                          # employees leave information
insertedCount = 0                       # number of records inserted to Employee_Attendance table

form = NobleInterfaceMainWindow()
form.show()
sys.exit(app.exec_())
