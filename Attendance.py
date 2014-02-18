# ****************************************
# Program Name : Attendance.py
# Date Written : April 27, 2013
# Purpose      : A class for the daily attendance of an employee. Includes date, login, logout, and overbreak total for each pause code.
# Author       : Eleazer L. Erandio
# ****************************************

import datetime

class Attendance:

    def __init__(self, date=None):
        if isinstance(date, datetime.date):
            self.loginDate = date
            self.logoutDate = date
        elif isinstance(date, basestring):
            if 'T' in date:
                (date, null) = date.split('T')
            (year, month, day) = date.split('-')
            self.loginDate = datetime.date(int(year), int(month), int(day))
            self.logoutDate = datetime.date(int(year), int(month), int(day))
        elif date is None:
            self.loginDate = datetime.date.today()          # set default login date to current date
            self.logoutDate = datetime.date.today()         # set default logout date to current date

        self.loginTime = datetime.time()
        self.logoutTime = datetime.time()
        self.tgen = datetime.time()
        self.logoutTimePrev = datetime.time()           # logout for previous day's night shift schedule
        self.pauses = {}                                # dictionary of each pause overbreak
        self.schedule = ''
        self.doc6ss_break_from = None
        self.doc6ss_break_to = None
        self.doc6ss_breaktime = 0
        self.isCrossOver = False                        # True if employee schedule is night shift

    def getLoginDateStr(self):
        """
        Get the login date as a string 'YYYY-MM-DD'
        """
        return self.loginDate.isoformat()

    def getLogoutDateStr(self):
        """
        Get the logout date as a string 'YYYY-MM-DD'
        """
        return self.logoutDate

    def getLoginTimeStr(self):
        """
        Get the login time as a string 'hh:mm:ss'
        """
        return self.loginTime.isoformat()


    def setLoginTimeStr(self, value):
        """
        Set the login time using a string time i.e. 'hh:mm:ss'
        """
        (hour, min, sec) = value.split(':')
        self.loginTime = datetime.time(int(hour), int(min), int(sec))


    def getLogoutTimeStr(self):
        """
        Get the logout time as a string
        """
        return self.logoutTime.isoformat()


    def setLogoutTimeStr(self, value):
        """
        Set the logout time using a string time
        """
        (hour, min, sec) = value.split(':')
        self.logouTime = datetime.time(int(hour), int(min), int(sec))

    def getLogoutTimePrevStr(self):
        """
        Get the logout time for previous day i.e. crossover shift as a string 'hh:mm:ss'
        """
        return self.logoutTimePrev.isoformat()

    def getTgenStr(self):
        """
        Get the TGEN time as a string
        """
        return self.tgen.isoformat()


    def setTgenStr(self, value):
        """
        Set the TGEN time using a string time
        """
        (hour, min, sec) = value.split(':')
        self.tgen = datetime.time(int(hour), int(min), int(sec))

    def addPauseCode(self, value):
        """
        Add a new pause code to the pauses dictionary and initialize to zero
        """
        if value not in self.pauses:
            self.pauses[value] = 0


    def addPauseTime(self, pause, value):
        """
        Add positive value to the pause total time
        """
        if value < 0:
            return

        if pause in self.pauses:
            self.pauses[pause] += value
        else:
            self.pauses[pause] = value


    def getPauseTime(self, key):
        """
        Get the total time of the pause code given
        """
        if key in self.pauses:
            return self.pauses[key]
        else:
            return 0

    def __str__(self):
        result = "Date : %s\tLogin : %s\tLogout : %s\tLogoutTimePrev : %s\tPauses : %s\n" % (self.loginDate, self.tgen, self.logoutTime, self.logoutTimePrev, self.pauses)
        return result


if __name__ == '__main__':
    att = Attendance(datetime.date.today())
    att.date = datetime.date.today()
    #att.dateStr = '1970-02-18'
    #att.setDateStr('1970-02-18')
    att.setLoginTimeStr('19:25:20')
    att.logoutTime = datetime.time(19, 35, 59)
    print att.date, att.getDateStr(), att.loginTime, att.getLoginTimeStr(), att.getLogoutTimeStr()
    att.addPauseCodeTime('1HR', 65)
    print 'pause code', att.pauses.keys(), att.getPauseCodeTime('1HR')