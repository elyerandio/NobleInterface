# ****************************************
# Program Name : EmployeeAttendance.py
# Date Written : April 27, 2013
# Purpose      : A class for the daily attendance of an employee. Includes date, login, logout, and overbreak total for each pause code.
# Author       : Eleazer L. Erandio
# ****************************************
from Attendance import Attendance

class EmployeeAttendance(Attendance):

    def __init__(self):
        self.attendance = {}

    def append(self, key, value):
        if key not in self.attendance:
            self.attendance[key] = value

