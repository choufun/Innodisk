import openpyxl

from pprint import pprint

from datetime import datetime


class Writer:
    def __init__(self, stack):
        self.workbook = openpyxl.Workbook()
        self.s = stack

    def report(self):
        self.summary('Summary', 0)
        self.customers('Customers', 1)
        self.flash('FLASH', 2)
        self.dram('DRAM', 3)
        self.ep('EP', 4)
        self.save()

    def summary(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 8
        sheet.column_dimensions['B'].width = 15
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 15
        sheet.column_dimensions['E'].width = 15
        sheet.column_dimensions['F'].width = 25
        sheet.column_dimensions['G'].width = 15
        sheet.column_dimensions['H'].width = 20
        sheet.column_dimensions['I'].width = 50

        sheet.append(['Year', 'Case', 'FAE', 'Region', 'Type', 'Customer', 'BU', 'Series', 'Model'])

        for case in self.s:
            sheet.append([case.y, case.n, case.fo, case.st, case.t, case.cu, case.bu, case.se, case.mn])

    def customers(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 8
        sheet.column_dimensions['B'].width = 15
        sheet.column_dimensions['C'].width = 10
        sheet.column_dimensions['D'].width = 10
        sheet.column_dimensions['E'].width = 10
        sheet.column_dimensions['F'].width = 10
        sheet.column_dimensions['G'].width = 35
        sheet.column_dimensions['H'].width = 35

        sheet.append(['Year', 'Case', 'Region', 'Type', 'BU', 'Rank', 'Customer', 'End Customer'])
        for case in self.s:
            sheet.append([case.y, case.n, case.st, case.t, case.bu, case.cr, case.cu, case.ec])

    def flash(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 15
        sheet.column_dimensions['E'].width = 20
        sheet.column_dimensions['F'].width = 20
        sheet.column_dimensions['G'].width = 28

        sheet.append(['Case', 'Customer', 'Interface', 'Series', 'Model', 'Root Cause 1', 'Root Cause 2'])
        for case in self.s:
            if case.bu == 'FLASH':
                sheet.append([case.n, case.cu, case.it, case.se, case.ff, case.r1, case.r2])

    def dram(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 8
        sheet.column_dimensions['D'].width = 30
        sheet.column_dimensions['E'].width = 15
        sheet.column_dimensions['F'].width = 20
        sheet.column_dimensions['G'].width = 28

        sheet.append(['Case', 'Customer', 'Series', 'Model', 'Speed', 'Root Cause 1', 'Root Cause 2'])
        for case in self.s:
            if case.bu == 'DRAM':
                sheet.append([case.n, case.cu, case.se, case.mn, case.sp, case.r1, case.r2])

    def ep(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 50
        sheet.column_dimensions['E'].width = 20
        sheet.column_dimensions['F'].width = 28

        sheet.append(['Case', 'Customer', 'Series', 'Model', 'Root Cause 1', 'Root Cause 2'])
        for case in self.s:
            if case.bu == 'EP':
                sheet.append([case.n, case.cu, case.se, case.mn, case.r1, case.r2])

    def save(self):
        now = datetime.now()
        date_time = now.strftime("%m.%d.%Y - %H.%M.%S")
        new_filename = str("{} - FAE Report.xlsx".format(date_time))
        self.workbook.save(new_filename)
