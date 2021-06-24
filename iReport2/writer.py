import openpyxl
from pprint import pprint


class Writer:
    def __init__(self, stack):
        self.workbook = openpyxl.Workbook()
        self.s = stack

    def report(self):
        self.summary('Summary', 0)
        self.customers('Customers', 1)
        self.dram('DRAM', 2)
        self.ep('EP', 3)
        self.save()

    def summary(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 25
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 20
        sheet.column_dimensions['E'].width = 50
        sheet.append(['Case', 'Customer', 'BU', 'Series', 'Model'])

        for case in self.s:
            sheet.append([case.n, case.cu, case.t, case.se, case.mn])

    def customers(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 8
        sheet.column_dimensions['C'].width = 8
        sheet.column_dimensions['D'].width = 30
        sheet.column_dimensions['E'].width = 35
        sheet.column_dimensions['F'].width = 30
        sheet.column_dimensions['G'].width = 30

        sheet.append(['Case', 'Type', 'Rank', 'Customer', 'End Customer', 'Application', 'Subject'])
        for case in self.s:
            sheet.append([case.n, case.t, case.cr, case.cu, case.ec, '', ''])

    def dram(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 8
        sheet.column_dimensions['D'].width = 30
        sheet.column_dimensions['E'].width = 15

        sheet.append(['Case', 'Customer', 'Series', 'Model', 'Speed'])
        for case in self.s:
            if case.t == 'DRAM':
                sheet.append([case.n, case.cu, case.se, case.mn, case.sp])

    def ep(self, title, page):
        sheet = self.workbook.create_sheet(title, page)
        sheet.column_dimensions['A'].width = 15
        sheet.column_dimensions['B'].width = 30
        sheet.column_dimensions['C'].width = 15
        sheet.column_dimensions['D'].width = 50

        sheet.append(['Case', 'Customer', 'Series', 'Model'])
        for case in self.s:
            if case.t == 'EP':
                sheet.append([case.n, case.cu, case.se, case.mn])

    def save(self):
        self.workbook.save('FAE Report.xlsx')
