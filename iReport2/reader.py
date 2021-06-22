import openpyxl


class Reader:
    def __init__(self, filename):
        self.stack = []
        self.excel = openpyxl.load_workbook(filename).active

        for row in self.excel.iter_rows(min_row=1, max_row=1):
            for cell in row:
                print(cell)
