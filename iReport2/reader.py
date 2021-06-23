import openpyxl
from utility import *
from pprint import pprint
from case import *


class Reader:
    def __init__(self, filename):
        self.s = []
        self.e = openpyxl.load_workbook(filename).active
        self.c = {}

        for index, row in enumerate(self.e.iter_rows(min_row=1, max_row=1)):
            for cell_obj in row:
                debug(cell_obj)
                self.c[cell_obj.value] = cell_obj.column_letter

        pprint(self.c)

    def scan(self):
        case = None

        for index, row in enumerate(self.e.iter_rows(min_row=2, max_row=len(self.e['A']))):
            if row[0].value == "*":
                for cell_obj in row:
                    if cell_obj.value == 'FLASH':
                        case = Flash(self.e, self.c, cell_obj.row)
                        print("index: {} - Flash case created".format(cell_obj.row))
                        debug(case)

                    elif cell_obj.value == 'DRAM':
                        case = DRAM(self.e, self.c, cell_obj.row)
                        print("index: {} - DRAM case created".format(cell_obj.row))
                        debug(case)

                    elif cell_obj.value == 'EP':
                        case = EP(self.e, self.c, cell_obj.row)
                        print("index: {} - EP case created".format(cell_obj.row))
                        debug(case)

            else:
                print("index: {}   - updating case...")


def debug(obj):
    if isinstance(obj, openpyxl.cell.cell.Cell):
        print("DEBUG::\tcolumn: {}".format(obj.column))
        print("DEBUG::\tcolumn_letter: {}".format(obj.column_letter))
        print("DEBUG::\tvalue: {}".format(obj.value))
        print("DEBUG::\trow: {}".format(obj.row))
        print("DEBUG::\tcol_idx: {}".format(obj.col_idx))
        print('\n\n')

    elif isinstance(obj, Flash):
        print("DEBUG::\tnumber: {}".format(obj.n))

    elif isinstance(obj, DRAM):
        print("DEBUG::\tnumber: {}".format(obj.n))

    elif isinstance(obj, EP):
        print("DEBUG::\tnumber: {}".format(obj.n))
