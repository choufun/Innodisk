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

        # pprint(self.c)

    def lookahead(self, case, row):
        print("case index: {} {}".format(case.i, self.e['A' + str(case.i)].value))
        print("curr index: {} {}".format(row[0].row, row[0].value))
        print("next index: {} {}".format(row[0].row+1, self.e['A'+str(row[0].row+1)].value))
        print("last index: {} {}".format(len(self.e['A'])-1, self.e['A'+str(len(self.e['A'])-1)].value))

        if case.i == len(self.e['A']) - 1:
            self.s.append(case)
            print('...last case added...')
            print('stack size:', len(self.s), '\n')
        else:
            if self.e['A'+str(row[0].row+1)].value == '*':
                self.s.append(case)
                print('...case added...')
                print('stack size:', len(self.s), '\n')
            else:
                print('...case updated...')
                print('stack size:', len(self.s), '\n')

    def scan(self):
        case = None
        for index, row in enumerate(self.e.iter_rows(min_row=2, max_row=len(self.e['A']))):
            if row[0].value == "*":
                for cell_obj in row:
                    if self.e['F'+str(cell_obj.row)].value != 'DOA':
                        # debug(cell_obj)
                        if cell_obj.value == 'FLASH':
                            case = Flash(self.e, self.c, cell_obj.row)
                            debug(case)
                            self.lookahead(case, row)

                        elif cell_obj.value == 'DRAM':
                            case = DRAM(self.e, self.c, cell_obj.row)
                            debug(case)
                            self.lookahead(case, row)

                        elif cell_obj.value == 'EP':
                            case = EP(self.e, self.c, cell_obj.row)
                            debug(case)
                            self.lookahead(case, row)
            else:
                self.lookahead(case, row)

    def stack(self):
        return self.s


def debug(obj):
    if isinstance(obj, openpyxl.cell.cell.Cell):
        print("DEBUG::\tcolumn: {}".format(obj.column))
        print("DEBUG::\tcolumn_letter: {}".format(obj.column_letter))
        print("DEBUG::\tvalue: {}".format(obj.value))
        print("DEBUG::\trow: {}".format(obj.row))
        print("DEBUG::\tcol_idx: {}".format(obj.col_idx))
        print('\n\n')

    elif isinstance(obj, Flash) or isinstance(obj, DRAM) or isinstance(obj, EP):
        print("index: {} - {} case created".format(obj.i, obj.t))
        print("DEBUG::")
        pprint(obj.spec())
