from pprint import pprint
from reader import *
from writer import *
from clerk import *

if __name__ == "__main__":
    reader = Reader()
    '''
    for case in reader.stack:
        pprint(vars(case))
    '''
    clerk = Clerk(reader.stack)
    writer = Writer(clerk)

