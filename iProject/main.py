if __name__ == "__main__":
    from pprint import pprint
    from reader import *
    from writer import *

    print('Innodisk iReport:\nAuthor: Steven Chou (FAE)\n')
    reader = Reader()
    '''
    for case in reader.cas:
        pprint(vars(case), indent=4)
        pprint(vars(case.product), indent=4)
        print()
    '''
    writer = Writer(reader.cas)
