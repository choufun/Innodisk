from pprint import pprint
import datetime


class Clerk:
    def __init__(self, stack):
        self.stack = stack
        self.stack.reverse()

        self.engineers = []
        self.engineerStack = {}
        self.annualStack = {}
        self.closedStack = {}
        self.productStack = {'Flash': {}, 'DRAM': {}, 'EP': {}}
        self.customerStack = {}

        self.addsEngineers()
        self.sortsByEngineer()
        self.sortsByDate()
        self.sortByProduct()
        self.sortByCustomer()
        self.sortByClosure()

        pprint(self.productStack)

    def addsEngineers(self):
        for case in self.stack:
            if case.fae not in self.engineers:
                self.engineers.append(case.fae)

        for engineer in self.engineers:
            self.engineerStack[engineer] = []

    def sortsByEngineer(self):
        for case in self.stack:
            self.engineerStack[case.fae].append(case)

        for engineer in self.engineerStack.keys():
            self.engineerStack[engineer].reverse()

    def sortsByDate(self):
        for engineer, cases in self.engineerStack.items():
            for case in cases:
                month, day, year = case.openDate.split('/')
                if engineer not in self.annualStack.keys():
                    self.annualStack[engineer] = {}
                if year not in self.annualStack[engineer].keys():
                    self.annualStack[engineer][year] = {}
                if month not in self.annualStack[engineer][year].keys():
                    self.annualStack[engineer][year][month] = []
                self.annualStack[engineer][year][month].append(case)

    def sortByProduct(self):
        for case in self.stack:
            month, day, year = case.openDate.split('/')
            if case.type == 'Flash':
                if year not in self.productStack[case.type].keys():
                    self.productStack[case.type][year] = {}
                if month not in self.productStack[case.type][year].keys():
                    self.productStack[case.type][year][month] = {}
                if case.interface not in self.productStack[case.type][year][month].keys():
                    self.productStack[case.type][year][month][case.interface] = []
                self.productStack[case.type][year][month][case.interface].append(case)

            if case.type == 'DRAM':
                if year not in self.productStack[case.type].keys():
                    self.productStack[case.type][year] = {}
                if month not in self.productStack[case.type][year].keys():
                    self.productStack[case.type][year][month] = []
                self.productStack[case.type][year][month].append(case)

            if case.type == 'EP':
                if year not in self.productStack[case.type].keys():
                    self.productStack[case.type][year] = {}
                if month not in self.productStack[case.type][year].keys():
                    self.productStack[case.type][year][month] = []
                self.productStack[case.type][year][month].append(case)

    def sortByCustomer(self):
        for case in self.stack:
            month, day, year = case.openDate.split('/')
            if year not in self.customerStack:
                self.customerStack[year] = {}
            if month not in self.customerStack[year].keys():
                self.customerStack[year][month] = []
            self.customerStack[year][month].append(case)

    def sortByClosure(self):
        for engineer, cases in self.engineerStack.items():
            for case in cases:
                if case.closeDate is None:
                    pass
                else:
                    month, day, year = case.closeDate.split('/')
                    if engineer not in self.closedStack.keys():
                        self.closedStack[engineer] = {}
                    if year not in self.closedStack[engineer].keys():
                        self.closedStack[engineer][year] = {}
                    if month not in self.closedStack[engineer][year].keys():
                        self.closedStack[engineer][year][month] = []
                    self.closedStack[engineer][year][month].append(case)
