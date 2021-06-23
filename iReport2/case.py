class Case:
    def __init__(self, excel, col, row):
        self.n = excel[col["Doc#"]+str(row)].value
        # self.f = excel[col[""]+str(row)].value
        self.s = excel[col["Status"]+str(row)].value
        self.c = excel[col["Customer name"]+str(row)].value.title()
        self.ec = excel[col["End customer name"]+str(row)].value
        # self.od = excel[col[""]+str(row)].value
        # self.cd = excel[col[""]+str(row)].value


class Flash(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.t = "FlASH"
        self.pn = excel[col["Innodisk PN"]+str(row)].value
        self.mn = excel[col["Model name"]+str(row)].value
        self.it = excel[col["ERP Interface"]+str(row)].value
        self.r1 = excel[col["Root cause1"]+str(row)].value
        self.r2 = excel[col["Root cause2"]+str(row)].value


class DRAM(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.t = "DRAM"
        self.pn = excel[col["Innodisk PN"] + str(row)].value
        self.mn = excel[col["DRAM Decode-DIMM Type(3rd)"] + str(row)].value
        self.sp = excel[col["DRAM Decode-IC Data Rate(4th)"] + str(row)].value
        self.se = excel[col["ERP Family"]+str(row)].value
        self.r1 = excel[col["Root cause1"] + str(row)].value
        self.r2 = excel[col["Root cause2"] + str(row)].value


class EP(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.t = "EP"
        self.pn = excel[col["Innodisk PN"] + str(row)].value
        self.mn = excel[col["Model name"] + str(row)].value
        self.se = excel[col["ERP Family"] + str(row)].value
        self.r1 = excel[col["Root cause1"] + str(row)].value
        self.r2 = excel[col["Root cause2"] + str(row)].value
