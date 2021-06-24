series = {
    'D06' : {'MLC': '3ME',    'iSLC': '3IE',    'SLC': '3SE', 'iSLC (MLC)': '3IE'},
    'D07' : {'MLC': '3ME',    'iSLC': '3IE',    'SLC': '3SE', 'iSLC (MLC)': '3IE'},
    'D08' : {'MLC': '3ME3',   'iSLC': '3IE3',   'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'D09' : {'MLC': '3ME3',   'iSLC': '3IE3',   'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'M41' : {'MLC': '3ME4',   'iSLC': '3IE4',   'SLC': '3SE4', 'iSLC (MLC)': '3IE4'},
    'D67' : {'MLC': '3MG-P',  'iSLC': '3IE-P',  'SLC': '3SE-P'},
    'D81' : {'MLC': '3MG2-P', 'iSLC': '3IE2-P', 'SLC': '3SE2-P'},
    'D82' : {'MLC': '3MG2-P', 'iSLC': '3IE2-P', 'SLC': '3SE2-P'},
    'D70' : {'MLC': '3MG3-P', 'iSLC': '3IE3-P', 'SLC': '3SE3-P'},
    'D72' : {'MLC': '3ME2'},
    'DK1' : {'3D TLC': '3TE7'},
    'J30' : {'MLC': 'D150QV', 'SLC': 'D150SV-L'},
    'I81' : {'SLC': 'SD'},
    'I68' : {'SLC': '3SE2'},
    'I61' : {'MLC': '3ME', 'SLC': '3SE'},
    'I72' : {'SLC': '2SE'},
    'E21' : {'MLC': '3ME2'},
    'Y91' : {'MLC': '3ME3',   'iSLC': '3IE3',   'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'M61' : {'MLC': '3ME2',   '3D TLC' : '3TE2'},
    'M71' : {'MLC': '3MG6-P', '3D TLC' : '3TG6-P'},
    'M72' : {'3D TLC': '3TG6-P'},
    'D31' : {'SLC': '4000'},
    'D51' : {'SLC': '4000 Plus'},
    'D53' : {'MLC': '1ME',    'iSLC': '1IE',    'SLC': '1SE'},
    'D71' : {'SLC': '9000'},
    'M81' : {},
    'DC1' : {'3D TLC' : '3TG6-P'},
    'DD1' : {'3D TLC' : '3TE6'},
    'DB1' : {'3D TLC' : '3TE4'},
    'S02' : {'SLC': '3SE2'},
    'Y81' : {'MLC': 'MLC', 'iSLC': 'iSLC', 'SLC': 'SLC'},
    'D41' : {'SLC': '1SE'},
    'J20' : {'MLC': '25000 SATA', 'SLC': '10000 Plus SATA'},
    'I21' : {'SLC': '3SE'},
    'D21' : {'SLC': '2000+'}
}

nand = {
    'SEASK': 'MLC',
    'C1': 'SLC',
    '3C2QC': 'iSLC',
    'BC1DC': 'MLC',
    'BW1DC': 'MLC',
    'ECBQF': '3D TLC',
    'EWBQF': '3D TLC',
    'GCCDL': '3D TLC',
    'BWADC': 'MLC',
    'EC3QF': '3D TLC',
    'BC2SC': 'MLC',
    'SCBSB': 'SLC',
    'EWAQF': '3D TLC',
    'BC1QC': 'MLC',
    'BWBQC': 'MLC',
    'BCAQC': 'MLC',
    'BC2DC': 'MLC',
}

model = {
    'SATA': {
        'M24': 'M.2 (S42)',
        'M28': 'M.2 (S80)',
        'SSH': 'SATADOM',
        'SSL': 'SATADOM',
        'SSF': 'SATADOM',
        'SML': 'SATADOM',
        'SMV': 'SATADOM',
        'SSV': 'SATADOM',
        'SSC': 'SATADOM',
        'SMH': 'SATADOM',
        'S18': '1.8" SATA SSD',
        'S25': '2.5" SATA SSD',
        'NSD': 'nanoSSD',
        'MSR': 'mSATA',
        'MSM': 'mSATA mini',
        'CFA': 'CFast',
        'SLM': 'SATA Slim',
        'ST2': '2.5"',
        'SNH': 'ServerDOM',
    },
    'PCIe': {
        'M24': 'M.2 (P42)',
        'M28': 'M.2 (P80)',
        'CFX': 'CFexpress',
    },
    'PATA': {
        'CFC': 'iCF',
        'P25': '2.5" PATA SSD',
    },
    'USB': {
        'UH1': 'USB EDC H',
        'UA1': 'USB Drive',
        'UF' : 'USB EDC',
    },

    'SD': {
        'SDM': "Micro SD",
        'SDC': 'Industrial SD Card'
    }
}


class Case:
    def __init__(self, excel, col, row):
        self.i = row
        self.n = excel[col["Doc#"]+str(row)].value
        # self.f = excel[col[""]+str(row)].value
        self.s = excel[col["Status"]+str(row)].value
        self.cu = excel[col["Customer name"]+str(row)].value.title()
        self.ec = excel[col["End customer name"]+str(row)].value
        self.cr = excel[col["Customer attr"]+str(row)].value
        # self.od = excel[col[""]+str(row)].value
        # self.cd = excel[col[""]+str(row)].value


class Flash(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.t = "FLASH"
        self.pn = excel[col["Innodisk PN"]+str(row)].value
        # print(self.pn)
        self.ct = ''.join(self.pn.split('-')[1][3:6])
        # self.nf = nand[''.join(self.pn.split('-')[1][6::])]
        self.mn = excel[col["Model name"]+str(row)].value
        self.it = excel[col["ERP Interface"]+str(row)].value
        # self.se = series[self.ct][self.nf]
        self.se = excel[col["ERP Family"]+str(row)].value
        self.r1 = excel[col["Root cause1"]+str(row)].value
        self.r2 = excel[col["Root cause2"]+str(row)].value
        # self.fa = [' '.join([self.r1, self.r2])]

    def spec(self):
        return {
            'type:': self.t,
            'part numbers': self.pn,
            'model name': self.mn,
            'interface': self.it,
            'ctrl': self.ct
        }


class DRAM(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.t = "DRAM"
        self.pn = excel[col["Innodisk PN"] + str(row)].value
        self.mn = excel[col["DRAM Decode-DIMM Type(3rd)"] + str(row)].value
        self.sp = excel[col["DRAM Decode-IC Data Rate(4th)"] + str(row)].value  # MHz
        self.se = excel[col["ERP Family"]+str(row)].value  # DDR3 or DDR4
        self.r1 = excel[col["Root cause1"] + str(row)].value
        self.r2 = excel[col["Root cause2"] + str(row)].value

    def spec(self):
        return {
            'type:': self.t,
            'part numbers': self.pn,
            'model name': self.mn,
            'speed': self.sp,
            'series': self.se,
        }


class EP(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.t = "EP"
        self.pn = excel[col["Innodisk PN"] + str(row)].value
        self.mn = excel[col["Model name"] + str(row)].value
        self.se = excel[col["ERP Family"] + str(row)].value
        self.r1 = excel[col["Root cause1"] + str(row)].value
        self.r2 = excel[col["Root cause2"] + str(row)].value

    def spec(self):
        return {
            'type:': self.t,
            'part numbers': self.pn,
            'model name': self.mn,
            'series': self.se,
        }
