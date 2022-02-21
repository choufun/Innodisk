series = {
    'D06': {'MLC': '3ME', 'iSLC': '3IE', 'SLC': '3SE', 'iSLC (MLC)': '3IE'},
    'D07': {'MLC': '3ME', 'iSLC': '3IE', 'SLC': '3SE', 'iSLC (MLC)': '3IE'},
    'D08': {'MLC': '3ME3', 'iSLC': '3IE3', 'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'D09': {'MLC': '3ME3', 'iSLC': '3IE3', 'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'M41': {'MLC': '3ME4', 'iSLC': '3IE4', 'SLC': '3SE4', 'iSLC (MLC)': '3IE4'},
    'D67': {'MLC': '3MG-P', 'iSLC': '3IE-P', 'SLC': '3SE-P'},
    'D81': {'MLC': '3MG2-P', 'iSLC': '3IE2-P', 'SLC': '3SE2-P'},
    'D82': {'MLC': '3MG2-P', 'iSLC': '3IE2-P', 'SLC': '3SE2-P'},
    'D70': {'MLC': '3MG3-P', 'iSLC': '3IE3-P', 'SLC': '3SE3-P'},
    'D72': {'MLC': '3ME2'},
    'DK1': {'3D TLC': '3TE7', 'iSLC (3D TLC)': '3IE7'},
    'J30': {'MLC': 'D150QV', 'SLC': 'D150SV-L'},
    'I81': {'SLC': 'SD'},
    'I68': {'SLC': '3SE2'},
    'I61': {'MLC': '3ME', 'SLC': '3SE'},
    'I72': {'SLC': '2SE'},
    'E21': {'MLC': '3ME2'},
    'Y91': {'MLC': '3ME3', 'iSLC': '3IE3', 'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'M61': {'MLC': '3ME2', '3D TLC': '3TE2'},
    'M71': {'MLC': '3MG6-P', '3D TLC': '3TG6-P'},
    'M72': {'3D TLC': '3TG6-P'},
    'D31': {'SLC': '4000'},
    'D51': {'SLC': '4000 Plus'},
    'D53': {'MLC': '1ME', 'iSLC': '1IE', 'SLC': '1SE'},
    'D71': {'SLC': '9000'},
    'M81': {},
    'DC1': {'3D TLC': '3TG6-P'},
    'DD1': {'3D TLC': '3TE6'},
    'DB1': {'3D TLC': '3TE4'},
    'S02': {'SLC': '3SE3', 'iSLC (MLC)': '3IE3'},
    'Y81': {'MLC': 'MLC', 'iSLC': 'iSLC', 'SLC': 'SLC'},
    'D41': {'SLC': '1SE'},
    'J20': {'MLC': '25000 SATA', 'SLC': '10000 Plus SATA'},
    'I21': {'SLC': '3SE'},
    'D21': {'SLC': "2000+"},
    'DA1': {'3D TLC': '3TG3-P'},
    'S06': {'3D TLC': '3TE4'},
    'DQ1': {'SLC': '2SE2',},
    'YA2': {'SLC': '1SE3'},
}

model1 = {
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
        'RPS': 'D150Q',
    },
    'PCIe': {
        'SBC': 'Server Boot Card',
        'M24': 'M.2 (P42)',
        'M28': 'M.2 (P80)',
        'CFX': 'CFexpress',
        'EU2': 'U.2 SSD',
        'EDM': 'Mini PCIeDOM',
    },
    'PATA': {
        'CFC': 'iCF',
        'CF': 'iCF',
        'P25': '2.5" PATA SSD',
    },
    'USB': {
        'UH1': 'USB EDC H',
        'UH2': 'USB EDC H',
        'UA1': 'USB Drive',
        'UF': 'USB EDC',
    },

    'SD': {
        'SDM': "Micro SD",
        'SDC': 'Industrial SD Card'
    }
}

model2 = {
    'DEMSR': 'MSR',
    'DESSL': 'SSL',
    'DHMSR': 'MSR',
    'DGMSR': 'MSR',
    'DGM24': 'M24',
    'DHM28': 'M28',
    'DHM24': 'M24',
    'DESDC': 'SDC',
    'DEM24': 'M24',
    'DECFA': 'CFA',
    'DC1M': 'CF',
    'DGS25': 'S25',
    'DRS25': 'S25',
    'DHS25': 'S25',
    'DUM28': 'M28',
    'DHUH1': 'UH1',
    'DESBC': 'SBC',
    'DS2M': 'SDM',
    'DRPS': 'RPS',
    'DEM28': 'M28',
    'DGM28': 'M28',
    'DS2A': 'SDC',
    'DHSSL': 'SSL',
    'DENSD': 'NSD',
    'DESMV': 'SMV',
    'R2DGMSR': 'MSR',
    'DECFX': 'CFX',
    'DES25': 'S25',
    'DGEU2': 'EU2',
    'DEEDM': 'EDM',
    'R2DEUH1': 'UH1',
    'DGSML': 'SML',
    'DESDM': 'SDM',
    'DHCFA': 'CFA',
    'DHSDC': 'SDC',
    'DGSLM': 'SLM',
    'DTM28': 'M28',
    'DVS25': 'S25',
    'DESSH': 'SSH',
    'DSSML': 'SML',
    'DEUH2': 'UH2',
    'DESSV': 'SSV',
    'DEMSM': 'MSM',
    'DECFC': 'CFC',
}


class Failure:
    def __init__(self, pn, se, ff, r1, r2, ty=None):
        self.pn = pn
        self.se = se
        self.mn = ff
        self.r1 = r1
        self.r2 = r2
        self.ty = ty

    def check(self, pn, r1, r2):
        if self.pn == pn and self.r1 == r1 or self.r2 == r2:
            return False
        else:
            return False

    def info(self):
        return {
            'pn': self.pn,
            'se': self.se,
            'mn': self.mn,
            'r1': self.r1,
            'r2': self.r2,
        }


class Case:
    def __init__(self, excel, col, row):
        self.i = row
        self.n = excel[col["Doc#"] + str(row)].value
        self.t = excel[col["Type"] + str(row)].value
        self.y = excel[col['Submit date'] + str(row)].value.split()[0].split('-')[-1]
        self.m = excel[col['Submit date'] + str(row)].value.split()[0].split('-')[0]

        self.st = excel[col["SubTerritory"] + str(row)].value
        if self.st in ['ZOTHER']:
            self.st = 'USA'

        self.fo = ' '.join(excel[col["FAE Owner"] + str(row)].value.split('_')).title()
        # self.hf = ' '.join(excel[col["Forward FAE Owner"]+str(row)].value.split('_')).title()

        self.s = excel[col["Status"] + str(row)].value
        if self.s in ['FA Report']:
            self.s = 'Closed'

        if self.s in ['Submit', 'Pending']:
            self.s = 'In progress'

        self.cu = excel[col["Customer name"] + str(row)].value.title()
        self.ec = excel[col["End customer name"] + str(row)].value
        self.cr = excel[col["Customer attr"] + str(row)].value
        self.od = excel[col["Submit date"] + str(row)].value
        self.cd = excel[col["FA date"] + str(row)].value


class Flash(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.bu = "FLASH"
        self.fa = None
        self.ft = None
        self.add(excel, col, row, 'c')

    def add(self, excel, col, row, method):
        pn = excel[col["Innodisk PN"] + str(row)].value
        ct = excel[col["FLASH Decode-controller code"] + str(row)].value
        if ct == "M81":
            it = 'PCIe'
        else:
            it = excel[col["ERP Interface"] + str(row)].value
        if excel[col["DRAM Decode-Memory type"] + str(row)].value is None:
            ff = model1[it][model2[pn.split('-')[0]]]
        else:
            ff = model1[it][excel[col["DRAM Decode-Memory type"] + str(row)].value]
        mn = ff
        if ct == "M81":
            se = 'Server'
        else:
            se = series[ct][excel[col["ERP Flash type"] + str(row)].value]

        self.ft = excel[col["ERP Flash type"] + str(row)].value
        print('FLASH Type: ', self.ft)

        if excel[col["Root cause1"] + str(row)].value is not None:
            r1 = excel[col["Root cause1"] + str(row)].value
        else:
            r1 = "Undetermined"
        if excel[col["Root cause2"] + str(row)].value is not None:
            r2 = excel[col["Root cause2"] + str(row)].value
        else:
            r2 = ""

        if method == 'c':
            self.fa = [Failure(pn, se, mn, r1, r2, it)]

        elif method == 'u':
            flag = True
            for fa in self.fa:
                if pn == fa.pn and r1 == fa.r1 and r2 == fa.r2:
                    flag = False

            if flag:
                self.fa.append(Failure(pn, se, mn, r1, r2, it))

    def spec(self):
        return {
            'case': self.n,
            'type:': self.bu,
            'issues': len(self.fa),
        }


class DRAM(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.bu = "DRAM"
        self.fa = None
        self.add(excel, col, row, 'c')

    def add(self, excel, col, row, method):
        pn = excel[col["Innodisk PN"] + str(row)].value
        # mn = excel[col["DRAM Decode-DIMM Type(3rd)"] + str(row)].value
        mn = excel[col["ERP Productline"] + str(row)].value
        sp = excel[col["DRAM Decode-IC Data Rate(4th)"] + str(row)].value
        se = excel[col["ERP Family"] + str(row)].value

        if excel[col["Root cause1"] + str(row)].value is not None:
            r1 = excel[col["Root cause1"] + str(row)].value
        else:
            r1 = "Undetermined"
        if excel[col["Root cause2"] + str(row)].value is not None:
            r2 = excel[col["Root cause2"] + str(row)].value
        else:
            r2 = ""

        if method == 'c':
            self.fa = [Failure(pn, se, mn, r1, r2, sp)]

        elif method == 'u':
            flag = True
            for fa in self.fa:
                if pn == fa.pn and r1 == fa.r1 and r2 == fa.r2:
                    flag = False
            if flag:
                self.fa.append(Failure(pn, se, mn, r1, r2, sp))

    def spec(self):
        return {
            'case': self.n,
            'type:': self.bu,
            'issues': len(self.fa),
        }


class EP(Case):
    def __init__(self, excel, col, row):
        super().__init__(excel, col, row)
        self.bu = "EP"
        self.fa = None
        self.add(excel, col, row, 'c')

    def add(self, excel, col, row, method):
        pn = excel[col["Innodisk PN"] + str(row)].value
        mn = excel[col["Model name"] + str(row)].value
        ty = excel[col["ERP Family"] + str(row)].value
        se = excel[col["ERP Productline"] + str(row)].value
        if excel[col["Root cause1"] + str(row)].value is not None:
            r1 = excel[col["Root cause1"] + str(row)].value
        else:
            r1 = "Undetermined"
        if excel[col["Root cause2"] + str(row)].value is not None:
            r2 = excel[col["Root cause2"] + str(row)].value
        else:
            r2 = ""

        if method == 'c':
            self.fa = [Failure(pn, se, mn, r1, r2, ty)]

        elif method == 'u':
            flag = True
            for fa in self.fa:
                if pn == fa.pn and r1 == fa.r1 and r2 == fa.r2:
                    flag = False
            if flag:
                self.fa.append(Failure(pn, se, mn, r1, r2, ty))

    def spec(self):
        return {
            'case': self.n,
            'type:': self.bu,
            'issues': len(self.fa),
        }
