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


class Case(object):
    def __init__(self, excel, index):
        self.number      = None
        self.status      = None
        self.fae         = None
        self.customer    = None
        self.endCustomer = None
        self.openDate    = None
        self.closeDate   = None
        self.localTAT    = None
        self.hqTAT       = None
        self.totalTAT    = None
        self.target      = 1
        self.infoCol     = None
        self.rank        = None
        self.application = None
        self.subject     = ""

        self.loadCaseInfo(excel, index)
        print(self.number, ':', self.customer)

    def loadCaseInfo(self, excel, index):
        self.infoCol = {'Doc#': None, 'Status': None, 'FAE': None, 'Submit Date': None, 'FA Date': None,
                        'Customer Name': None, 'Forward FA TAT': None, 'Local FA TAT': None, 'End Customer Name': None,
                        'All TAT': None, 'Customer Attr.': None, 'Customer Application': None, 'Subject': None}

        for row in excel.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value in self.infoCol.keys():
                    self.infoCol[cell.value] = cell.column_letter

        self.number      = excel[self.infoCol['Doc#']+index].value
        self.status      = excel[self.infoCol['Status']+index].value
        self.fae         = ' '.join(excel[self.infoCol['FAE']+index].value.split('/')[0].split('_')).title()
        self.customer    = excel[self.infoCol['Customer Name']+index].value.title()
        if excel[self.infoCol['End Customer Name']+index].value is not None:
            self.endCustomer = excel[self.infoCol['End Customer Name']+index].value.title()
        self.openDate    = excel[self.infoCol['Submit Date']+index].value
        self.closeDate   = excel[self.infoCol['FA Date']+index].value
        self.localTAT    = int(excel[self.infoCol['Local FA TAT']+index].value)
        self.hqTAT       = int(excel[self.infoCol['Forward FA TAT']+index].value)
        self.totalTAT    = int(excel[self.infoCol['All TAT']+index].value)
        self.rank        = excel[self.infoCol['Customer Attr.']+index].value
        self.application = excel[self.infoCol['Customer Application']+index].value
        if excel[self.infoCol['Subject']+index].value is not None:
            self.subject = excel[self.infoCol['Subject']+index].value.title()

        if self.totalTAT > 10:
            self.target = 0


class Flash(Case):
    def __init__(self, excel, index):
        super().__init__(excel, index)
        self.type             = "Flash"
        self.series           = None
        self.partNumber       = None
        self.modelName        = None
        self.model            = None
        self.failure          = []
        self.attrCol          = None
        self.flashType        = None
        self.interface        = None
        self.controller       = None
        self.controllerModel  = None

        self.loadProdInfo(excel, index)

    def loadProdInfo(self, excel, index):
        self.attrCol = {'*': None, 'Product BU': None, 'Innodisk PN': None, 'Label PN': None,
                        'ERP-Familes': None, 'ERP-Flash': None, 'Root cause1': None, 'Root cause2': None,
                        'Model Name': None, 'ERP-Interface': None, 'ERP-Controller Type': None,
                        'DRAM Decode-Memory Type(2nd)': None, 'FLASH Decode-controller(8th~10th)': None,
                        }

        for row in excel.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value in self.attrCol.keys():
                    self.attrCol[cell.value] = cell.column_letter

        code = excel[self.attrCol['DRAM Decode-Memory Type(2nd)']+index].value

        self.partNumber      = excel[self.attrCol['Innodisk PN']+index].value
        self.modelName       = excel[self.attrCol['Model Name']+index].value
        self.flashType       = excel[self.attrCol['ERP-Flash']+index].value
        self.controller      = excel[self.attrCol['Innodisk PN']+index].value.split('-')[1][3:6]
        self.controllerModel = excel[self.attrCol['ERP-Controller Type']+index].value
        self.interface       = excel[self.attrCol['ERP-Interface']+index].value

        self.debug(code)

        if self.controller == 'M81':
            self.series    = 'PCIe NVMe'
            self.model     = 'Server Boot Card'
            self.interface = 'PCIe'

        elif self.controller == 'D31' or self.controller == 'D41' or self.controller == 'D51' or \
                self.controller == 'D71':
            self.series = series[self.controller][self.flashType]
            self.model  = 'iCF'

        elif self.controller == 'I81':
            self.series = series[self.controller][self.flashType]
            self.model = 'Industrial Micro SD Card'

        elif self.controller == 'J30':
            if code == 'MSR':
                self.series = '2IE'
                self.model = 'mSATA'
            else:
                self.series = series[self.controller][self.flashType]
                self.model = 'InnoLite SATADOM'

        else:
            self.series = series[self.controller][self.flashType]
            self.model  = model[self.interface][code]

        self.addFailure(excel, index)

    def addFailure(self, excel, index):
        rc1 = excel[self.attrCol['Root cause1'] + index].value
        rc2 = excel[self.attrCol['Root cause2'] + index].value

        if rc2 is not None:
            if rc2 != 'NPF':
                if ' '.join([rc1, rc2]) not in self.failure:
                    self.failure.append(' '.join([rc1, rc2]))
            else:
                if rc1 not in self.failure:
                    self.failure.append(rc1)
        else:
            if rc1 not in self.failure:
                self.failure.append(rc1)

    def debug(self, code):
        print(code)
        print('Part Number:', self.partNumber)
        print('Model Name:', self.modelName)
        print('Flash:', self.flashType)
        print('Controller:', self.controller)
        print('Controller Model:', self.controllerModel)
        print('Interface:', self.interface)


class DRAM(Case):
    def __init__(self, excel, index):
        super().__init__(excel, index)
        self.type    = "DRAM"
        self.series  = None
        self.model   = None
        self.speed   = None
        self.IC      = None
        self.config  = None
        self.failure = []
        self.attrCol = None

        self.loadProdInfo(excel, index)
        self.setProdInfo(excel, index)

    def loadProdInfo(self, excel, index):
        self.attrCol = {'*': None, 'Product BU': None, 'Innodisk PN': None, 'Label PN': None,
                        'ERP-Familes': None, 'Root cause L1': None, 'Root cause L2': None,
                        'Model Spec': None, 'ERP-Product Line': None, 'Model Name': None,
                        'ERP-IC': None, 'DRAM Decode-IC Config.(9th)': None,
                        'DRAM Decode-IC Data Rate(4th)': None,
                        'DRAM Decode-DIMM Type(3rd)': None,
        }

        for row in excel.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value in self.attrCol.keys():
                    self.attrCol[cell.value] = cell.column_letter

        self.addFailure(excel, index)

    def addFailure(self, excel, index):
        rc1 = excel[self.attrCol['Root cause L1'] + index].value
        rc2 = excel[self.attrCol['Root cause L2'] + index].value
        if rc2 is not None:
            if rc2 != 'NPF':
                if ' '.join([rc1, rc2]) not in self.failure:
                    self.failure.append(' '.join([rc1, rc2]))
            else:
                if rc1 not in self.failure:
                    self.failure.append(rc1)
        else:
            if rc1 not in self.failure:
                self.failure.append(rc1)

    def setProdInfo(self, excel, index):
        self.speed = excel[self.attrCol['Model Name'] + index].value
        self.series = ' '.join([excel[self.attrCol['ERP-Familes']+index].value, self.speed])
        self.model  = excel[self.attrCol['DRAM Decode-DIMM Type(3rd)']+index].value
        self.IC     = excel[self.attrCol['ERP-IC']+index].value
        self.config = excel[self.attrCol['DRAM Decode-IC Config.(9th)']+index].value


class EP(Case):
    def __init__(self, excel, index):
        super().__init__(excel, index)
        self.type    = "EP"
        self.series  = None
        self.model   = None
        self.failure = []
        self.attrCol = None

        self.loadProdInfo(excel, index)

    def loadProdInfo(self, excel, index):
        self.attrCol = {'*': None, 'Product BU': None, 'Innodisk PN': None, 'ERP-Familes': None,
                        'Root cause L1': None, 'Root cause L2': None, 'Model Spec': None,
                        'ERP-Product Line': None, 'Model Desc': None}

        for row in excel.iter_rows(min_row=1, max_row=1):
            for cell in row:
                if cell.value in self.attrCol.keys():
                    self.attrCol[cell.value] = cell.column_letter

        # self.series = excel[self.attrCol['ERP-Familes']+index].value
        self.series = excel[self.attrCol['ERP-Product Line']+index].value
        # self.model  = excel[self.attrCol['Model Spec']+index].value
        self.model = excel[self.attrCol['Model Desc'] + index].value
        self.addFailure(excel, index)

    def addFailure(self, excel, index):
        rc1 = excel[self.attrCol['Root cause L1'] + index].value
        rc2 = excel[self.attrCol['Root cause L2'] + index].value
        if rc2 is not None:
            if rc2 != 'NPF':
                if ' '.join([rc1, rc2]) not in self.failure:
                    self.failure.append(' '.join([rc1, rc2]))
            else:
                if rc1 not in self.failure:
                    self.failure.append(rc1)
        else:
            if rc1 not in self.failure:
                self.failure.append(rc1)
