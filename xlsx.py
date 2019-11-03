import xlrd
import xlwt
import string
from openpyxl import Workbook, load_workbook
from xlutils.copy import copy as shcopy
import xlsxwriter
import copy

class Data(object):

    def __init__(self, target = 0,target_loc = (0,0), 
                 var_loc = (0,0), data = []):
        self.target = target
        self.target_loc = target_loc
        self.var_loc = var_loc
        self.data = data
    
    def __str__(self):
        str = 'target : {0} \ntarget_loc : {1} \nvar_loc : {2} \ndata : {3}'.format(self.target, self.target_loc, self.var_loc, self.data)
        return str

class UkpManager(object):


    def generate_using_template_xlsx(self, path, encoding = None):
        self.wbread = xlrd.open_workbook(path,formatting_info=True, encoding_override = encoding) if encoding else xlrd.open_workbook(path,formatting_info=True)

        worksheet = self.wbread.sheets()

        sheet_name = self.wbread.sheet_names()

        # get target
        target = self.ref_data[0].target_loc.split(',')
        pos_target = string.ascii_lowercase.index(target[0])
        target_val = self.ref_data[0].target * 0.15

        var_loc = self.ref_data[0].var_loc.split(',')
        pos_var = string.ascii_lowercase.index(var_loc[0])

        for ref in self.ref_data[0].data:

            wbwrite = shcopy(self.wbread)

            next = wbwrite._Workbook__worksheets[0]
            next.set_name('add new')
            print(str(pos_target) + ';' + str(pos_var))
            next.write(int(target[1]) - 1, pos_target, target_val , xlwt.easyxf("font: name Candara, height 280; align: horiz center, vert center"))
            print(ref[2]) 
            next.write(int(var_loc[1]) - 1, pos_var, ref[2] , xlwt.easyxf("font: name Candara, height 280; align: horiz center, vert center"))

            #wbwrite._Workbook__worksheets = [next]

            wbwrite.save('pkp_pkm_cianjur' + ref[0] +'.xls')

    def load_data_xls(self, path):
        self.data_wb_read = xlrd.open_workbook(path)

        worksheets = self.data_wb_read.sheets()

        data = []
        for sh in worksheets:
            target = sh.cell(0,0).value
            target_loc = sh.cell(0,1).value
            var_loc = sh.cell(0,2).value
            d = Data(target = target, 
                     target_loc = target_loc, var_loc = var_loc)
            data_rows = []
            is_empty = False
            for row in  range(1, sh.nrows):
                
                data_cols = []
                for col in range(0, sh.ncols):

                    temp = sh.cell_value(row, col)
                    if temp != '' and row == 0:
                        is_empty = True

                    data_cols.append(temp)

                    if is_empty:
                        break

                if is_empty:
                    break
                data_rows.append(copy.deepcopy(data_cols))

            d.data = copy.deepcopy(data_rows)
            data.append(d)

        self.ref_data = data
