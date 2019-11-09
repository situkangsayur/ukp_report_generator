import xlrd
import xlwt
import string
from openpyxl import Workbook, load_workbook
from xlutils.copy import copy as shcopy
import xlsxwriter
import copy

class Data(object):

    def __init__(self, fields =  [], data = []):
        self.fields = fields
        self.data = data
    
    def __str__(self):
        str = 'fields : {0}  \ndata : {1}'.format(self.fields, self.data)
        return str

class UkpManager(object):


    def generate_using_template_xlsx(self, path, encoding = None, output_path = None):
        self.wbread = xlrd.open_workbook(path,formatting_info=True, encoding_override = encoding) if encoding else xlrd.open_workbook(path,formatting_info=True)

        worksheet = self.wbread.sheets()

        sheet_name = self.wbread.sheet_names()

        style = "font: name Candara, height 280; align: horiz center, vert center; borders: left 1, top 1, bottom 1, right 1;"

        
        for data_sh in self.ref_data:
            

            # create workbook write and copy workbook read
            wbwrite = shcopy(self.wbread)
            
            template_sheet = wbwrite.get_sheet(0)

            new_sheets = []

            for ref in data_sh.data:
                
                #next = wbwrite.add_sheet(ref[0])

                next = copy.copy(template_sheet)

                for col in data_sh.fields[1:]:

                    # get target
                    coor = col.split(',')
                    pos_x = string.ascii_uppercase.index(coor[0])
                    pos_y = int(coor[1])
                    data_idx = data_sh.fields.index(col) 
                    value = ref[data_idx]

                    real_val =  value if (type(value) != str) else (value  if value[0] != '~' else xlwt.Formula(value[1:]))

                    next.write(pos_y - 1, pos_x, real_val, xlwt.easyxf(style))

                next.set_name('capaian kegiatan ' + ref[0])

                new_sheets.append(copy.deepcopy(next))


            wbwrite._Workbook__worksheets = new_sheets
            wbwrite.save('pkp_pkm_cianjur.xls' if output_path == None else output_path+ '/pkp_pkm_cianjur.xls')
        print('done')

    def load_data_xls(self, path):
        self.data_wb_read = xlrd.open_workbook(path)

        worksheets = self.data_wb_read.sheets()

        data = []
        for sh in worksheets:
            
            d = Data()
            data_rows = []
            is_empty = False
           
            for row in  range(0, sh.nrows):
                
                data_cols = []
                empty_col = 0
                for col in range(1, sh.ncols):

                    if sh.cell_value(0, col) != '':
                        temp = sh.cell_value(row, col)

                        if temp == '':
                            empty_col += 1
                            break

                        data_cols.append(temp)

                if empty_col > 1:
                    break

                if row == 0:
                    d.fields = copy.deepcopy(data_cols)
                else:
                    data_rows.append(copy.deepcopy(data_cols))

            d.data = copy.deepcopy(data_rows)
            data.append(d)

        self.ref_data = data
