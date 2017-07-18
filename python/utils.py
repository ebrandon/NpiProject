import xlrd
import xlwt
from xlrd import open_workbook

def output(self, filename, sheet, col_names, data):
    book = xlwt.Workbook()
    sh = book.add_sheet(sheet)
   
    style = xlwt.easyxf('font: bold 1')
    style2 = xlwt.easyxf('pattern: pattern solid, fore_color light_yellow;')
    n = 0
    for i in range(len(col_names)):
        sh.write(0, n, col_names[i], style)
        n += 1
     
    if len(data) == 0:
        print('SOMETHING WENT WRONG WITH THE DATA')
             
    for row_index, row in enumerate(data):
        if len(row) > len(col_names) :
            for col_index, cell_value in enumerate(row):
                sh.write(row_index + 1, col_index, cell_value, style2)
        else:
            for col_index, cell_value in enumerate(row):
                sh.write(row_index + 1, col_index, cell_value)
        book.save(filename)
        # self.finish()
           
def getHeadings(filename):
    wb = open_workbook(filename)
        
    for sheet in wb.sheets():
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols
       
        headings = []
        for col in range(0, number_of_columns):
           heading = (sheet.cell(0, col).value)
           headings.append(heading)
        return headings

def getAddressDetails(data, index):
    if 'address_details_line' in data[index]['practice_address']:
        address2 = data[index]['practice_address']['address_details_line']
        return address2

def getNpi(data, index):
    if 'npi' in data[index]:
        npi = data[index]['npi']
        return npi

def getSpecialty(data, index, curs):
    tax = []
    for index, details in enumerate(data[index]['provider_details']):
        taxCode = details['healthcare_taxonomy_code']
        t = (taxCode, )
        for row in curs.execute('SELECT (Type || \' - \' || Classification || \' - \' || Specialization)'
                                   'FROM HcpSpecialties WHERE Code = ?;', t):
            specialty = row
            tax.append(specialty[0])
    return ', '.join(tax) 

def getLicense(data, index):
    licCode = []
    licState = []
    for num, details in enumerate(data[index]['provider_details']):
        if 'license_number' in details and 'license_number_state' in details and details['taxonomy_switch'] == 'yes':
            licCode = details['license_number']
            licState = details['license_number_state']
        lic = [licCode, licState]
        return lic

if __name__=='__main__':
    main()
