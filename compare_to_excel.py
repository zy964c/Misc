import win32com.client
from xlrd import open_workbook
from xlwt import Workbook


def compare_to_excel(path, column_number):
    
    gls_parts = []
    excel = []
    catia = win32com.client.Dispatch('catia.application')
    productDocument1 = catia.ActiveDocument
    collector = catia.ActiveDocument.Product.Products
    for part in range(1, collector.Count + 1):
        gls_parts.append(str(collector.Item(part).PartNumber))

    # acquiring data from excel
    #---------------------------------------------------------------
    
    book = open_workbook(path)
    sheet = book.sheet_by_index(0)

    for row_index in range(sheet.nrows):
        excel.append(str(sheet.cell(row_index, column_number).value))              

    result = []
    for part in excel:
        if part not in gls_parts:
            result.append(part)

    # saving new excel file
    #---------------------------------------------------------------

    book1 = Workbook()
    sheet1 = book1.add_sheet('Sheet 1')
    for i,j in enumerate(result):
        sheet1.write(i,0,j)
    book1.save(path[:-5] + '_new' + '.xls')
    

if __name__ == "__main__":

    compare_to_excel('C:\Temp\zy964c\gls_lib\GL_SEED_report_nomarkers.xlsx', 0)
