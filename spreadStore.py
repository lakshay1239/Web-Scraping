import xlrd
def store(url):
    filelocation = url                 #opens a spreadsheet
    workbook=xlrd.open_workbook(filelocation)
    sheet = workbook.sheet_by_index(0)
    return sheet
