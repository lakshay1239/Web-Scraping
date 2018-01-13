from bs4 import BeautifulSoup
import urllib.request
import xlsxwriter
import xlrd
import spreadStore
import codeFolder
import crawl

cod=[]
codeFolder.code()          #calling codeFolder file
industry=['energy','basic-materials','industrials','cyclicals','non-cyclicals','financials','healthcare','technology','telecom','utilities']
for i in range(0,len(industry)):                                                                                     
        sheet=spreadStore.store('C:\\Users\\Admin\\Desktop\\blueOptima\\newFolder\\'+industry[i]+'\\mysheet.xlsx')               #calling spreadStore file
        for r in range(sheet.nrows):
                cod.append(sheet.cell_value(r,1))                                                                               #strong data in list from file


detail=[]
for h in range(1,len(cod)):                                                                                                     #loop to make all spreadsheets (comapny_detail and employee_detail)
        page = urllib.request.urlopen('https://www.reuters.com/sectors/industries/rankings?industryCode='+ str(int(cod[h])) +'&view=size&page=-1&sortby=mktcap&sortdir=DESC')
        soup=BeautifulSoup(page,"html.parser")                                                                                   # data scrapping
        for i in soup.find("tbody").find_all("td"):
                detail.append('' + i.text + '')
           
        workbook = xlsxwriter.Workbook('COMPANY_'+str(int(cod[h]))+'_DETAIL.xlsx',{'constant_memory':True})                     #storing data in worksheet
        worksheet = workbook.add_worksheet()


        row = 0
        col = 0
        for key in range(0,len(detail)):
            worksheet.write(row, col, detail[key])
            col +=1
            if col%5==0:
                row +=1
                col = 0
                        
        workbook.close()
        detail.clear()
        
        #filelocation = 'C:\\Users\\Admin\\Desktop\\blueOptima\\newFolder\\COMPANY_'+str(int(u[h]))+'_DETAIL.xlsx'
        sheet=spreadStore.store('C:\\Users\\Admin\\Desktop\\blueOptima\\newFolder\\COMPANY_'+str(int(cod[h]))+'_DETAIL.xlsx')     #spreadStore file
        ticker=[]
        for row in range(sheet.nrows):
            ticker.append(""+ sheet.cell_value(row,0) +"")

        for i in range(1,len(ticker)):
            strng = 'https://www.reuters.com/finance/stocks/company-officers/' + ticker[i]                                         #using ticker value to maintain employee data
            spreadName = "Employee_Detail_"+ticker[i] 
            crawl.store(strng,spreadName)                                                                                          # calling crawl file
          
                 
