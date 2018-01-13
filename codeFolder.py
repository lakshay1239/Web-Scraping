import urllib.request
from bs4 import BeautifulSoup
import xlsxwriter
import xlrd
import requests
import os
import spreadStore

class code:
     def __init__(self):
        industry=['energy','basic-materials','industrials','cyclicals','non-cyclicals','financials','healthcare','technology','telecom','utilities']
        os.chdir('C:/Users/Admin/Desktop/blueOptima/newFolder')
        root_path = '' 
        for i in industry:
            os.mkdir(os.path.join(root_path, i))
        
        links = []
        d=[]
        for indus in range(0,len(industry)):
            d.clear()
            links.clear()
            page1 = urllib.request.urlopen('https://www.reuters.com/sectors/'+industry[indus])
            soup=BeautifulSoup(page1,"html.parser")
            os.chdir('C:/Users/Admin/Desktop/blueOptima/newFolder/'+industry[indus])
            root_path ='' 
            for lio in soup.find("div",{"class":"sectionRelatedTopics"}).find_all("li"):
                v = lio.find('a')
                lin="https://www.reuters.com"+v['href']
                if "industryCode" in v['href']:
                    links.append(v['href'])
                    d.append(str(v.text))
                    page2 = urllib.request.urlopen(lin)
                    soup=BeautifulSoup(page2,"html.parser")
                    j={}
                    for lio in soup.find("div",{"class":"sectionRelatedTopics relatedIndustries"}).find_all("li"):
                        v = lio.find('a')
                        links.append(v['href'])
                        d.append(str(v.text))    
                        '''s=links[0]
                        print(v.text)'''
                    for i in range(0,len(links)):
                        ind=links[i].index('=')
                        var=len(links[i])
                        key = int(str(links[i])[ind+1:var])
                        j[d[i]]=key
                    workbook = xlsxwriter.Workbook('mysheet.xlsx',{'constant_memory':True})
                    worksheet = workbook.add_worksheet()
                    sheet=spreadStore.store("C:\\Users\\Admin\\Desktop\\blueOptima\\blue.xlsx")
                    sheet.cell_value(0,0)
                    sheet.nrows
                    sheet.ncols
                    row = 0
                    col = 0
                    worksheet.write(row,col,'SECTOR')
                    worksheet.write(row,col+1,'CODE')
                    worksheet.write(row,col+2,'PERM_ID')
                    worksheet.set_column(2,2,15)
                    worksheet.set_column(0,0,25)
                    row = 1
                    col = 0
                    for i in range(0,len(links)):
                        if(d[i]=="Aerospace / Defense"):
                           os.mkdir(os.path.join(root_path, "Aerospace"))
                           os.chdir('C:/Users/Admin/Desktop/blueOptima/newFolder/'+industry[indus]+'/'+"Aerospace")
                        else:
                            os.mkdir(os.path.join(root_path, d[i]))
                            os.chdir('C:/Users/Admin/Desktop/blueOptima/newFolder/'+industry[indus]+'/'+d[i])
                        worksheet.write(row,col,d[i])
                        worksheet.write(row,col+1,j[d[i]])
                        for r in range(sheet.nrows):
                            if sheet.cell_value(r,3)==d[i]:
                                worksheet.write(row,col+2,sheet.cell_value(r,4))
                        row+=1
                        os.chdir('C:/Users/Admin/Desktop/blueOptima/newFolder/'+industry[indus])
                    workbook.close()
                    os.chdir('C:/Users/Admin/Desktop/blueOptima/newFolder/')
                    break
