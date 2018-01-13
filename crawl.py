from bs4 import BeautifulSoup
import requests
import unicodedata2
import re
import xlsxwriter

def store(url,sname):
    r = requests.get(url)
    data = r.text
    soup = BeautifulSoup(data, "html.parser")                                                   #data scrapping
    flag=0
    content = []
    if len(soup.find_all("tbody"))>0:                                                           #checking the availability of tbody
        if len(soup.find_all("tbody"))==1:
            flag=1
        for d in soup.find_all("tbody",{"class" :"dataSmall"}):                           
            content.append(d.get_text())

        #emp_data = unicodedata.normalize('NFKD',l[0]).encode('ascii', 'ignore')
        emp_data = content[0]
        emp_data_list = re.split(r'\s{2,}',emp_data)                                            # regular expression

        name = []                                                                               # lists storing different data as required
        Age = []
        Since = []
        Cur_Position = []
        c = 1
        for i in range(2,len(emp_data_list)-1):                                                 # extracting data
            s = emp_data_list[i]
            if c==1:
              name.append(s)
              c= c+1
              continue
            if c==2:
                if len(s)>7 :                                                                   #when 'age' and 'since' not available
                    Age.append("UNavailable")
                    Since.append("Unavailable")
                    Cur_Position.append(s)
                    c = 1
                    continue
                if len(s)==7 :                                                                  #when 'age' and 'since' both available
                   Age.append(s.split()[0])
                   Since.append(s.split()[1])
                   c = c+1
                   continue
                elif len(s)==4:                                                                 #when only 'since' is available
                    Age.append("Unavailable")
                    Since.append(s)
                    c = c+1
                    continue
                elif len(s)==2:                                                                 #when only 'age' is available
                    Since.append("Unavailable")
                    Age.append(s)
                    c = c+1
                    continue
            if c==3:                                                                            # for current_position
                Cur_Position.append(s)
                c=1

        #print len(Age)
        #print (Age)
        # Extraction of Description and name from given URL
        #emp_data = unicodedata.normalize('NFKD',l[1]).encode('ascii', 'ignore')
        if flag==0:
            emp_data = content[1]
            emp_data_list = re.split(r'\s{2,}',emp_data)                                                # for biography


            name2 = []
            desc = []

            c = 1
            for i in range(2,len(emp_data_list)-1):                                                 #extracting data
                s = emp_data_list[i]
                if c==1:
                  name2.append(s)
                  c= c+1
                  continue
                if c==2:
                    if len(s)>25 :                                                                  #biographies
                        desc.append(s)
                        c = 1
                        continue
                    if len(s)<=25 :                                                                 #names
                       desc.append("NONE")
                       name2.append(s)
                       c = 2
                       continue


            workbook = xlsxwriter.Workbook(sname+'.xlsx',{'constant_memory':True})                  #storing data in spreadsheet
            worksheet = workbook.add_worksheet()
            row = 0
            col1 = 0
            col2 = 1
            col3 = 2
            col4 = 3
            col5 = 4
            worksheet.write(row,col1,'NAME')
            worksheet.write(row,col2,'AGE')
            worksheet.write(row,col3,'SINCE')
            worksheet.set_column(3,3,70)
            worksheet.write(row,col4,'CURRENT_POSITION')
            worksheet.write(row,col5,'DESCRIPTION')
            row = 1
            
            for key in range(0,len(name)-1):
                worksheet.write(row,col1,name[key])
                worksheet.set_column(3,3,70)
                worksheet.write(row,col2,Age[key])
                worksheet.write(row,col3,Since[key])
                worksheet.write(row,col4,Cur_Position[key])
                worksheet.write(row,col5,desc[key])
                row +=1
           
            workbook.close()
        

    else:
       emp=['No Employee data']
       workbook = xlsxwriter.Workbook(sname+'.xlsx',{'constant_memory':True})
       worksheet = workbook.add_worksheet()
       row = 0
       col = 0
       for key in range(0,len(emp)):
            worksheet.write(row,col,emp[key])
            row +=1
            
       workbook.close()
       
