import logging
import time
import traceback
import xlrd
import csv
import sys
import re
import datetime
import shutil

files = ["Dashboard_REVPAR_month.xlsx","employment_data.xlsx","fhfa_data.xlsx","labor_force_data.xls","local_option_sales_data.xlsx","qcew_wklywages 2000 to 2011Q4 VBA Old data.xlsx","qcew_wklywages.xlsx","RGDP_MSA_all.xlsx","Staunton Waynesboro Employment data  Latest.xlsx","Staunton Waynesboro Labor force Latest.xlsx","uiclaims.xls","VA_gdp.xlsx","Washington dc MSA Employment latest.xlsx"]

import os
l = os.listdir("./..")
# print l
for i in l:
    if i in files:
        shutil.copyfile('./../'+i, './files/'+i)


dic = { "Virginia Beach-Norfolk-Newport News, VA-NC (Metropolitan Statistical Area)" : "VirginiaBeach"	,
        "Blacksburg-Christiansburg-Radford, VA (Metropolitan Statistical Area)" : "Blacksburg",
        "Charlottesville, VA (Metropolitan Statistical Area)" : "Charlottesville",
        "Harrisonburg, VA (Metropolitan Statistical Area)" : "Harrisonburg",
        "Lynchburg, VA (Metropolitan Statistical Area)" : "Lynchburg",
        "Richmond, VA (Metropolitan Statistical Area)" : "Richmond",
        "Roanoke, VA (Metropolitan Statistical Area)" : "Roanoke",
        "Staunton-Waynesboro, VA (Metropolitan Statistical Area)" : "Staunton",
        "Washington-Arlington-Alexandria, DC-VA-MD-WV (Metropolitan Statistical Area)" : "Washington",
        "Winchester, VA-WV (Metropolitan Statistical Area)" : "Winchester",
        "Virginia" : "Virginia"}

#-------------------------------------------------------------------------------------------------------------------------------
def RGDP_VA():
    xls = "./files/VA_gdp.xlsx"
    target = "./Csv/VA_gdp.csv"
    try:
        start_time = time.time()
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(0)
        row_count = sh.nrows
        col_count = sh.ncols


        excel=[]
        for k in range(0,row_count):
            row = []
            for m in range(0,col_count):
                a1 = sh.cell_value(rowx=k,colx=m)
                if m==0 and k>=1:
                    a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                    row.append(a1_as_datetime)
                else:
                    row.append(a1)
            excel.append(row)
        excel=zip(*excel)
        #print excel
        del excel[3]
        excel=zip(*excel)
        csvFile = open(target, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    if count==0:
                        try:
                            s = s.strftime('%m/%d/%Y')
                        except:
                            s=str(s)
                        count+=1
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            if count > 1:
                break;
            new.append(newValues)
        #print new
        new[0].append("avagdp")
        new[1].append("0")
	    # print new
        # for k in range(2,len(a1[l])):
        #     if float(a1[l][k])!=-400:
        #         row.append(float(a1[l][k]))


        for i in range(2,len(new)):
            a = ((float(new[i][3])/float(new[i-1][3])) - 1) * 400
            if(a!=-400):
                new[i].append(str(a))
        #new[len(new)-1].append(str(0))
        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            x = new[i][0].split('/')
            flag=0
            if int(x[-1]) >= 2005:
                for k in range(0,len(new[i])):
                    if k!=0:
                        if new[i][k]==0:
                            flag+=1

                if flag!=len(new[i])-1:
                    final.append (new[i])
        #print final
        for item in final:
                wr.writerow(item)

        #print new
        csvFile.close()
        print "RGDP_VA completed"
    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())




def RGDP_MSA_All():
    xls = "./files/RGDP_MSA_all.xlsx"
    target = "./Csv/RGDP.csv"
    try:
        start_time = time.time()
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(1)
        csvFile = open(target, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        for row in xrange(sh.nrows):
            rowValues = sh.row_values(row)
            newValues = []
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue='0'
                newValues.append(str(strValue))
            new.append(newValues)
        new=zip(*new)
        length = len(new)
        final=[]
        for j in range(0,length):
            i= new[j]
            if j==0:
                temp=[]
                for k in i:
                    if k=="GeoName":
                        k="Year"
                    if k in dic:
                        temp.append(dic[k])
                    else:
                        temp.append(k)
                final.append(temp)
            else:
                # print new[j]
                # temp2=[]
                # for mm in new[j]:
                #     # print m
                #     if len(str(mm))==0:
                #         temp2.append('0')
                #     else:
                #         temp2.append(mm)
                final.append(new[j])
        # print final
        for item in final:
            # print item
            if item[0]!="GeoFips" and len(item[0].split())==1:
                wr.writerow(item)
        csvFile.close()
        print "RGDP_MSA completed"
    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())

#-------------------------------------------------------------------------------------------------------------------------------


def Labor_Force_Data_Ssa():
    xls = "./files/labor_force_data.xls"
    target_ssa = "./Csv/labor_force_data.csv"
    flaglfssa = 0
    try:
        start_time = time.time()
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(0)
        row_count = sh.nrows
        col_count = sh.ncols
        excel=[]
        for k in range(0,row_count):
            if flaglfssa==0:
                row = []
                for m in range(0,col_count):
                    a1 = sh.cell_value(rowx=k,colx=m)
                    if m==0 and k>=1:
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                            row.append(a1_as_datetime)
                        except:
                            flaglfssa=1
                    else:
                        row.append(a1)
                if flaglfssa==0:
                    excel.append(row)
        excel=zip(*excel)
        #print excel
        del excel[1]
        del excel[1]
        del excel[-1]
        del excel[5]

        #print excel
        excel=zip(*excel)

        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    if count==0:
                        try:
                            s = s.strftime('%m/%d/%Y')
                        except:
                            s=str(s)
                        count+=1
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            if count > 1:
                break;
            new.append(newValues)

# flag=0
# for k in range(0,len(new[i])):
#
#     if k!=0 and k!=1:
#         if new[i][k]==0:
#             flag+=1

        #print new[0]
        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            x = new[i][0].split('/')
            flag=0
            if int(x[-1]) >= 2005:
                for k in range(0,len(new[i])):
                    if k!=0:
                        if new[i][k]==0:
                            flag+=1

                if flag!=len(new[i])-1:
                    final.append (new[i])
        for item in final:
                wr.writerow(item)
        #print final
        csvFile.close()
        print "LaborForce SSA completed"
    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())

#-------------------------------------------------------------------------------------------------------------------------------


def Labor_Force_Data_Nsa():
    xls = "./files/labor_force_data.xls"
    target_nsa = "./Csv/labor_force_dataNsa.csv"
    xls_staunton = "./files/Staunton Waynesboro Labor force Latest.xlsx"
    wb1 = xlrd.open_workbook(xls_staunton)
    sh1 = wb1.sheet_by_index(0)
    row_count1 = sh1.nrows
    col_count1 = sh1.ncols
    excel1=[]
    startx = 0
    for k in range(0,row_count1):
        for m in range(0,col_count1):
            a1 = sh1.cell_value(rowx=k,colx=m)
            if m==0:
                if a1 == "Year":
                    startx = k+1
    row1 = []
    for k in range(startx,row_count1):
        a1 = sh1.cell_value(rowx=k,colx=5)
        # print a1
        if a1=="":
            a1=0
        row1.append(a1)

    #print "afbfbdfa"


    wb = xlrd.open_workbook(xls)
    sh = wb.sheet_by_index(1)
    row_count = sh.nrows
    col_count = sh.ncols
    excel=[]
    flaglf = 0
    for k in range(0,row_count):
        if flaglf == 0:
            row = []
            for m in range(0,col_count):
               # print k,m,a1
                a1 = sh.cell_value(rowx=k,colx=m)
                #print a1
                #print a1,k,m
                if m==0 and k>=1:
                    try:
                        a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                        row.append(a1_as_datetime)
                        #print a1_as_datetime
                    except:
                        flaglf = 1
                       # print "random data starts"
                else:
                    #print a1
                    row.append(a1)
            if flaglf==0:
                excel.append(row)
    #print "completed"
    #print excel
    excel=zip(*excel)
    #print excel
    del excel[1]
    del excel[5]
    del excel[-1]
    #del excel[5]

    #print excel
    excel=zip(*excel)
    csvFile = open(target_nsa, 'wb')
    wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
    new=[]
    count=0
    row_count = len(excel)
    for row in range(row_count):
        rowValues = excel[row]
        newValues = []
        count=0
        # print rowValues
        for s in rowValues:
            if isinstance(s, unicode):
                strValue = (str(s.encode("utf-8")))
            else:
                if count==0:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        s=str(s)
                    count+=1
                strValue = (str(s))
            isInt = bool(re.match("^([0-9]+)\.0$", strValue))
            if isInt:
                strValue = int(float(strValue))
            else:
                isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                if isFloat:
                    strValue = float(strValue)
                if isLong:
                    strValue = int(float(strValue))
            if strValue=="":
                strValue = 0
            newValues.append(strValue)
        if count > 1:
            break;
        new.append(newValues)
    # print len(new[0])
    new=zip(*new)
    # print len(new[0])
    
    temp = []
    temp.append("staunton")
    for l in row1:
        temp.append(l)
    new.append(temp)
    # print len(new[0])
    new=zip(*new)
    # print len(new[0])
    print new
    final=[]
    length = len(new)
    final.append(new[0])
    for i in range(1,length):
        x = new[i][0].split('/')
        flag=0
        if int(x[-1]) >= 2005:
            for k in range(0,len(new[i])):
                if k!=0:
                    if new[i][k]==0:
                        flag+=1

            if flag!=len(new[i])-1:
                final.append (new[i])
    for item in final:
            wr.writerow(item)
    csvFile.close()
    print "LaborForce NSA completed"


#-------------------------------------------------------------------------------------------------------------------------------



def Employment_Data():
    xls = "./files/employment_data.xlsx"
    xls1 = "./files/Washington dc MSA Employment latest.xlsx"
    target = "./Csv/EMPLOYEMENT_TNFALL_emp_Monthly.csv"
    flag_empssa = 0
    try:
        wb1 = xlrd.open_workbook(xls1)
        sh1 = wb1.sheet_by_index(0)
        row_count1 = sh1.nrows
        col_count1 = sh1.ncols
        excel1=[]
        for k in range(7,row_count1):
            if flag_empssa == 0 :
                row1 = []
                for m in range(0,col_count1):
                    a1 = sh1.cell_value(rowx=k,colx=m)
                    if m==0 and k>=7:
                        #print k,m,a1
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb1.datemode))
                            row1.append(a1_as_datetime)
                        except:
                            flag_empssa = 1
                    else:
                        if a1=="date":
                            row1.append("Date")
                        else:
                            row1.append(a1)
                if flag_empssa == 0:
                    excel1.append(row1)
        
        
        for i in excel1:
            del(i[2])
            del(i[0])
            del(i[0])
            i[0] = float(i[0]) * 1000
        
        #print len(excel1)
        excel1 = zip(*excel1)
        
        flag_empssa = 0
        start_time = time.time()
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(11)
        row_count = sh.nrows
        col_count = sh.ncols
        excel=[]
        for k in range(0,row_count):
            if flag_empssa == 0:
                row = []
                for m in range(0,col_count):
                    a1 = sh.cell_value(rowx=k,colx=m)
                    if m==0 and k>=1:
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                            row.append(a1_as_datetime)
                        except:
                            flag_empssa=1
                    else:
                        if a1=="date":
                            row.append("Date")
                        else:
                            row.append(a1)
                if flag_empssa == 0:
                    excel.append(row)
        # print len(excel)
        excel = zip(*excel)
        del(excel[-2])
        #print excel[0],excel[1]
        #print excel1[0],excel1[1]
            #i.insert(1,'Total Nonfarm NSA')
        excel2 = list(excel1[0])
        excel2.insert(0,'Washington-Arlington-Alexandria, DC-VA-MD-WV MSA, VA part')
        excel1 = tuple(excel2)
        excel.append(excel1)
        # print excel[0],excel[1],len(excel)
        excel = zip(*excel)
        #print len(excel)
        row_count = len(excel)
        csvFile = open(target, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        for row in range(row_count):
            # print row
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    if count==0:
                        try:
                            s = s.strftime('%m/%d/%Y')
                        except:
                            s=str(s)
                    if s in dic:
                        temp=dic[s]
                        s=temp
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                if not (str(strValue)=="Total Nonfarm SA" or str(strValue)=="variable"):
                    newValues.append(strValue)
            if count > 1:
                break;
            new.append(newValues)
        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            x = new[i][0].split('/')
            flag=0
            if int(x[-1]) >= 2005:
                for k in range(0,len(new[i])):
                    if k!=0:
                        if new[i][k]==0:
                            flag+=1

                if flag!=len(new[i])-1:
                    final.append (new[i])
        for item in final:
                wr.writerow(item)
        csvFile.close()
        print "employment_data completed"
    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())

#-------------------------------------------------------------------------------------------------------------------------------


def Employment_Data_Nsa():
    xls = "./files/employment_data.xlsx"
    xls1 = "./files/Washington dc MSA Employment latest.xlsx"
    target = "./Csv/EMPLOYEMENT_TNFALLNSA_emp_Monthly2.csv"
    xls_staunton = "./files/Staunton Waynesboro Employment data  Latest.xlsx"

    wb2 = xlrd.open_workbook(xls1)

    sh2 = wb2.sheet_by_index(0)
    row_count2 = sh2.nrows
    col_count2 = sh2.ncols
    excel2=[]
    flag_empnsa = 0
    for k in range(7,row_count2):
        if flag_empnsa == 0:
            row2 = []
            for m in range(0,col_count2):
                a1 = sh2.cell_value(rowx=k,colx=m)
                if m==0 and k>=7:
                	#print k,m,a1
                    try:
                        a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb2.datemode))
                        row2.append(a1_as_datetime)
                    except:
                        flag_empnsa = 1
                else:
                    if a1=="date":
                        row2.append("Date")
                    else:
                        row2.append(a1)
            if flag_empnsa==0:
                excel2.append(row2)
    for i in excel2:
        del(i[2])
        del(i[0])
        del(i[-1])
        i[0] = float(i[0]) * 1000
    #print excel1
    excel2 = zip(*excel2)

    flag_empnsa=0
    wb1 = xlrd.open_workbook(xls_staunton)
    sh1 = wb1.sheet_by_index(0)
    row_count1 = sh1.nrows
    col_count1 = sh1.ncols
    excel1=[]
    startx = 0
    for k in range(0,row_count1):
        for m in range(0,col_count1):
            a1 = sh1.cell_value(rowx=k,colx=m)
            if m==0:
                if a1 == "Year":
                    startx = k+1
    row1 = []
    for k in range(startx,row_count1):
        for m in range(1,col_count1):
            a1 = sh1.cell_value(rowx=k,colx=m)
            #print k,m,a1
            if a1=="":
                a1=0
            row1.append(float(a1)* 1000)

    wb = xlrd.open_workbook(xls)
    sh = wb.sheet_by_index(12)
    row_count = sh.nrows
    col_count = sh.ncols
    excel=[]
    #print row_count
    for k in range(0,row_count):
        if flag_empnsa==0:
            row = []
            for m in range(0,col_count):
                a1 = sh.cell_value(rowx=k,colx=m)
                if m==0 and k>=1:
                    #print k,m,a1
                	#print a1,k,m, sh.cell_value(rowx=k,colx=m+1),sh.cell_value(rowx=k,colx=m+2),sh.cell_value(rowx=k,colx=m+3)
                    try:
                        a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                        row.append(a1_as_datetime)
                    except:
                        flag_empnsa=1
                else:
                    if a1=="date":
                        row.append("Date")
                    else:
                        row.append(a1)
            if flag_empnsa==0:
                excel.append(row)

    excel = zip(*excel)
    #print excel
    del(excel[1])
    #del(excel[-1])   For extra empty column 
    # del(excel[-1])
    # del(excel[-1])
    del(excel[-2])
    #print excel[]
        #print excel[0],excel[1]
        #print excel1[0],excel1[1]
            #i.insert(1,'Total Nonfarm NSA')
    
    excel3 = list(excel2[0])
    excel3.insert(0,'Washington-Arlington-Alexandria, DC-VA-MD-WV MSA, VA part')
    excel2 = tuple(excel3)
    excel.append(excel2)
    excel = zip(*excel)

    csvFile = open(target, 'wb')
    wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
    new=[]
    count=0
    row_count = len(excel)
    for row in range(row_count):
        rowValues = excel[row]
        newValues = []
        count=0
        for s in rowValues:
            if isinstance(s, unicode):
                strValue = (str(s.encode("utf-8")))
            else:
                try:
                    s = s.strftime('%m/%d/%Y')
                except:
                    s=s
                strValue = (str(s))
            isInt = bool(re.match("^([0-9]+)\.0$", strValue))
            if isInt:
                strValue = int(float(strValue))
            else:
                isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                if isFloat:
                    strValue = float(strValue)
                if isLong:
                    strValue = int(float(strValue))
            if strValue=="":
                strValue = 0
            if not (str(strValue)=="Total Nonfarm NSA" or str(strValue)=="variable"):
                newValues.append(strValue)
        if count > 1:
            break;
        new.append(newValues)
    new=zip(*new)
    temp = []
    temp.append("Staunton-Waynesboro, VA")
    for l in row1:
        temp.append(l)
    new.append(temp)
    new=zip(*new)
    final=[]
    length = len(new)
    final.append(new[0])
    for i in range(1,length):
        x = new[i][0].split('/')
        flag=0
        if int(x[-1]) >= 2005:
            for k in range(0,len(new[i])):
                if k!=0:
                    if new[i][k]==0:
                        flag+=1

            if flag!=len(new[i])-1:
                final.append (new[i])
    for item in final:
            wr.writerow(item)
    csvFile.close()
    print "TNF NSA completed"

############################################################################################

def Lf_Ssa():
    xls = "./files/labor_force_data.xls"
    target_ssa = "./Csv/LABOR_FORCE_lf_ssa_Monthly.csv"
    try:
        start_time = time.time()
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(2)
        row_count = sh.nrows
        col_count = sh.ncols

        flag_lfssa = 0
        excel=[]
        for k in range(0,row_count):
            if flag_lfssa==0:
                row = []
                for m in range(0,col_count):
                    a1 = sh.cell_value(rowx=k,colx=m)
                    if m==0 and k>=1:
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                            #print a1_as_datetime
                            row.append(a1_as_datetime)
                        except:
                            flag_lfssa=1
                    elif m==1 or m==col_count-1:
                        oo=1
                    else:
                        if a1=="date":
                            row.append("Date")
                        else:
                            row.append(a1)
                if flag_lfssa==0:
                    excel.append(row)
        #print excel
        excel=zip(*excel)
        #print excel
        del excel[5]
        #print excel
        excel=zip(*excel)
        #print excel
        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    if count==0:
                        try:
                            s = s.strftime('%m/%d/%Y')
                        except:
                            s=str(s)
                        count+=1
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            if count > 1:
                break;
            new.append(newValues)

        final=[]
        length = len(new)
        final.append(new[0])
        #print new[0]
        for i in range(1,length):
            x = new[i][0].split('/')
            flag=0
            if int(x[-1]) >= 2005:
                for k in range(0,len(new[i])):
                    if k!=0:
                        if new[i][k]==0:
                            flag+=1

                if flag!=len(new[i])-1:
                    final.append (new[i])
        for item in final:
                wr.writerow(item)
        #print final
        csvFile.close()
        print "LF_ssa completed"
    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())

############################################################################################


def Lf_Nsa():
    xls = "./files/labor_force_data.xls"
    target_ssa = "./Csv/LABOR_FORCE_lf_nsa_Monthly2.csv"
    xls_staunton = "./files/Staunton Waynesboro Labor force Latest.xlsx"
    wb1 = xlrd.open_workbook(xls_staunton)
    sh1 = wb1.sheet_by_index(0)
    row_count1 = sh1.nrows
    col_count1 = sh1.ncols
    excel1=[]
    startx = 0
    for k in range(0,row_count1):
        for m in range(0,col_count1):
            a1 = sh1.cell_value(rowx=k,colx=m)
            if m==0:
                if a1 == "Year":
                    startx = k+1
    row1 = []
    for k in range(startx,row_count1):
        a1 = sh1.cell_value(rowx=k,colx=2)
        if a1=="":
            a1=0
        row1.append(float(a1))

    wb = xlrd.open_workbook(xls)
    sh = wb.sheet_by_index(3)
    row_count = sh.nrows
    col_count = sh.ncols

    excel=[]
    flag_lfnsa=0
    for k in range(0,row_count):
        if flag_lfnsa==0:
            row = []
            for m in range(0,col_count):
                a1 = sh.cell_value(rowx=k,colx=m)
                if m==0 and k>=1:
                    try:
                        a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                        #print a1_as_datetime
                        row.append(a1_as_datetime)
                    except:
                        flag_lfnsa=1
                elif m==1 or m==col_count-1:
                    oo=1
                else:
                    if a1=="date":
                        row.append("Date")
                    else:
                        row.append(a1)
            if flag_lfnsa==0:
                excel.append(row)

    excel=zip(*excel)
    # print excel
    del excel[5]
    #print excel
    excel=zip(*excel)
    csvFile = open(target_ssa, 'wb')
    wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
    new=[]
    count=0
    row_count=len(excel)
    for row in range(row_count):
        rowValues = excel[row]
        newValues = []
        count=0
        for s in rowValues:
            if isinstance(s, unicode):
                strValue = (str(s.encode("utf-8")))
            else:
                if count==0:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        s=str(s)
                    count+=1
                strValue = (str(s))
            isInt = bool(re.match("^([0-9]+)\.0$", strValue))
            if isInt:
                strValue = int(float(strValue))
            else:
                isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                if isFloat:
                    strValue = float(strValue)
                if isLong:
                    strValue = int(float(strValue))
            if strValue=="":
                strValue = 0
            newValues.append(strValue)
        if count > 1:
            break;
        new.append(newValues)

    new=zip(*new)
    temp = []
    temp.append("staunton")
    for l in row1:
        temp.append(l)
    new.append(temp)
    new=zip(*new)
    final=[]
    length = len(new)
    final.append(new[0])
    for i in range(1,length):
        x = new[i][0].split('/')
        flag=0
        if int(x[-1]) >= 2005:
            for k in range(0,len(new[i])):
                if k!=0:
                    if new[i][k]==0:
                        flag+=1

            if flag!=len(new[i])-1:
                final.append (new[i])
    for item in final:
            wr.writerow(item)
    csvFile.close()
    print "LF_Nsa completed"

############################################################################################

def Washington_MSA():
    xls = "./files/Washington dc MSA Employment latest.xlsx"
    target = "./Csv/Washington dc MSA Employment latest.csv"


    csvFile = open(target, 'wb')
    wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)

    wb1 = xlrd.open_workbook(xls)
    sh1 = wb1.sheet_by_index(0)
    row_count1 = sh1.nrows
    col_count1 = sh1.ncols
    excel1=[]
    startx = 0
    for k in range(0,row_count1):
        for m in range(0,col_count1):
            a1 = sh1.cell_value(rowx=k,colx=m)
            if m==0:
                if a1 == "date":
                    startx = k+1
    excel=[]
    flag_DCmsa = 0
    for k in range(startx,row_count1):
        if flag_DCmsa==0:
            row = []
            for m in range(0,col_count1):
                a1 = sh1.cell_value(rowx=k,colx=m)
                if m==0 :
                    #print a1
                    try:
                        a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb1.datemode))
                        #print a1_as_datetime
                        row.append(a1_as_datetime)
                    except:
                        flag_DCmsa=1
                elif m==2:
                    oo=1
                else:
                    row.append(a1)
            if flag_DCmsa==0:
                excel.append(row)

    # print excel

    new=[]
    count=0
    row_count1 = len(excel)
    for row in range(0,row_count1):
        rowValues = excel[row]
        newValues = []
        count=0
        for s in rowValues:
            if isinstance(s, unicode):
                strValue = (str(s.encode("utf-8")))
            else:
                if count==0:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        s=str(s)
                    count+=1
                strValue = (str(s))
            isInt = bool(re.match("^([0-9]+)\.0$", strValue))
            if isInt:
                strValue = int(float(strValue))
            else:
                isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                if isFloat:
                    strValue = float(strValue)
                if isLong:
                    strValue = int(float(strValue))
            if strValue=="":
                strValue = 0
            newValues.append(strValue)
        if count > 1:
            break;
        new.append(newValues)

    final=[]
    length = len(new)
    x=["date","Nsa_value","Ssa_value"]
    final.append(x)
    for i in range(1,length):
        x = new[i][0].split('/')
        flag=0
        if int(x[-1]) >= 2005:
            for k in range(0,len(new[i])):
                if k!=0:
                    if new[i][k]==0:
                        flag+=1

            if flag!=len(new[i])-1:
                final.append (new[i])
    for item in final:
            wr.writerow(item)
    csvFile.close()
    print "Washington MSA completed"

######################################################################################

def Weekly_wages_quarterly():
    xls = "./files/qcew_wklywages.xlsx"
    xls1 = "./files/qcew_wklywages 2000 to 2011Q4 VBA Old data.xlsx"

    target_ssa = "./Csv/WEEKLY_WAGES_Quarterly.csv"
    try:
        wb1 = xlrd.open_workbook(xls1)
        sh1 = wb1.sheet_by_index(0)
        row_count1 = sh1.nrows
        col_count1 = sh1.ncols
        excel=[]
        dic1 = [3,4,5,6,7,8,9,10,11,12,13,15]
        for k in range(0,row_count1):
            row = []
            for m in range(0,col_count1):
                a1 = sh1.cell_value(rowx=k,colx=m)
                if m in dic1:
                    if a1 == "date.x":
                        row.append("Date")
                    else:
                        row.append(a1)
            excel.append(row)


        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(0)
        row_count = sh.nrows
        col_count = sh.ncols
        dic = [2,3,4,5,6,7,8,9,10,11,12,13]
        
        for k in range(1,row_count):
            row = []
            for m in range(0,col_count):
                a1 = sh.cell_value(rowx=k,colx=m)
                if m in dic:
                    if a1 == "date.x":
                        row.append("Date")
                    else:
                        row.append(a1)
            excel.append(row)

        #print excel
        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        for row in range(row_count+row_count1-1):
            rowValues = excel[row]
            newValues = []
            count=0
            #print rowValues
            for s in rowValues:
                #print s
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                strValue = str(s)
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
                #print newValues
            new.append(newValues)
        count = 0
        # print new
        for item in new:
            if count >=1:
                #print item
                if int(item[0].split(' ')[0])>=2005:
                        wr.writerow(item)
            else:
                wr.writerow(item)
            count+= 1
        csvFile.close()
        print "Weekly wages quarterly"
    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())


######################################################################################

def Ui_claims():
    xls = "./files/uiclaims.xls"
    target_ssa = "./Csv/uiclaims.csv"
    try:
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(1)
        row_count = sh.nrows
        col_count = sh.ncols
        excel=[]
        flag_uiclaims = 0
        for k in range(0,row_count):
            row = []
            if flag_uiclaims==0:
                for m in range(0,col_count):
                    a1 = sh.cell_value(rowx=k,colx=m)
                    if m==0 and k>=1:
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                            row.append(a1_as_datetime)
                        except:
                            flag_uiclaims=1
                    else:
                        row.append(a1)
                if flag_uiclaims==0:
                    excel.append(row)
        excel=zip(*excel)

        del excel[4]
        del excel[4]
        del excel[9]
        del excel[10]
        del excel[11]
        excel=zip(*excel)

        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        row_count = len(excel)
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        strValue = (str(s))
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            new.append(newValues)

        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            final.append(new[i])
        for item in final:
                wr.writerow(item)
        csvFile.close()
        print "UI claims completed"

    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())


######################################################################################

def local_option_sales_data():
    xls = "./files/local_option_sales_data.xlsx"
    target_ssa = "./Csv/local_option_sales_data.csv"
    try:
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(0)
        row_count = sh.nrows
        col_count = sh.ncols
        excel=[]
        flag_localoptions = 0
        for k in range(0,row_count):
            row = []
            if flag_localoptions==0:
                for m in range(0,col_count):
                    a1 = sh.cell_value(rowx=k,colx=m)
                    if m==0 and k>=1:
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                            row.append(a1_as_datetime)
                        except:
                            flag_localoptions=1
                    else:
                        row.append(a1)
                if flag_localoptions==0:
                    excel.append(row)

        excel=zip(*excel)

        del excel[4]
        del excel[4]

        # del excel[4]
        # del excel[9]
        # del excel[10]
        # del excel[11]
        del excel[-1]
        del excel[-1]

        del excel[-2]
        del excel[-3]
        #del excel[8]
        excel=zip(*excel)

        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        row_count = len(excel)
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        strValue = (str(s))
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            new.append(newValues)

        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            x = new[i][0].split('/')
            flag=0
            if int(x[-1]) >= 2005:
                for k in range(0,len(new[i])):
                    if k!=0:
                        if new[i][k]==0:
                            flag+=1

                if flag!=len(new[i])-1:
                    final.append (new[i])
        for item in final:
                wr.writerow(item)
        csvFile.close()
        print "local options sales data completed"

    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())


######################################################################################
def FHFA_Quarterly():
    xls = "./files/fhfa_data.xlsx"
    target_ssa = "./Csv/FHFA_Quarterly.csv"
    try:
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(0)
        row_count = sh.nrows
        col_count = sh.ncols
        excel=[]
        flag_fhfa=0
        for k in range(0,row_count):
            row = []
            if flag_fhfa==0:
                for m in range(0,col_count):
                    a1 = sh.cell_value(rowx=k,colx=m)
                    if m==0 and k>=1:
                        try:
                            a1_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(a1, wb.datemode))
                            row.append(a1_as_datetime)
                        except:
                            flag_fhfa=1
                    else:
                        row.append(a1)
                if flag_fhfa==0:
                    excel.append(row)
        excel=zip(*excel)

        del excel[4]
        del excel[-2]
        excel=zip(*excel)
        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        row_count = len(excel)
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        strValue = (str(s))
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            new.append(newValues)

        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            x = new[i][0].split('/')
            flag=0
            if int(x[-1]) >= 2005:
                for k in range(0,len(new[i])):
                    if k!=0:
                        if new[i][k]==0:
                            flag+=1

                if flag!=len(new[i])-1:
                    final.append (new[i])
        for item in final:
                wr.writerow(item)
        csvFile.close()
        print "FHFA quarterly completed"

    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())

######################################################################################



def Revpar():
    xls = "./files/Dashboard_REVPAR_Month.xlsx"
    target_ssa = "./Csv/Dashboard_REVPAR_Month1.csv"

    dic={"Washington DC-MD-VA Market":"WashingtonDC-MD-VAMarket",
    	 "Richmond/Petersburg Market":"Richmond/PetersburgMarket",
         "Staunton/Harrisonburg Market":"Staunton/HarrisonburgMarket",
         "Virginia Portion of Washington DC":"VirginiaPortionofWashingtonDC",
         "Blacksburg/Wytheville Market":"Blacksburg/WythevilleMarket",
         "the Commonwealth of Virginia":"theCommonwealthofVirginia",
         "Charlottesville Market":"CharlottesvilleMarket",
         "Lynchburg Market":"LynchburgMarket",
         "Roanoke Market":"RoanokeMarket",
         "Hampton Roads Market":"HamptonRoadsMarket"}
    try:
        wb = xlrd.open_workbook(xls)
        sh = wb.sheet_by_index(0)
        row_count = sh.nrows
        col_count = sh.ncols
        excel=[]
        for k in range(0,row_count):
            row = []
            for m in range(0,col_count):
                a1 = sh.cell_value(rowx=k,colx=m)
                if a1 in dic:
                    row.append(dic[a1])
                else:
                    row.append(a1)
            excel.append(row)
        print excel[-1]
        csvFile = open(target_ssa, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
        new=[]
        count=0
        for row in range(row_count):
            rowValues = excel[row]
            newValues = []
            count=0
            for s in rowValues:
                if isinstance(s, unicode):
                    strValue = (str(s.encode("utf-8")))
                else:
                    try:
                        s = s.strftime('%m/%d/%Y')
                    except:
                        strValue = (str(s))
                    strValue = (str(s))
                isInt = bool(re.match("^([0-9]+)\.0$", strValue))
                if isInt:
                    strValue = int(float(strValue))
                else:
                    isFloat = bool(re.match("^([0-9]+)\.([0-9]+)$", strValue))
                    isLong  = bool(re.match("^([0-9]+)\.([0-9]+)e\+([0-9]+)$", strValue))
                    if isFloat:
                        strValue = float(strValue)
                    if isLong:
                        strValue = int(float(strValue))
                if strValue=="":
                    strValue = 0
                newValues.append(strValue)
            new.append(newValues)

        final=[]
        length = len(new)
        final.append(new[0])
        for i in range(1,length):
            flag=0
            for k in range(0,len(new[i])):                        

                if k!=0 and k!=1:
                    if new[i][k]==0:
                        flag+=1

            if flag!=len(new[i])-2:
                if int(new[i][0] ) >= 2005:
                    final.append (new[i])
        for item in final:
                wr.writerow(item)
        csvFile.close()
        print "REVPAR completed"

    except Exception as e:
        print (str(e) + " " +  traceback.format_exc())

######################################################################################

RGDP_MSA_All() # RGDP_MSA_All.csv
Labor_Force_Data_Nsa() # unemprate_nsa.csv
Labor_Force_Data_Ssa()  # unemprate_ssa.csv
Employment_Data()   # TN_Fall_Monthly.csv
Employment_Data_Nsa() # TN_Fall_Monthly_Nsa.csv
Lf_Ssa()    #lf_ssa.csv
Lf_Nsa()    #lf_nsa.csv
Washington_MSA() # Washington_MSA.csv
Weekly_wages_quarterly() # Weekly_wages_quarterly.csv
Ui_claims() #Ui_claims.csv
local_option_sales_data() # local_option_sales_data.csv
FHFA_Quarterly() # FHFA_Quarterly.csv
Revpar() # Revpar.csv
RGDP_VA()
