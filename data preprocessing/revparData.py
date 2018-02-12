import logging
import time
import traceback
import xlrd
import csv
import sys
import re
import datetime
import os

xlnames={}

xlnames["Country United States"]="United States"
xlnames["State of Virginia"]="the Commonwealth of Virginia"
xlnames["Market Norfolk-Virginia Beach, VA"]="Hampton Roads Market"
xlnames["Market Myrtle Beach, SC"]="Myrtle Beach, SC"
xlnames["Tract Chesapeake/Ocean City"]="Ocean City, MD"
xlnames["Tract Virginia Beach"]="Virginia Beach Market"
xlnames["Tract Norfolk/ Portsmouth"]="Norfolk/Portsmouth Market"
xlnames["Tract Williamsburg"]="Williamsburg Market"
xlnames["Tract Chesapeake/ Suffolk"]="Chesapeake/Suffolk Market"
xlnames["Tract Newport News/Hampton"]="Newport News/Hampton Market"
xlnames["City of Norfolk, VA "]="City of Norfolk"
xlnames["City of Chesapeake, VA  "]="City of Chesapeake"
xlnames["City of Suffolk, VA  "]="City of Suffolk"
xlnames["City of Hampton, VA  "]="City of Hampton"
xlnames["City of Newport News, VA  "]="City of Newport News"
xlnames["Tract Coastal Carolina, NC"]="Coastal Carolina"
xlnames["Gates, Pasquotank, Camden, and Currituck NC Counties"]="Camden, Currituck, Gates, and Pasquotank"
xlnames["Tract Blacksburg/Wytheville, VA"]="Blacksburg/Wytheville Market"
xlnames["Tract Charlottesville, VA"]="Charlottesville Market"
xlnames["Tract Bristol/Kingsport, TN"]="Bristol/Kingsport, TN Market"
xlnames["Tract Lynchburg,VA"]="Lynchburg Market"
xlnames["Market Richmond/Petersburg, VA"]="Richmond/Petersburg Market"
xlnames["Tract Roanoke, VA"]="Roanoke Market"
xlnames["Tract Staunton/Harrisonburg, VA"]="Staunton/Harrisonburg Market"
xlnames["Market Washington, DC-MD-VA"]="Washington, DC-MD-VA Market"
xlnames["Virginia portion of Tract   Bristol/Kingsport, TN"]="Virginia Portion of Bristol/Kingsport"
xlnames["Virginia Portion of Market  Washington, DC-MD-VA"]="Virginia Portion of Washington DC"
xlnames["Tract Chesapeake/Ocean City, MD"]="Chesapeake/Ocean City, MD"
xlnames["Washington DC CBD (District of Columbia)"]="DC CBD"
xlnames["Maryland Portion of Washington, DC-MD-VA"]="Maryland Portion"

array=["the Commonwealth of Virginia","Blacksburg/Wytheville Market","Virginia Portion of Washington DC","Richmond/Petersburg Market","Washington, DC-MD-VA Market","Lynchburg Market","Staunton/Harrisonburg Market",
		"Charlottesville Market","Roanoke Market","Hampton Roads Market"]


naming={"Washington, DC-MD-VA Market":"WashingtonDC-MD-VAMarket"	,"Richmond/Petersburg Market":"Richmond/PetersburgMarket","Staunton/Harrisonburg Market":	"Staunton/HarrisonburgMarket",
	"Virginia Portion of Washington DC":"VirginiaPortionofWashingtonDC"	,"Blacksburg/Wytheville Market":"Blacksburg/WythevilleMarket",
	"the Commonwealth of Virginia":"theCommonwealthofVirginia"	,"Charlottesville Market":"CharlottesvilleMarket",
		"Lynchburg Market":"LynchburgMarket"	,
		"Roanoke Market":"RoanokeMarket",
			"Hampton Roads Market":"HamptonRoadsMarket"
}


#print os.listdir("./../../../../Hotel fILES/Strome Business Data/")

def REVPAR():
	finalarray=[]
	columnFlag=0
	for filename in os.listdir("./../../../../Hotel fILES/Strome Business Data/"):
		if filename!= "CPIU.xls" and filename!="header.png" and filename!="TopBanner.png" and filename!="Thumbs.db":
                        print filename
			xls = "./../../../../Hotel fILES/Strome Business Data/"+ filename
			target = "./Csv/Dashboard_REVPAR_Month1.csv"
			#print xls
			wb = xlrd.open_workbook(xls)
			sh = wb.sheet_by_index(1)
			cityName = sh.cell_value(rowx=1,colx=1)
			year=[]
			months=[]
			if cityName in xlnames:
				#print cityName,xlnames[cityName]
				if xlnames[cityName] in array:
					#print xlnames[cityName]
					temp=[]
					temp.append(naming[xlnames[cityName]])
					row_count=sh.nrows
					col_count = sh.ncols
					flag=0
					start=0
					end=0
					for i in range(0,row_count):
						if sh.cell_value(rowx=i,colx=1) == "RevPAR ($)":
							flag=1
						elif sh.cell_value(rowx=i,colx=1) == 2005 and flag==1:
							start=i
						elif sh.cell_value(rowx=i,colx=1) == "Avg" and flag==1:
							end = i-1
							break;
					if columnFlag==0:
						a=2005
						columnFlag=1
						year.append("Year")
						for i in range(start,end+1):
							for j in range(1,13):
								year.append(a)
							a+=1
						dic={1:"Jan",2:"Feb",3:"Mar",4:"Apr",5:"May",6:"Jun",7:"Jul",8:"Aug",9:"Sep",10:"Oct",11:"Nov",12:"Dec"}
						months.append("Month")
						for i in range(start,end+1):
							for j in range(1,13):
								months.append(dic[j])
						finalarray.append(year)
						finalarray.append(months)
					#print start,end
					for i in range(start,end+1):
						for j in range(2,14):
							temp.append(sh.cell_value(rowx=i,colx=j))
					finalarray.append(temp)
	
	finalarray=zip(*finalarray)	
	csvFile = open(target, 'wb')
        wr = csv.writer(csvFile, quoting=csv.QUOTE_ALL)
	for item in finalarray:
                wr.writerow(item)			
REVPAR()
