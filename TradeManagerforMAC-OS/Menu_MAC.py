###################################
#Swing Capital Trading Manager

#Written by: 
#Mehdi Ezatabadi
#Parth Singh
#Sanjit Singh Batra
#Version: 1 Dec 2017
###################################


import csv
import sys
from shutil import copyfile
import os
from Tkinter import *
import pprint
from tabulate import tabulate
import pandas as pd
import xlsxwriter 
from datetime import datetime


#We hard code the path to two input files which this framework modifies
#PLEASE MAKE SURE THIS IS SET TO YOUR SPECIFIC MACHINE
cwd = os.getcwd()
Investors=cwd+"/Investors.csv"
File_xlsx=cwd+"/Trades.xlsx"
#Investors="/Users/sanjitsbatra/Desktop/Chaabi/Menu/Nov17/Investors.csv"
#File_xlsx="/Users/sanjitsbatra/Desktop/Chaabi/Menu/Nov17/Trades.xlsx"
#Investors="/Users/mezzatab/Dropbox/SwingCapital/Trading_Manager/Nov_13/Investors.csv"
#File_xlsx="/Users/mezzatab/Dropbox/SwingCapital/Trading_Manager/Nov_13/Trades.xlsx"








######################################################################################################
#This function converts the monthly report generated in text format into a pretty xlsx format
######################################################################################################

def txtToxlsx(inputFile, outputfile):

##  Defining the excel file ## 
    workbook = xlsxwriter.Workbook(outputfile)
    worksheet = workbook.add_worksheet()



## Setting up the excel file ## 


# the structure parameter #

    hRow=32 # the row where  header titles starts # 
    sCol=9 # the number of sepration column, trade information || investors gain and loose 
    maxBlank=100 # the number of rows in specific cols to be set to be white empty # like the height of gap between each investor 
    invCols=7 # nCol needed for each investor for each trade 
#    nInvestors=3 # number of investors # ?
#nameInvestors=['Fabian Ng','Parth Singh','MM Ezzat'] # List of the name of investors # ?
# End of structure paramters #

#####                               ###############
    worksheet.insert_image('A1', 'Logo.png',{'x_offset': 8, 'y_offset': 4,'x_scale':0.5,'y_scale':0.5})
#    worksheet.insert_image('D1', 'FN4X.png',{'x_offset': 15, 'y_offset': 10,'x_scale':1,'y_scale':1})


############################
# defining the format of the cells 
# the green P/L pips # 
    Pl_format=workbook.add_format({'bg_color':'#CCFFCC','align':'center','border':1})

## Leave the top cells blank 
    blankFormat = workbook.add_format({'bg_color': '#FFFFFF'})

# Add a format for headers # 
    headerFormat = workbook.add_format({'bold': 1,'bg_color': '#FFCC99','align':'center','border':2})


# Casual format #
    casFormat=workbook.add_format({'align':'center','border':1})

# the border of gap 
    cgFormat=workbook.add_format({'bg_color': '#FFFFFF','left':1,'right':1,'border':1})
#cgFormat.

# the border of the very end gap 
    cgLastFormat=workbook.add_format({'bg_color': '#FFFFFF','right':1,'border':1})
#cgFormat.

#   #

    investFormat = workbook.add_format({'bold': 1,'bg_color': '#CCFFFF','align':'center','border':2})
    invNameFormat = workbook.add_format({'bold': 10,'bg_color': '#CCFFFF','align':'center','border':2})




#################

 # Add a number format for cells with money.
    money_format = workbook.add_format({'num_format': '$#,##0.00','align':'center','border':1})
    Smoney_format = workbook.add_format({'num_format': '$#,##0.00','align':'center','bg_color':'yellow','border':2})

# Add number format for open price, close price , numbers with no $ sign 
    number_format = workbook.add_format({'num_format': '0.00','align':'center','border':1})
    format_oc_price = workbook.add_format({'num_format': '0.00','align':'center','border':1})    

# Add number format for share %, numbers with no $ sign 
    numPer_format = workbook.add_format({'num_format': '0.00\%','align':'center','border':1})
    SnumPer_format = workbook.add_format({'num_format': '0.00\%','align':'center','bg_color':'yellow','border':2})

# numbers in P/L pips cells with green color 
    num_pips_format=workbook.add_format({'num_format': '0.00','align':'center','bg_color':'#CCFFCC', 'border':1})
 # Add an Excel date format.
    date_format = workbook.add_format({'num_format': 'mmmm d yyyy','border':1})

## Summary first columns format ## 
    tFormat=workbook.add_format({'bold': 10,'bg_color': '#CCFFFF','align':'center','border':2})


## Merge cells ## 
        ## Merge cells format for names and investors #
    merge_Name = workbook.add_format({
    'bold':     25,
    'border':   6,
    'align':    'center',
    'valign':   'vcenter',
    'fg_color': '#D7E4BC',
})


    ## Text Box options ## 
    option1 = {
    'width': 256,
    'height': 100,
    'x_offset': 1,
    'y_offset': 1,
    'font': {'size': 15},

#    'font': {'color': '#CCFFFF',
#             'size': 15},
    'align': {'vertical': 'middle',
              'horizontal': 'center'
              },
    'gradient': {'colors': ['#DDEBCF',
                            '#9CB86E',
                            '#156B13']},
}

    option2 = {
    'width': 1000,
    'height': 40,
    'x_offset': 1,
    'y_offset': 1,
    'font': {'size': 15},

#    'font': {'color': '#CCFFFF',
#             'size': 15},
    'align': {'vertical': 'middle',
              'horizontal': 'center'
              },
    'fill': {'color': '#FF9900'}#    'gradient': {'colors': ['#DDEBCF',
 #                           '#9CB86E',
 #                           '#156B13']},
}

    option3 = {
    'width': 500,
    'height': 60,
    'x_offset': 1,
    'y_offset': 1,
    'font': {'size': 15},

#    'font': {'color': '#CCFFFF',
#             'size': 15},
    'align': {'vertical': 'middle',
              'horizontal': 'center'
              },
    'fill': {'color': 'green'}#    'gradient': {'colors': ['#DDEBCF',
 #                           '#9CB86E',
 #                           '#156B13']},
}


# the width of columns # 
    worksheet.set_column(2, 2, 20) # Date open size
    worksheet.set_column(3, 3, 15) # open price
    worksheet.set_column(5, 5, 15) # Close price
    worksheet.set_column(4, 4, 20) # date close size
    worksheet.set_column(5, 5, 15) # date close size
    worksheet.set_column(9, 9, 5) # date close size
    worksheet.set_column(8, 8, 25) # Date open size
    worksheet.set_column(13, 13, 25) # Date open size    
    worksheet.set_column(6, 6, 15) # Date open size      
    worksheet.set_column(7, 7, 15) # Date open size            
    for i in range(9,13):
        worksheet.set_column(i, i, 20) # Date open size

###########################################

# Blanking empty cells #
# Blanking the white empty cells 
    for i in range(3*hRow):
       for j in range(2*sCol):
            worksheet.write(i,j,None,blankFormat)


## End of decorating excel file ###




    f = open(inputFile)
## Header line handler ## 
    context=f.readlines()
    headerLine=context[0]
    List= headerLine.strip().split("\t") 
    Name=List[1].strip()
    headerLine=context[1]
    List= headerLine.strip().split("\t")     
    Month=List[1].strip()
    headerLine=context[2]
    List= headerLine.strip().split("\t") 
    year='20'+List[1].strip()
    Date=Month +'  of  '+year;
    text='Investor: ' +Name+'\n'+ 'Month: '+Month +', Year: '+year
    ## Writing the name and the date at the top of report ## 
## Merge cells ## 

    worksheet.insert_textbox(hRow-12,0, text, option1)
    openTrades='Trades open as of end of this month ( ' +Date+ ' ) , that '+ Name+ ' is part of:'
    worksheet.insert_textbox(hRow-3,0, openTrades, option2)

#    worksheet.merge_range(hRow-10,0,hRow-8,3, Name, merge_Name)        
#    worksheet.merge_range(hRow-7,0,hRow-5,3, Date, merge_Name)        
#    print year
#    print Name
## Empty lines ##
    line=3
    while context[line]=='\n':
        line=line+1
#    print line 
## Open Trades ## 
#    oTradeText=context[line]
#   line=line+1
#    print oTradeText
	headers=context[line].strip().split("\t")   ## Tab seprated ## 
#    print headers
#    print headers
#    headers=['Currency',    'B/S',     'Date Open',   'Open Price',  'Date close',  'Close Price',  'P/L Pips',  'P/L ($)',  'Total Pips',   
#     'Value for opening trade',    '% Share at open',    'P/L ($)',   'Commission',      'Added Funds',    'Value after closing trade',    '% Share at close']
    line=line+1
     # Reading values of each trade#

### Setting up the header in excel ##

# Trade section #
    for i in range(len(headers)):
        worksheet.write_string(hRow, i,headers[i],headerFormat)


####


## Reading txt and writing  xlsx Open Trades Data ## 
## @@ ##
    xLine=0
    while context[line] != '\n':
        List= context[line].strip().split("\t") 
        ## Writing ## 
        for i in range(len(List)): 
            if List[i]!='':           
                worksheet.write_string(hRow+line-4, i, List[i], casFormat) # Currency                            
                if i==6:
                    worksheet.write_number(hRow+line-4, i, float(List[i]), number_format) # open close price                                 
                if i==3 or i==5:
                    worksheet.write_number(hRow+line-4, i, float(List[i]), format_oc_price) # open close price                 
                if i==6:
                   worksheet.write_number(hRow+line-4, i, float(List[i]), num_pips_format) # Pl_Format
                if i==7 or i==8:
                    worksheet.write_number(hRow+line-4, i, float(List[i]), money_format) # $ format
                if i==9:
                    worksheet.write_number(hRow+line-4, i, float(List[i]), numPer_format) # % format  
                if 9<i<=13:               
                    worksheet.write_number(hRow+line-4, i, float(List[i]), money_format) # $ format
#        print List
            else:
                worksheet.write(hRow+line-4, i, None, casFormat)         
        for i in range(9,14):
                worksheet.write(hRow+line-4, i, None, casFormat)         

        line=line+1
## @@ ##






## Skipping the empty lines ## 

    while context[line]=='\n':
        line=line+1

## Writing close trades header 
    closeTrades='Trades closed during this month, that '+Name+' was part of:'
    worksheet.insert_textbox(hRow+line-3,0,closeTrades,option2)


    cTradeText=context[line]
#    line=line+1

    headsUp=context[line].strip().split("\t")   ## Tab seprated ## 
    for i in range(len(headsUp)):
        worksheet.write_string(hRow+line-1, i,headers[i],headerFormat)

#    print headers
    line=line+1
     # Reading values of each trade#



    while context[line] != '\n':
        List= context[line].strip().split("\t") 
        for i in range(len(List)):
            if List[i]!='':           
                worksheet.write_string(hRow+line-1, i, List[i], casFormat) # Currency                            
                if i==6:
                    worksheet.write_number(hRow+line-1, i, float(List[i]), number_format) # open close price                                 
                if i==3 or i==5:
                    worksheet.write_number(hRow+line-1, i, float(List[i]), format_oc_price ) # open close price                 
                if i==6:
                   worksheet.write_number(hRow+line-1, i, float(List[i]), num_pips_format) # Pl_Format
                if i==7 or i==8:
                    worksheet.write_number(hRow+line-1, i, float(List[i]), money_format) # $ format
                if i==9:
                    worksheet.write_number(hRow+line-1, i, float(List[i]), numPer_format) # % format  
                if 9<i<=13:               
                    worksheet.write_number(hRow+line-1, i, float(List[i]), money_format) # $ format
            else:
                worksheet.write(hRow+line-1, i, None, casFormat) # Currency                            

#        print List
#        print List
        line=line+1

    while context[line]=='\n':
        line=line+1
    List= context[line].strip().split("\t")     
    Summary=context[line]    
    line=line+1
    worksheet.set_column(sCol, sCol, 80) # Date open size 
    worksheet.insert_textbox(1,sCol,Summary,option3)   
    row=5
    while context[line] != '\n':
        List= context[line].strip().split("\t") 
        worksheet.write_string(row,sCol,List[0],tFormat)
        x,y=divmod(row,2)
        if y==1:
            worksheet.write_number(row,sCol+1, float(List[1]), Smoney_format) # $ format     
        if y==0:   
            worksheet.write_number(row,sCol+1, float(List[1]), SnumPer_format) # $ format                 
        Summary=Summary+context[line]+'\n'
        line=line+1
        row=row+1




#    print(headers)
#    print context[line]
#    print Month
    f.close()
    workbook.close()





































#################################################################################################################
#This function converts the Trading csv file used as the intermediate format, into a well-formatted xlsx file
#################################################################################################################


# inputFile 1 : the csv file having trade information
# inputFile 2: csv file having investors information , reading their names 
# outputFile: xlsx output file 
def csvToxlsx(inputFile1, inputFile2, outputFile):
# the structure parameter #

    hRow=10 # the row where  header titles starts # 
    sCol=9 # the number of sepration column, trade information || investors gain and loose 
    maxBlank=100 # the number of rows in specific cols to be set to be white empty # like the height of gap between each investor 
    invCols=7 # nCol needed for each investor for each trade 
#    nInvestors=3 # number of investors # ?
#nameInvestors=['Fabian Ng','Parth Singh','MM Ezzat'] # List of the name of investors # ?
# End of structure paramters #


# reading the trade data and inveestors track versus trades 
    reader = csv.reader(open(inputFile1, 'rU'))
    header = reader.next()


    sepIndex=header.index('Total Pips')


#Trade is reading the information of csv file # 
    trade=[]
    for row in reader:
        trade.append(row)

    nRows=len(trade) # N trades
#print len(trade)+1

# End of reading trades date 



# Reading the investors original data file #

    invReader = csv.reader(open(inputFile2, 'rU'))
    dummy = invReader.next()

#Trade is reading the information of csv file # 
    listInv=[]
    nameInvestors=[]
    for row in invReader:
        listInv.append(row)
        nameInvestors.append(row[0])
    nInvestors=len(listInv) # N trades




# End of knowing investors #

# Start forming the structure of xlsx file #
# Headers titles #

    Title_Dict=["Date Open","Currency", 'B/S', 'Open Price', 'Date Close', 'Close Price', 'P/L pips', 'P/L ($)', 'Total Pips']
    tradeHeads=len(Title_Dict)
# Headers for columns of investors information #
    Inv_Dict=["Open Trade Investment ($)","Open Trade Share (%)", 'Trade P/L ($)', 'Commission ($)', 'Amount Changed ($)', 'Closed Trade Investment ($)', 'Closed Trade Share (%)']
    invHeads=len(Inv_Dict)


    workbook = xlsxwriter.Workbook(outputFile)
    worksheet = workbook.add_worksheet()


# the width of columns # 
    worksheet.set_column(0, 0, 20) # Date open size
    worksheet.set_column(3, 3, 15) # Close price
    worksheet.set_column(4, 4, 20) # date close size
    worksheet.set_column(5, 5, 15) # date close size
    worksheet.set_column(6, 6, 20) # date close size
    worksheet.set_column(7, 7, 20) # date close size
    worksheet.set_column(8, 8, 20) # date close size
    worksheet.set_column(9, 9, 5) # date close size

# width of Investor section cols # 
    for j in range(nInvestors):
        for i in range(sCol+1, sCol+invHeads+1):
            worksheet.set_column(sCol+1+j+j*invHeads,sCol+j+j*invHeads+invHeads,22)
        worksheet.set_column(sCol+j+j*invHeads,sCol+j+j*invHeads,5) # Gap between investors         


# defining the format of the cells 
# the green P/L pips # 
    Pl_format=workbook.add_format({'bg_color':'#CCFFCC','border':1})

## Leave the top cells blank 
    blankFormat = workbook.add_format({'bg_color': '#FFFFFF'})

# Add a format for headers # 
    headerFormat = workbook.add_format({'bold': 1,'bg_color': '#FFCC99','align':'center','border':2})


# Casual format #
    casFormat=workbook.add_format({'num_format': '@','align':'center'})

# the border of gap 
    cgFormat=workbook.add_format({'bg_color': '#FFFFFF','left':1,'right':1})
#cgFormat.

# the border of the very end gap 
    cgLastFormat=workbook.add_format({'bg_color': '#FFFFFF','right':1})
#cgFormat.

#   #

    investFormat = workbook.add_format({'bold': 1,'bg_color': '#CCFFFF','align':'center','border':2})
    invNameFormat = workbook.add_format({'bold': 2,'bg_color': '#FFFFFF','align':'center','border':2})


 # Add a number format for cells with money.
    money_format = workbook.add_format({'num_format': '$#,##0.00000','align':'center'})

# Add number format for open price, close price , numbers with no $ sign 
    number_format = workbook.add_format({'num_format': '0.00000','align':'center'})

# Add number format for share %, numbers with no $ sign 
    numPer_format = workbook.add_format({'num_format': '0.00000\%','align':'center'})


# numbers in P/L pips cells with green color 
    num_pips_format=workbook.add_format({'num_format': '0.00000','align':'center','bg_color':'#CCFFCC', 'border':1})





# Blanking the white empty cells 
    for i in range(hRow):
       for j in range(sCol+nInvestors+nInvestors*invHeads):
            worksheet.write(i,j,None,blankFormat)

# Blanking the gaps between investors #
    for i in range(hRow):
        for j in range(nInvestors+1):
            worksheet.write(i,sCol+j+j*invHeads,None,blankFormat)

    for i in range(hRow,maxBlank):
        for j in range(nInvestors+1):
            worksheet.write(i,sCol+j+j*invHeads,None,cgFormat)


# the right side of last gap #
    for i in range(hRow):
        j=nInvestors
        worksheet.write(i,sCol+j+j*invHeads,None,cgLastFormat)


#for i in range(maxBlank):
#    worksheet.write(i,9,None,blankFormat)


## Insert Images on the top left #

    worksheet.insert_image('A1', 'Prog.png',{'x_offset': 15, 'y_offset': 10,'x_scale':1,'y_scale':1})
    worksheet.insert_image('D1', 'FN4X.png',{'x_offset': 15, 'y_offset': 10,'x_scale':1,'y_scale':1})



# Start filling the header cells 

# Trade section #
    for i in range(tradeHeads):
        worksheet.write_string(hRow, i,Title_Dict[i],headerFormat)

# Investor sections 
    for j in range(nInvestors):
        for i in range(sCol+1, sCol+invHeads+1):
            worksheet.write_string(hRow, i+j+j*invHeads,Inv_Dict[i-sCol-1],headerFormat)
        strInv="Investor "+str(j+1)
        worksheet.write_string(hRow-2, sCol+1+j+j*invHeads,strInv,invNameFormat)        
        worksheet.write_string(hRow-1, sCol+1+j+j*invHeads,nameInvestors[j],investFormat)

#P/L pips Coloring for the columns 
    for i in range(maxBlank):
        worksheet.write(hRow+i+1, 6,None,Pl_format)



## Inserting data of trade information ## 

 # Start from the first cell below the headers.

    for i in range(nRows):
#    print trade[i][2]
#    print trade[i][7]    
#    date=datetime.strptime(trade[i][2], "mm/dd/yyyy") # 2 is the columns of date from cvs, format of date # ?! 
        worksheet.write_string(hRow+i+1, 0, trade[i][2],casFormat) # Date
#    worksheet.write_datetime(hRow+i, 0, date, date_format)
    

        worksheet.write_string(hRow+i+1, 1, trade[i][0], casFormat) # Currency
        worksheet.write_string(hRow+i+1, 2, trade[i][1], casFormat) # $ B/S
        if trade[i][3] != '':
            worksheet.write_number(hRow+i+1, 3, float(trade[i][3]), number_format)#Open Price 

        worksheet.write_string(hRow+i+1, 4, trade[i][4],casFormat) # Date close
        if trade[i][5] != '':
            worksheet.write_number(hRow+i+1, 5, float(trade[i][5]), number_format)#close Price 

        if trade[i][6] != '':
            worksheet.write_number(hRow+i+1, 6, float(trade[i][6]), num_pips_format)# P/L Pips  


        if trade[i][7] != '':
            worksheet.write_number(hRow+i+1, 7, float(trade[i][7]), money_format);## P/L $ Money





# Filling the inf of investors # 
        for j in range(nInvestors):
                for k in range(sCol+1, sCol+invHeads+1):
                    if trade[i][j+j*invHeads+k] != '':
                        if k-sCol==2 or k-sCol==7:  # No $ sign 
#                        print trade[i][j+j*invHeads+k]
                            worksheet.write_number(hRow+i+1, j+j*invHeads+k,float(trade[i][j+j*invHeads+k]), numPer_format)
                        else:
                            worksheet.write_number(hRow+i+1, j+j*invHeads+k,float(trade[i][j+j*invHeads+k]), money_format)                        




    workbook.close()
















#################################################################################################################
#This function converts the Trades' xlsx file into the intermediate csv file used by this program
#################################################################################################################


## The first input os the excel file the second is outcome  csv file and the third column is number of rows 
# would skipped from excel file to csv due to the presence of pictures and empty rows  

def xlsxTocsv(inputFile,outputFile,skipRows):
	data_xlsx = pd.read_excel(inputFile, 'Sheet1', index_col=None,skiprows=skipRows)
	#new_columns = [data_xlsx.columns[i] + "Count" if data_xlsx.columns[i].find("Unnamed") >= 0 
	#else data_xlsx.columns[i] for i in range(len(data_xlsx.columns))]

#	print data_xlsx.columns[3:]
	columnsTitles=[data_xlsx.columns[1]]+[data_xlsx.columns[2]]+[data_xlsx.columns[0]]
	#+data_xlsx.columns[3:]		
	for i in range(3,len(data_xlsx.columns)):
		columnsTitles=columnsTitles+[data_xlsx.columns[i]]
#	print columnsTitles
	data_xlsx=data_xlsx[columnsTitles]

	data_xlsx.as_matrix()

	data=data_xlsx[:]
	data_xlsx.to_csv(outputFile, encoding='utf-8',index=False)
#	return 

	new_columns=[]
	# Mannualy titling the cols #
	nCol=len(data_xlsx.columns)
	Titles=["Currency", 'B/S', "Date Open", 'Open Price', 'Date Close', 'Close Price', 'P/L pips', 'P/L ($)', 'Total Pips']
	nTitles=len(Titles)

# Headers for columns of investors information #
	Invs=["Open Trade Investment ($)","Open Trade Share (%)", 'Trade P/L ($)', 'Commission ($)', 'Amount Changed ($)', 'Closed Trade Investment ($)', 'Closed Trade Share (%)']
	nInvs=len(Invs)
	NumInvs=(nCol-(nTitles))/nInvs

	new_columns=Titles
	for n in range(NumInvs):
		new_columns=new_columns+[None]+Invs
	data_xlsx.columns=new_columns

#	columnsTitles=[new_columns[1]]+[new_columns[2]]+[new_columns[0]]+new_columns[3:]
#	print len(columnsTitles)
#	data_xlsx=data_xlsx[columnsTitles]
#	data_xlsx=data_xlsx.reindex(columns=columnsTitles)
	
#

	data_xlsx.as_matrix()

	data=data_xlsx[:]
	data_xlsx.to_csv(outputFile, encoding='utf-8',index=False)


## The first input os the excel file the second is outcome  csv file and the third column is number of rows 
# would skipped from excel file to csv due to the presence of pictures and empty rows  

## Testing the function ## 
# xlsxTocsv('Trades.xlsx','Trades.csv',10)







#################################################################################################################
#This function checks if the latest month in the trades' spreadsheet has been lossy for investors
#################################################################################################################



#We first compute a month and year whose end is the closest to when flush commissions is called
#The way this is done is, to go back in time in terms of close dates and check the latest month
#Another safeguard is that this month should have a close date larger than 10th of that month
def compute_flush_commission_flags(file):
	
	N1=9
	N2=8

	months={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}


	latest_month = -1
	latest_year = -1
	iteration = 0
	iteration_latest_month = -1
	iteration_latest_year = -1

	with open(file,'rU') as csvfile:
		F = csv.reader(csvfile)

		for row in reversed(list(F)):
			#We've crossed trades lines if length of row is too small
			if(len(row)<6):
				break

			if( (row[0]!="Currency") and (row[2]!="") and (row[4]!="") ):
				#We are in a close trade (traversing in reverse order of open time)
				#Check close month and date -> if > 10 this must be the latest month for flushing commissions
				dmy=row[4].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100

				iteration+=1
				if(iteration==1):
					iteration_latest_month = month
					iteration_latest_year = year

				if(day>10):
					latest_month=month
					latest_year=year				
					break
				
	#if latest month hasn't been found, then maybe we need to set latest month to the closing month of the most recent opened trade
	if(latest_month<0):
		# we already know something went wrong -> no trade in this month after 10th of this month???
		latest_month = iteration_latest_month
		latest_year = iteration_latest_year

	print "For flush commissions Latest month = "+str(latest_month)+" Latest year = "+str(latest_year)

	#Now we read this file knowing what month and year to compute opening and closing balances for
	with open(file,'rU') as csvfile:
		F = csv.reader(csvfile)
		j=1

		found_entry_trade = -1

		max_day=-1
		max_hour=-1
		max_minute=-1

		for row in F:			
			if (j==1):
				num_investors=(len(row)-N1)/(N2)
				entry_balance=[0]*num_investors
				final_balance=[0]*num_investors								

				flush_commission_flags = [1]*num_investors

			
			# We check if open date matches the latest month
			if ( (j>1) and (row[2]!="") ):
				#print str(row[2])
				dmy=row[2].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100


				if( (month==latest_month) and (year==latest_year) and (found_entry_trade<0) ):
					#print "Found entry trade at row = "+str(j)
					for i in range(num_investors):
						entry_balance[i]=float(row[N1+i*N2+1])
						found_entry_trade = 1


			#Now we check each closed trade and find the latest one in the latest month
			if( (j>1) and (row[6]!="") ):

				dmy=row[4].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100
			
				if( (month==latest_month) and (year==latest_year) ):
					if( (day>max_day) or ( (day==max_day) and (hour>max_hour) ) or ( (day==max_day) and (hour==max_hour) and (minute>max_minute) ) ):
						for i in range(num_investors):
							final_balance[i]=float(row[N1+i*N2+6])
			j+=1
						
	#Now that we have both the entry balance and the final balance for each investor for the latest month
	#We can compute the flush_commission flags

	for i in range(num_investors):
		print "Initial balance for investor "+str(i)+" = "+str(entry_balance[i])+" final balance = "+str(final_balance[i])
		if(final_balance[i]<entry_balance[i]):
			flush_commission_flags[i]=-1

	return flush_commission_flags
















#################################################################################################################
#This function spreads the financing amount across investors, across open trades as per open share percentages
#################################################################################################################



def compute_financing_upon_flush(file, amount,n_investors):
	
	N1=9
	N2=8

	months={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}




	#We first want to figure out the month and year for which financing is happening
	latest_month = -1
	latest_year = -1
	iteration = 0
	iteration_latest_month = -1
	iteration_latest_year = -1

	with open(file,'rU') as csvfile:
		F = csv.reader(csvfile)

		for row in reversed(list(F)):
			#We've crossed trades lines if length of row is too small
			if(len(row)<6):
				break

			if( (row[0]!="Currency") and (row[2]!="") and (row[4]!="") ):
				#We are in a close trade (traversing in reverse order of open time)
				#Check close month and date -> if > 10 this must be the latest month for flushing commissions
				dmy=row[4].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100

				iteration+=1
				if(iteration==1):
					iteration_latest_month = month
					iteration_latest_year = year

				if(day>10):
					latest_month=month
					latest_year=year				
					break
				
	#if latest month hasn't been found, then maybe we need to set latest month to the closing month of the most recent opened trade
	if(latest_month<0):
		# we already know something went wrong -> no trade in this month after 10th of this month???
		latest_month = iteration_latest_month
		latest_year = iteration_latest_year

	print "For financing Latest month = "+str(latest_month)+" Latest year = "+str(latest_year)






	#Once we know that, given the financing amount, we partition this month-end financing amount
	#For this, we obtain the set of open shares as a dict of the open dates (iterating over this month backwards)

	dates = [0]
	open_shares = {}
	seen_correct_month_flag = 0
	with open(file,'rU') as csvfile:
		F = csv.reader(csvfile)

		for row in reversed(list(F)):
			#We've crossed trades lines if length of row is too small
			if(len(row)<6):
				break

			if( (row[0]!="Currency") and (row[2]!="") ):
				dmy=row[2].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100

				if( (month==latest_month) and (year==latest_year) ):
					seen_correct_month_flag = 1
					if(day not in dates):
						dates.append(day)
						open_shares[day] = []
						for i in range(n_investors):
							if(row[N1+i*N2+2]==""):
								open_shares[day].append(0)
							else:
								open_shares[day].append(row[N1+i*N2+2])
                                '''
				elif( (seen_correct_month_flag==1) and ((month!=latest_month) or (year!=latest_year)) ):
					#This is done to record the open trade shares from the 1st of the month until the first trade of the latest month has been opened
					open_shares[0] = []
					for i in range(n_investors):
						if(row[N1+i*N2+2]==""):
							open_shares[day].append(0)
						else:
							open_shares[day].append(row[N1+i*N2+2])
					seen_correct_month_flag+=1
                    '''

	#If open_shares[0] does not exist by now, then default it to: open_shares of first element
	min_open_shares_list = []
	for key in open_shares.keys():
  		min_open_shares_list.append(key)

  	min_open_shares = min(min_open_shares_list)
	#print "min_open_shares = "+str(min_open_shares)

	if(min_open_shares!=0):
		open_shares[0]=open_shares[min_open_shares]



	#Now that we have the list of open dates and the corresponding shares, we divide the financing amount into pieces
	financing_values = [0.0]*n_investors

	if(len(dates)<2):
		#No trades opened this month? Shouldn't happen
		print "No trades opened this month! Something wrong!"

	if(31 not in dates):
		dates.append(31)

	dates.sort()








	#print "The open dates for this month are: "+str(dates)

	for d_index in range(1,len(dates)):
		prev = dates[d_index-1]
		d = dates[d_index]
		financing_portion = float(d-prev)*float(amount)/31.000
		#print "The portion from date = "+str(prev)+" to date = "+str(d)+" is "+str(financing_portion)
		for i in range(n_investors):
			#print "Share of investor "+str(i)+" is = "+str(open_shares[prev][i])
			financing_values[i]+=float(float(float(open_shares[prev][i])*financing_portion)/100.0)

	#Our output is:
	#We perform the sanity check:
	print "For financing to have worked, this number should be close to 0: "+str(float(sum(financing_values))-float(amount))

	return financing_values		









	#corresponding to intervals of open trades (where we break ties for open trades on the same date, by choosing the later date)
	#then for each such interval, the financing amount of an investor comes from the open share of the investor 
	#multiplied by the fraction of the financing amount partitioned for this open trade














#################################################################################################################
#This segment of code produces intermediate files
#################################################################################################################


#We process the files here:

#At each step we create an updated version of the "File" which is the spreadsheet containing the trades.

#For safety, at each step, we create a backup file which is just a copy of the input trades file
File_backup=File_xlsx.rstrip('.xlsx')+"s_backup.xlsx"
copyfile(File_xlsx,File_backup)

Investors_backup=Investors.rstrip('.csv')+"_backup.csv"
copyfile(Investors, Investors_backup)


#First we convert the excel file to csv, for the code to be able to process it
File = File_xlsx.rstrip('.xlsx')+".csv"
xlsxTocsv(File_xlsx,File,10)
File_new=File.rstrip('.csv')+".new.csv"




#################################################################################################################
#This function performs all the tasks involved in squaring off a trade
#These tasks are:
# 1) Check if there is incompatibility between the number of investors in the investors file and trades' file
# 2) Check how many trades are ripe to be closed; user needs to make sure to close one trade at a time
# 3) 
#################################################################################################################


#This function checks if a trade is open and closes it by computing necessary quantities for each investor.
#PLEASE MAKE SURE TO CLOSE EACH TRADE SEPARATELY (in serial order). This is required for the values to be computed correctly.
#To close a trade, the user puts in the Close Date, the Close Price, and the Pips columns herself. Rest of the columns are computed by the program.
def close_trades(): 

	N1=9
	N2=8

	months={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}

	names = []
	commissions = []
	current_commission = []
	added_funds = []
	number_of_investors = 0

	#We first read the file containing the list of investors to store the commission percentage to be subtracted for each investor
	with open(Investors,'rU') as investors_file:
		F_investors = csv.reader(investors_file)
		j=1
		for row in F_investors:
			if(j==1):
				j=1
				header_row = row
			else:
				number_of_investors += 1
				names.append(row[0])
				commissions.append(row[1])
				current_commission.append(row[2])
				added_funds.append(row[3])
			j+=1	

	#We then read the spreadsheet containing the trades and compute the most recently closed trade for each investor. This is required to compute the balance of the investor after the trade will be closed.
	with open(File,'rU') as csvfile:
		F1 = csv.reader(csvfile)
		j=1
		for row in F1:
			if (j==1):
				num_investors=(len(row)-N1)/(N2)

				#We check that the number of investors in the file containing the names and commission percentage of each investor, and the spreadsheet containing the trades, are compatible
				if(num_investors!=number_of_investors):
					print "Number of investors in investors file = "+str(number_of_investors)
					print "Number of investors in trades' spreadsheet = "+str(num_investors)
					print "investors file is incompatible with trades' spreadsheet. Exiting!"

					os.remove(File)
					sys.exit(-1)


			##### We don't need this!!!
			# 	#print "number of investors = "+str(num_investors)
			# 	latest_closing_balance=[0]*num_investors
			# 	max_day = 1
			# 	max_month = 1
			# 	max_year = 1
			# else:
			# 	if(row[6]!=""):
			# 		#print "Open"+str(j-1)

			# 		dmy=row[4].split("-")
			# 		day=int(dmy[0])
			# 		month=dmy[1]
			# 		year=int(dmy[2])

			# 		#print str(max_year)+" "+str(year)+" "+str(months[month])+" "+str(max_month)+" "+str(day)

			# 		if( (year>max_year) or ( (year==max_year) and (months[month]>max_month) ) or ( (year==max_year) and (months[month]==max_month) and (day>=max_day) ) ):
			# 			max_year=year
			# 			max_month=months[month]
			# 			max_day=day
		
			# 			#print str(max_year)+" "+str(year)+" "+str(months[month])+" "+str(max_month)+" "+str(max_day)+" "+str(day)

			# 			for i in range(num_investors):
			# 				latest_closing_balance[i]=float(row[N1+i*N2+6])
							
			# j+=1	

	num_trades_to_close=0

	with open(File_new, "wb") as csv_out:
		F_out = csv.writer(csv_out, delimiter=',')
		with open(File,'rU') as csvfile_2:
			F = csv.reader(csvfile_2)
			j=1
			for row in F:
				if (j==1):
					num_investors=(len(row)-N1)/(N2)
					F_out.writerow(row)
				else:
					if( (row[4]!="") and (row[6]=="") ):
						num_trades_to_close+=1

						#If the user accidentally tries to close multiple trades at once, i.e. she has entered close date/close price of more than one trade before running the close trades function, then we should exit and request them to close trades separately, in serial order.
						if(num_trades_to_close>1):
							print "More than one trade to close! Re-run and close each trade separately! Exiting!"
							os.remove(File_new)
							os.remove(File)
							sys.exit(-1)

						op=float(row[3])
						cp=float(row[5])
						
						if(row[1]=="B"):						
							buy_sell_sign=+1
						else:
							buy_sell_sign=-1	

						r_temp=str((cp-op)*buy_sell_sign)												

						if('JPY' in row[0]):
							row[6]=str(float(r_temp)*100)
						else:
							row[6]=str(float(r_temp)*10000)
						
						if( row[7]=="" ): #or (row[8]=="") ):
							print "The user has not filled out the Profit/Loss dollar amount or the Total Pips. Please enter them and relaunch!"
							os.remove(File_new)
							os.remove(File)
							sys.exit(-1)

						total_open = 0
						total = 0
						for i in range(num_investors):
							#If investor was added after trade was opened, they don't participate in open trade share but do in close trade share
							if(row[N1+i*N2+6]==""):	
								total_open+=float(row[N1+i*N2+1])

						for i in range(num_investors):
							
							#Commission for first investor (Primary investor) should be 0 in the investors file
							X=float(commissions[i])

							if(row[N1+i*N2+5]==""):
								amount_added = 0
							else:
								amount_added = float(row[N1+i*N2+5])

							#If this investor was added after this trade was opened, then they will have a non-empty close value
							#Their commission should be 0 so should be there P/L 
							if(row[N1+i*N2+6]!=""):
								row[N1+i*N2+2]=0
								row[N1+i*N2+3]=0
								row[N1+i*N2+4]=0
								row[N1+i*N2+6]=str(float(row[N1+i*N2+6])+amount_added)
							else:
								row[N1+i*N2+2]=str(100*(float(row[N1+i*N2+1])/total_open))
								row[N1+i*N2+3]=str((float(row[N1+i*N2+2])*float(row[7]))/100)
							

								#Commission
								#if commission is negative then make it 0
								if(float(row[N1+i*N2+3])<0):
									row[N1+i*N2+4] = 0
								else:
									row[N1+i*N2+4]=str(X*float(row[N1+i*N2+3]))

								#Update close value
								row[N1+i*N2+6]=str(float(row[N1+i*N2+1])+float(row[N1+i*N2+3])+amount_added) # We no longer subtract the commission at each trade -float(row[N1+i*N2+4])) #-row[N1+i*N2+4] because commision is taken away	

							current_commission[i] = str(float(current_commission[i])+float(row[N1+i*N2+4]))

						total_after=0
						for i in range(num_investors):
							total_after+=float(row[N1+i*N2+6])
						for i in range(num_investors):
							row[N1+i*N2+7]=str(100*(float(row[N1+i*N2+6])/total_after))

							

					F_out.writerow(row)
							
				j+=1
	os.remove(File)
	os.rename(File_new,File)
	
	out_row = ['','','','']
	#Now we rewrite the investors file with the updated commissions in the third column
	with open(Investors,'wb') as investors_file_out:
		F_investors_out = csv.writer(investors_file_out)
		for j in range(num_investors+1):
			if(j==0):
				F_investors_out.writerow(header_row)
			else:
				out_row[0] = str(names[j-1])
				out_row[1] = str(commissions[j-1])
				out_row[2] = str(current_commission[j-1])
				out_row[3] = str(0)
				F_investors_out.writerow(out_row)


	if(num_trades_to_close==0):
		print "No trades to close!\nUser might have entered the (green) pips column! That is computed by the program, so please leave it empty!"		
	else:
		print str(num_trades_to_close)+" trades have been closed!"




	csvToxlsx(File,Investors,File_xlsx)
	os.remove(File)
	sys.exit(1)





#This function simply finds the most recently closed trade and computes the current balance of each investor and writes that to a new row, so the user knows what will be the amount invested by each investor upon opening a new trade.
def open_trade(financing_text):

	if(financing_text==""):
		flush_commissions = False
	else:
		flush_commissions = True
		amount = float(financing_text)	

	#print "Flush commissions flag = "+str(flush_commissions)

	N1=9
	N2=8

	months={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}

	#First we read the spreadsheet containing the trades to find the row corresponding to the most recently closed trade
	with open(File,'rU') as csvfile:
		F1 = csv.reader(csvfile)
		j=1

		max_j = -1
		max_year=-1
		max_month=-1
		max_day=-1
		max_hour=-1
		max_minute=-1

		for row in F1:
			if (j==1):
				num_investors=(len(row)-N1)/(N2)
				value_1=[0]*num_investors
				value_2=[0]*num_investors
			elif (row[0]==""):
				print "There is a trade waiting to be opened before you can Open a new trade!"
				os.remove(File)
				sys.exit(1)
			else:
				if(row[6]!=""):
					#print "Open"+str(j-1)

					dmy=row[4].split("-")
					day=int(dmy[2])
					month=int(dmy[1])
					year=int(dmy[0])%100

					hour=int(dmy[3])
					minute=int(dmy[4])

					#print str(max_year)+" "+str(year)+" "+str(months[month])+" "+str(max_month)+" "+str(day)

					if( (year>max_year) or ( (year==max_year) and (month>max_month) ) or ( (year==max_year) and (month==max_month) and (day>=max_day) ) or ( (year==max_year) and (month==max_month) and (day==max_day) and (hour>=max_hour) ) or ( (year==max_year) and (month==max_month) and (day>=max_day) and (hour==max_hour) and (minute>max_minute) ) ):
						max_j = j
						max_year=year
						max_month=month
						max_day=day
						max_hour=hour
						max_minute=minute


			j+=1	

	#print "max_j was determined to be = "+str(max_j)





	#if the flush_commissions flag is set, then we move all the commissions to the primary investor's account and zero out the commission column in the investors file
	commission = []
	name = []
	af = []
	cf = []
	counter=0
	header_row=[]

	with open(Investors,'rU') as investors_file:
		F_investors = csv.reader(investors_file)
		j=1
		for row in F_investors:
			if(j==1):
				header_row=row
			else:
				commission.append(row[2])
				name.append(row[0])
				cf.append(row[1])
				af.append(row[3])
				counter+=1
			j+=1	

	# print "num_investors = "+str(counter)+" "+str(len(commission))+" "+str(len(name))+" "+str(len(cf))+" "+str(len(af))


	out_row = ['','','','']

	with open(Investors,'wb') as investors_file_out:
		F_investors_out = csv.writer(investors_file_out)
		for j in range(counter+1):
			if(j==0):
				F_investors_out.writerow(header_row)
			else:
				out_row[0] = name[j-1]
				out_row[1] = cf[j-1]
				if(flush_commissions==True):
					out_row[2] = 0
				else:
					out_row[2] = commission[j-1]
				out_row[3] = af[j-1]
				F_investors_out.writerow(out_row)


	with open(File_new, "wb") as csv_out:
		F_out = csv.writer(csv_out, delimiter=',')
		with open(File,'rU') as csvfile_2:
			F = csv.reader(csvfile_2)
			j=1
			
			for row in F:
				if (j==2):
					for i in range(num_investors):
						value_1[i]=float(row[N1+i*N2+1])
						#value_2[i]=float(row[N1+i*N2+6])
						if(row[N1+i*N2+6]==""):
							#print "random"
							value_2[i]=0
						else:	
							value_2[i]=float(row[N1+i*N2+6])								

				elif (j==max_j):
					for i in range(num_investors):
						if(row[N1+i*N2+6]==""):
							#print "random"
							value_2[i]=0
						else:	
							value_2[i]=float(row[N1+i*N2+6])								

				F_out.writerow(row)
							
				j+=1

		new_row=['']*(N1+num_investors*N2)

		#This is handling the special case when the trades spreadsheet has <3 trades
		if(max_j<0):
			for i in range(num_investors):
				new_row[N1+i*N2+1]=str(value_1[i]+float(af[i])) # We add the added funds to the open trade amount
		else:
			for i in range(num_investors):
				new_row[N1+i*N2+1]=str(value_2[i]+float(af[i])) # We add the added funds to the open trade amount


		#if flush commission flag is set, move the commissions to the primary investors account
		if(flush_commissions==True):

			flush_commission_flags_list = compute_flush_commission_flags(File)

			for i in range(num_investors):
				#If the month has been lossy for an investor then don't charge her any commission
				if(flush_commission_flags_list[i]==-1):
					print "Commission for investor "+str(name[i])+" will not be charged since it was a lossy month for him!"
					new_row[N1+i*N2+1]=str(float(new_row[N1+i*N2+1])+float(commission[i]))
				else:
					new_row[N1+0*N2+1]=str(float(new_row[N1+0*N2+1])+float(commission[i]))

			#We also reduce each investor's amount by distributing the financing across the investors
			financing = compute_financing_upon_flush(File, amount,num_investors)
			for i in range(num_investors):				
				new_row[N1+i*N2+1]=str(float(new_row[N1+i*N2+1])-float(financing[i]))

		F_out.writerow(new_row)			

	print "1 Trade has been opened!"

	os.remove(File)
	os.rename(File_new,File)
	csvToxlsx(File,Investors,File_xlsx)
	os.remove(File)
	sys.exit(1)







#This function adds a line for a new investor in the file containing investors' information. It also adds additional columns for this new investor in the spreadsheet containing the trades.
def add_investor(name_investor,commision_investor,initial_investment):
	
	# print "Name of investor to be added = "+str(name_investor)+" Commision fraction of new investor = "+str(commision_investor) 

	#Make sure the user entered non-null values for the arguments to this function in the corresponding text boxes
	if( (name_investor=="") or (commision_investor=="") or (initial_investment=="") ):
		print "Please enter values in Add Investor text boxes and then click the button!"
		#os.remove(File_new)
		os.remove(File)
		sys.exit(-1)

	with open(Investors,'ab') as investors_file:
		print "Opening investor file!"

		F_investors = csv.writer(investors_file,delimiter=',')
		row = [str(name_investor),str(commision_investor),str(0),str(0)] ###When adding a column in the investors file
		#F_investors.writerow([])    ###### Retain for Unix, Comment for windows - Parth 27/10/2017
		F_investors.writerow(row)

	with open(File_new, "wb") as csv_out:
		copyfile(File, File_backup)

		N1=9
		N2=8

		F_out = csv.writer(csv_out, delimiter=',')
		with open(File,'rU') as csvfile_2:
			F = csv.reader(csvfile_2)
			j=1
			
			for row in F:
				if(j==1):
					out_row = row
					out_row.extend(row[N1:N1+N2])
				else:
					out_row = row
					out_row.extend(['',str(initial_investment)])
					out_row.extend(['']*(N2-4))
					out_row.extend([str(initial_investment),''])
				F_out.writerow(out_row)
				j+=1

	print "A new investor named "+str(name_investor)+" with commission = "+str(commision_investor)+" and initial investment = "+str(initial_investment)+" has been added!"
	os.remove(File)
	os.rename(File_new,File)
	csvToxlsx(File,Investors,File_xlsx)
	os.remove(File)
	sys.exit(1)
	
Report_Header = ["Currency","B/S","Date Open","Open Price","Date Close","Close Price","P/L Pips","P/L ($)","Total Pips","Value for opening trade","% Share at open","P/L ($)","Commission","Added Funds","Value after closing trade","% Share at close"]			
	
def generate_report(name_of_investor,month_report,year_report):

	months={'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}

	freport_name = str(name_of_investor+"_"+month_report+"_"+year_report+"_Report") 

	freport=open(str(freport_name+".txt"),'w')

	freport.write("Name \t "+str(name_of_investor)+"\nMonth \t "+str(month_report)+"\nYear \t "+str(year_report)+"\n")

	Investor_by_serial_number=[]
	serial_number_of_investor=-1	
	with open(Investors,'rU') as investors_file:
		F_investors = csv.reader(investors_file)
		j=1
		for row in F_investors:
			if(j>1):
				Investor_by_serial_number.append(row[0])
				#print "Input = "+str(name_of_investor)+" Current = "+str(row[0])
				if(name_of_investor==row[0]):
					serial_number_of_investor = j-2
			j+=1		
	
	if(serial_number_of_investor<0):
		print "Something wrong with the investor name! Exiting!"
		#os.remove(File_new)
		os.remove(File)
		sys.exit(-1)

	N1=9
	N2=8


	Open_Trades=[]
	Closed_Trades=[]
	Header=[]

	with open(File,'rU') as csvfile:
		F = csv.reader(csvfile)
		j=1

		correct_month_flag=0

		num_trades_in_this_month=0
		max_day = -1

		for row in F:			
			if (j==1):
				num_investors=(len(row)-N1)/(N2)
				entry_balance=[0]*num_investors
				initial_balance=[0]*num_investors
				initial_share=[0]*num_investors
				current_balance=[0]*num_investors
				current_share=[0]*num_investors
				final_balance=[0]*num_investors								
				final_share=[0]*num_investors
				net_gain=[0]*num_investors
				net_percentage_gain=[0]*num_investors								
				net_entry_gain=[0]*num_investors
				net_percentage_entry_gain=[0]*num_investors		

				Header=row[0:8]
				Header.extend(row[10:16])


			if (j==2):
				for i in range(num_investors):
					#print str(i)+" "+row[N1+i*N2+1]
					entry_balance[i]=float(row[N1+i*N2+1])

			#First consider trades closed this month which this investor was a part of
			if( (j>1) and (row[6]!="") and (row[N1+serial_number_of_investor*N2+2]!="") ):

				dmy=row[4].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100
			
				if( (months[month_report]==month)) and (year_report==str(year)):

					close_trade_info = row[0:8]
					close_trade_info.extend(row[N1+serial_number_of_investor*N2+1:N1+serial_number_of_investor*N2+7])
					Closed_Trades.append(close_trade_info)

					
					#print str(month)+" "+sys.argv[2]+" "+str(year)+" "+sys.argv[3]

					num_trades_in_this_month+=1

					if(correct_month_flag==0):
						for i in range(num_investors):
							initial_balance[i]=float(row[N1+i*N2+1])
							if(row[N1+i*N2+2]==""):
								initial_share[i]=0
							else:
								initial_share[i]=float(row[N1+i*N2+2])
							
							correct_month_flag=1

					#The spreadsheet is not sorted in close dates; so we have to find the last closed trade for this month
					if(day>max_day):
						day = max_day

						for i in range(num_investors):
							if(row[N1+i*N2+2]==""):
								current_balance[i]=0
								current_share[i]=0
							else:			
								current_balance[i]=float(row[N1+i*N2+6])
								current_share[i]=float(row[N1+i*N2+7])	

			#Now consider all open trades this investor was a part of that were opened before this month:
			if( (j>1) and (row[0]!="") and (row[6]=="") and (row[N1+serial_number_of_investor*N2+6]=="") ):
				dmy=row[2].split("-")
				day=int(dmy[2])
				month=int(dmy[1])
				year=int(dmy[0])%100

				if( (year<year_report) or ( (year==year_report) and (month<=months[month_report]) ) ):
					open_trade_info = row[0:8]
					open_trade_info.extend(row[N1+serial_number_of_investor*N2+1:N1+serial_number_of_investor*N2+7])
					Open_Trades.append(open_trade_info)
																	
			j+=1		



		# if(num_trades_in_this_month==0):
		# 	print "No trades closed in this month! Reporting all trades open as of the end of this month, which this investor was a part of"


		freport.write("\n")
		s=""
		for hi in Header:
			s=s+hi+"\t"
		freport.write(s.rstrip("\t")+"\n")
		for i in Open_Trades:
			s=""
			for ii in i:
				s=s+ii+"\t"
			freport.write(s.rstrip("\t")+"\n")
		freport.write("\n")	


		s=""
		for hc in Header:
			s=s+hc+"\t"
		freport.write(s.rstrip("\t")+"\n")	
		for j in Closed_Trades:
			s=""
			for jj in j:
				s=s+jj+"\t"
			freport.write(s.rstrip("\t")+"\n")
		freport.write("\n")	

		for i in range(num_investors):
			if(i==serial_number_of_investor):

				final_balance[i]=current_balance[i]
				final_share[i]=current_share[i]	
				net_gain[i]=final_balance[i]-initial_balance[i]

				if(initial_balance[i]!=0):
					net_percentage_gain[i]=100*(net_gain[i]/initial_balance[i])
				else:
					net_percentage_gain[i]=0

				net_entry_gain[i]=final_balance[i]-entry_balance[i]
				net_percentage_entry_gain[i]=100*(net_entry_gain[i]/entry_balance[i])

	#print "Serial number equals " + str(serial_number_of_investor) 		


	# freport.write("\n\nTrades open as of end of this month, that "+name_of_investor+" is part of:\n")
	# freport.write(tabulate(Open_Trades,headers=Report_Header)+"\n\n\n")

	# freport.write("\n\nTrades closed during this month, that "+name_of_investor+" was part of:"+"\n")
	# freport.write(tabulate(Closed_Trades,headers=Report_Header)+"\n\n\n")

	if(correct_month_flag==0):
		print "No trades closed in this month for "+name_of_investor+". Hence his balance hasn't changed and we don't have any summary stats to report."
		os.remove(str(freport_name+".txt"))
		os.remove(File)
		sys.exit(-1)
	else:
		freport.write("\n\nSummary of Investor "+Investor_by_serial_number[serial_number_of_investor]+" for month "+month_report+" of year 20"+year_report+"\n")
		freport.write("")
		for i in range(num_investors):
			if(i==serial_number_of_investor):
				freport.write("Initial balance at the beginning of the month \t"+str(initial_balance[i])+"\n")
				freport.write("Percentage share in pool of money at the beginning of the month \t"+str(initial_share[i])+"\n")
				freport.write("Final balance at the end of the month \t"+str(final_balance[i])+"\n")
				freport.write("Percentage share at the end of the month \t"+str(final_share[i])+"\n")
				freport.write("Net gain since beginning of month \t"+str(net_gain[i])+"\n")
				freport.write("Percentage gain in this month w.r.t balance at the beginning of the month \t"+str(net_percentage_gain[i])+"\n")
				freport.write("Net gain since entry in pool \t"+str(net_entry_gain[i])+"\n")
				freport.write("Percentage gain since entry in pool \t"+str(net_percentage_entry_gain[i])+"\n")
				freport.write("\n\n")
	print "Generated report for Investor "+Investor_by_serial_number[serial_number_of_investor]+" for month "+month_report+" of year "+year_report



	freport.close()
	txtToxlsx(str(freport_name+".txt"),str(freport_name+".xlsx"))
	os.remove(str(freport_name+".txt"))

	#os.rename(File_new,File)
	csvToxlsx(File,Investors,File_xlsx)
	os.remove(File)
	sys.exit(1)



def add_funds(name_investor,amount_added):
	N1=9
	N2=8

	if(amount_added==""):
		print "Please enter amount!"
		#os.remove(File_new)
		os.remove(File)
		sys.exit(-1)


	found_flag=0
	out_row=[]
	output=[]
	serial_num = -1
	with open(Investors,'rU') as investors_file:
		F_investors = csv.reader(investors_file)
		j=1
		for row in F_investors:
			if(j==1):
				j=1
				header_row = row
				output.append(header_row)
			elif(name_investor==str(row[0])):
				serial_num = j-2
				found_flag+=1
				out_row.append(row[0])
				out_row.append(row[1])
				out_row.append(row[2])
				out_row.append(int(row[3])+int(amount_added)) ###When adding a column in the investors file
				output.append(out_row)
			else:
				output.append(row)
			j+=1	

	if(found_flag==0):
		print("Investor not found in Investor file. Unable to add funds. Exiting!")


	with open(Investors,'wb') as investors_file_out:
		F_investors_out = csv.writer(investors_file_out)
		for j in range(len(output)):
			F_investors_out.writerow(output[j])
			j+=1

	#Now we have to go to all open trades and add funds for this investor 
	with open(File_new, "wb") as csv_out:
		F_out = csv.writer(csv_out, delimiter=',')
		with open(File,'rU') as csvfile_2:
			F = csv.reader(csvfile_2)
			j=1
			
			for row in F:
				if( (j>1) and (row[0]!="") and (row[6]=="") ):
					if(row[N1+serial_num*N2+5]==""):
						row[N1+serial_num*N2+5] = amount_added
					else:
						row[N1+serial_num*N2+5] = str(float(row[N1+serial_num*N2+5]) + float(amount_added))							

				F_out.writerow(row)
							
				j+=1



	print "Succesfully added $"+str(amount_added)+" to "+str(name_investor)+"'s account."
        os.remove(File)
	os.rename(File_new,File)
	csvToxlsx(File,Investors,File_xlsx)
	os.remove(File)
	sys.exit(1)



















pos = [2,14,26,38,50]


win = Tk()
win.configure(background='white')


label_header_1 = Label(text='Swing Capital Trading Manager v1.0').grid(row=1,column=3)


#Creat three dropdowns next to Generate Report for Name, Month and Year

investor_name_for_report = ""
month_for_report = ""
year_for_report = ""

investor_name_for_report_choices_list=['']
month_for_report_choices_list=['']
year_for_report_choices_list=['']

def nameClicked(btn):
	investor_name_for_report=btn
	investor_name_for_report_choices_list.append(investor_name_for_report)
	btnMenu.config(text=btn)

def monthClicked(btn):
	month_for_report=btn
	month_for_report_choices_list.append(month_for_report)
	btnMenu_2.config(text=btn)

def yearClicked(btn):
	year_for_report=btn		
	year_for_report_choices_list.append(year_for_report)
	btnMenu_3.config(text=btn)

btnMenu = Menubutton(win, text='Select Investor')
contentMenu = Menu(btnMenu)
btnMenu.config(menu=contentMenu)
btnMenu.grid(row=pos[2],column=3,padx=40)

Investors_Names_List=[]
with open(Investors,'rU') as investors_file_read:
	F_investors_File = csv.reader(investors_file_read)
	jj=1
	for row in F_investors_File:
		if(jj>1):
			Investors_Names_List.append(row[0])
		jj+=1	

#print str(Investors_Names_List)

for btn in Investors_Names_List:
	contentMenu.add_command(label=btn, command = lambda btn=btn: nameClicked(btn))



btnMenu_2 = Menubutton(win, text='Select Month')
contentMenu_2 = Menu(btnMenu_2)
btnMenu_2.config(menu=contentMenu_2)
btnMenu_2.grid(row=pos[2],column=4,padx=40)

Month_List=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

for btn in Month_List:
	contentMenu_2.add_command(label=btn, command = lambda btn=btn: monthClicked(btn))

btnMenu_3 = Menubutton(win, text='Select Year')
contentMenu_3 = Menu(btnMenu_3)
btnMenu_3.config(menu=contentMenu_3)
btnMenu_3.grid(row=pos[2],column=5,padx=60)

Year_List=['15','16','17','18','19','20','21','22','23','24','25']

for btn in Year_List:
	contentMenu_3.add_command(label=btn, command = lambda btn=btn: yearClicked(btn))

def generate_report_call():
	if( (len(investor_name_for_report_choices_list)<2) or (len(month_for_report_choices_list)<2) or (len(year_for_report_choices_list)<2) ):
			print "Please select Investor, Month and Year of Report!"
	else:		
		generate_report(investor_name_for_report_choices_list[-1],month_for_report_choices_list[-1],year_for_report_choices_list[-1])


#Create two textboxes for Add Investor

label1 = Label(text='Name:')
label1.grid(row=pos[4],column=3,sticky="e")
entry1 = Entry(win, width=10)
entry1.grid(row=pos[4],column=4,sticky="w")

label2 = Label(text='Commission:').grid(row=pos[4],column=4,sticky="e")
entry2 = Entry(win, width=10)
entry2.grid(row=pos[4],column=5,sticky="w")

label3 = Label(text='Initial investment:').grid(row=pos[4],column=5,sticky="e")
entry3 = Entry(win, width=10)
entry3.grid(row=pos[4],column=6,sticky="w")


def add_investor_call():
	add_investor(entry1.get(),entry2.get(),entry3.get())

#Add textbox in front of Open Trade 

label101 = Label(text='Flush Commissions and Enter Financing Fee: ').grid(row=pos[1],column=3)
entry101 = Entry(win, width=10)
entry101.grid(row=pos[1],column=4,sticky="w")


def open_trade_call():
	open_trade(entry101.get())


#Add dropdown for investor names and a checkbox for amount added (subtracted)
investor_name_for_AF = ""
investor_name_for_AF_list=['']

def nameClicked_AF(btn_AF):
	investor_name_for_AF=btn_AF
	investor_name_for_AF_list.append(investor_name_for_AF)
	btnMenu_AF.config(text=btn_AF)

btnMenu_AF = Menubutton(win, text='Select Investor for Changing Funds')
contentMenu_AF = Menu(btnMenu_AF)
btnMenu_AF.config(menu=contentMenu_AF)
btnMenu_AF.grid(row=pos[3],column=3,padx=40)

Investors_Names_List_AF=[]
with open(Investors,'rU') as investors_file_for_AF:
	F_investors_File_for_AF = csv.reader(investors_file_for_AF)
	jj=1
	for row in F_investors_File_for_AF:
		if(jj>1):
			Investors_Names_List_AF.append(row[0])
		jj+=1	

# print "AF list = "+str(Investors_Names_List_AF)

for btn_AF in Investors_Names_List_AF:
	contentMenu_AF.add_command(label=btn_AF, command = lambda btn_AF=btn_AF: nameClicked_AF(btn_AF))


label_AF = Label(text='Amount to change: ').grid(row=pos[3],column=4,sticky="e")
entry_AF = Entry(win, width=10)
entry_AF.grid(row=pos[3],column=5,sticky="w")


def add_funds_call():
	# print str(investor_name_for_AF_list)
	add_funds(investor_name_for_AF_list[-1],entry_AF.get())




#All open trades are closed. No arguments to this.
b1 = Button(win,text="Close Trades",background="red",height=1,width=30,command=close_trades)
b1.grid(row=pos[0],column=2,pady=50,sticky="w")
b1.config(highlightbackground="red")

#A new row is created in the file. No arguments to this.
b2 = Button(win,text="Open a new Trade",height=1,width=30,command=open_trade_call)
b2.grid(row=pos[1],column=2,pady=50,sticky="w")
b2.config(highlightbackground="red")


#First argument is investor name. Also, two more arguments: month and year. Generates a pdf report for this investor.
b3 = Button(win,text="Generate Report",height=1,width=30,command=generate_report_call)
b3.grid(row=pos[2],column=2,pady=50,sticky="w")
b3.config(highlightbackground="red")

#No arguments to this. Assume that user has added the added value in the last open trade. Run only once for a trade.
b4 = Button(win,text="Add Funds",height=1,width=30,command=add_funds_call)
b4.grid(row=pos[3],column=2,pady=50,sticky="w")
b4.config(highlightbackground="red")


#Two arguments to this. Investor name and amount to begin with.
b5 = Button(win,text="Add Investor",height=1,width=30,command=add_investor_call)
b5.grid(row=pos[4],column=2,pady=50,sticky="w")
b5.config(highlightbackground="red")



##############################
#Add buttons to open the two files

def open_trades_file():
    print "open trades"
    os.system('open "Trades.xlsx"')

def open_investors_file():
    print "open investors"  
    os.system('open "Investors.csv"')

b6 = Button(win,text="Open Trades' File",height=1,width=30,command=open_trades_file)
b6.grid(row=pos[0],column=5,pady=50,sticky="w")
b6.config(highlightbackground="blue")

b7 = Button(win,text="Open Investors' File",height=1,width=30,command=open_investors_file)
b7.grid(row=pos[0],column=6,pady=50,sticky="w")
b7.config(highlightbackground="blue")


mainloop()


