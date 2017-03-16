from __future__ import print_function, division
import csv
import sys
from xlrd import open_workbook
import os, shutil
from __builtin__ import file
from shutil import copy
from vobject.icalendar import PRODID

"""
Script to extract named sheets from a workbook into csv

*** salesfile is passed from command line passed as -w salesfile  
*** Sheets are converted to csv in data/ folder
"""
# Command line Argument Handling
try:
    import argparse
    parser = argparse.ArgumentParser(description='Script for creating csv files from xls file')
    parser.add_argument('-m','--dictfile', help='e.g -m fido_dict.csv', required=True)
    parser.add_argument('-w','--salesfile', help='e.g -w salesfile.csv', required=True)
    parser.add_argument('-d','--date', help='e.g -d 2017-03-09', required=True)
    args = vars(parser.parse_args())
except ImportError:
    parser = None

if not os.path.exists('./data'):
    os.makedirs('./data')

if not os.path.exists('./OUT'):
    os.makedirs('./OUT')

    
DICTFILE = open(args['dictfile'], 'rt')
WBFILE = args['salesfile']
DATEINVOICE = args['date']
ERRORFILE = 'OUT/errfile' + DATEINVOICE + '.csv'


errfile = open(ERRORFILE, 'w')
DATAFOLDER = 'data'
salesid = {}
custid = {}

# delete all files in folder
def delfiles(folder):
    for the_file in os.listdir(folder):
        file_path = os.path.join(folder, the_file)
        try:
            if os.path.isfile(file_path):
                os.unlink(file_path)
        except Exception as e:
            print(e)

def customerqc(name,type):
    """
     DO quality control on customer name to reduce rejections
    """
    custname = name.strip()
    if type == 'CUSTOMER':
        
        
        custname = custname.replace('Jesus Love','Jesus-Love')
        custname = custname.replace('Roland','Rowland')
        custname = custname.replace('Christain','Christian')
        custname = custname.replace('Pishoh Gole','Pishon Gole')
        custname = custname.replace('Stella Amaran','Stella Amara')
        custname = custname.replace('Omorome','Omoreme')    
        custname = custname.replace('Egerekumo','Ederekumo')
        custname = custname.replace('Mathew','Matthew')
        custname = custname.replace('Mercy Ndubuisi','Mercy Ndubusi')        
        custname = custname.replace('Daniel Egiri','Erigi')
        
      
        custname = custname.replace('Okafor Priscilla','Priscilla Okafor')
        custname = custname.replace('Nigerian Neavy','Nig Navy')
        custname = custname.replace('Nigerian Navy','Nig Navy')
        custname = custname.replace('Doris Ogede','Ogede Doris')
        custname = custname.replace('Emeka Okolo','Okolo Emeka')
        custname = custname.replace('Sunday David','Sunny David')
        custname = custname.replace('Olayode Ujro','Olayode Ujiro')
        custname = custname.replace('Ganiyu Ayo','Ganiyu Motunrayo')
        custname = custname.replace('Chima Customer','Chima-Customer')
        custname = custname.replace('Ayodele Franca','Franca Ayodele')
        custname = custname.replace('Godspower Customer','Godspower-Customer')
        custname = custname.replace('New Integrated Service','New Integrated Services')
        custname = custname.replace('Florence','Flourence')
        custname = custname.replace('Seriyai Sam','Serieya Sam')
    if type == 'SALES':
       
        custname = custname.replace('Kingley','Kingsley')
        custname = custname.replace('Kroboh Oweibiagha','Kroboh Oweibigha')
        custname = custname.replace('Kroboh Owiebigha','Kroboh Oweibigha')

    return custname                
    
def reformat (file,prodtype):
    
    if 'OBUN' in file:
        PRODLOC = 'OBUNNA'
    if 'KPANSIA' in file:
        PRODLOC = 'KPANSIA'
    if 'DISP' in file:
        PRODLOC = 'DISPENSER'
    if 'CRATE' in file:
        PRODLOC = "CRATES"
    
    outfile = open("OUT/out_%s" %(file), 'w')
    reader1 = csv.reader(DICTFILE)
    
    file1 = 'data/'+file
    for row in reader1:
        sperson = row[2].upper()
        salesid[sperson] = row[3]
        cperson = row[0].upper()
        custid[cperson] = row[1]
        
    reader2 = csv.reader(open(file1, 'rt'))
    CSVHEADER = 'id,payment_term_id/id,account_id/id,user_id/id,user_id/name,partner_id/id,partner_id/display_name,date_invoice,invoice_line_ids/product_id/id,invoice_line_ids/name,invoice_line_ids/account_id/id,invoice_line_ids/quantity,invoice_line_ids/price_unit'
    print (CSVHEADER,file=outfile)
    rcount = 0
    ercount = 0

    for row in reader2:
        try:
            rcount = rcount + 1
            custname = customerqc(row[3],'CUSTOMER')
            custname = custname.upper()
            salesperson = customerqc(row[1],'SALES')
            salesperson = salesperson.upper()
            # print(custname,' master ',salesperson)
            sid = ('__export__.res_users_' + salesid[salesperson]).rstrip()
            cid = ('__export__.res_partner_' + custid[custname]).rstrip()
            if 'OBUN' in file:
                printstr = ',__export__.account_payment_term_7,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+'__export__.product_product_421'+','+prodtype+','+'__export__.account_account_204'+','+row[7]+','+row[8]
                
            if 'KPANSIA' in file:                
                printstr = ',__export__.account_payment_term_8,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+'__export__.product_product_421'+','+prodtype+','+'__export__.account_account_204'+','+row[7]+','+row[8]
            
            if 'DISP' in file:                
                printstr = ',__export__.account_payment_term_9,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+'__export__.product_product_433'+','+prodtype+','+'__export__.account_account_98'+','+row[7]+','+row[8]
            
            if 'CRATE' in file:
                prodtyp = row[5]
                if '50CL' in prodtyp:
                    prdid = '__export__.product_product_883'
                    acctid = '__export__.account_account_205'
                    
                if '60CL' in prodtyp:
                    prdid = '__export__.product_product_572' 
                    acctid = '__export__.account_account_206'
                
                printstr = ',__export__.account_payment_term_10,__export__.account_account_7,'+sid +','+salesperson+','+cid+','+custname+','+DATEINVOICE +','+prdid+','+prodtyp+','+acctid+','+row[7]+','+row[8]
            if (int(row[7]) != 0) and (int(row[8]) != 0):
                print (printstr,file=outfile)
        except KeyError as e:
            ercount = ercount + 1
            print('Customer,',row[3],',Salesperson,',row[1],',',e.args[0],',**** ',file,file=errfile)
            continue
    
    print ('PRODLOC,LINES,Errors,Errpct\n%s,%s,%s,%.2f%%\n' %(PRODLOC, rcount,ercount,((ercount-1)/(rcount-1))*100))
    outfile.close()
    
# Make import-ready files
def convfiles(folder):
    print ('ANALYSIS\n--------')
    for file in os.listdir(folder):
        if ('KPANSIA' in file) or ('OBUN' in file.upper()):
            reformat(file,'Purewater')
        elif ('DISPENSER' in file.upper()):
            reformat(file,'Dispenser')
        elif ('CRATE' in file.upper()):
            reformat(file,'Crate')
      
# Create csv files from sheets in Sales Workbook
def csvextract():
    wb = open_workbook(WBFILE)
    delfiles(DATAFOLDER)
    print ('SHEETS IN SALES FILE')
    
    for i in range(0, wb.nsheets-1):
        sheet = wb.sheet_by_index(i)
        print (sheet.name)
 
        path =  DATAFOLDER + '/%s.csv'
        with open( path %(sheet.name.replace(" ","")+DATEINVOICE), "w") as file:
            writer = csv.writer(file, delimiter = ",")
           
            header = [cell.value for cell in sheet.row(0)]
            writer.writerow(header)
 
            for row_idx in range(1, sheet.nrows):
                row = [int(cell.value) if isinstance(cell.value, float) else cell.value
                   for cell in sheet.row(row_idx)]
                writer.writerow(row)

def main():
    # extract csv from sheets in workbook
    csvextract()
    print ("\n")
    # Actual reformating
    convfiles(DATAFOLDER)
            
    print('See %s for error and data/ for source csv files\n and OUT/ directory for import-ready files' % ERRORFILE)        
            
if __name__ == '__main__':
    main()
