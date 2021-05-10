import openpyxl
import smptlib

pending_payment={}

def getDefaulters():
	
	global pending_payment
	excelObj=openpyxl.load_workbook('Vendor_payment.xlsx')
 
	sheet=excelObj.active.title
	
	for row in range(2,excelObj[sheet].max_row+1):
		for column in range(2,excelObj[sheet].max_column+1):
		     
			if excelObj[sheet].cell(row=row,column=column).value==None:
				pending_payment.update({excelObj[sheet].cell(row=row,column=2).value:[]})

				


	for key in pending_payment.keys():
	     
		pending_payment[key]=[excelObj[sheet].cell(row=1,column=column).value for row in range(2,excelObj[sheet].max_row+1) for column in range(3,excelObj[sheet].max_column+1) if excelObj[sheet].cell(row=row,column=2).value == key and excelObj[sheet].cell(row=row,column=column).value==None]
		
	
	
	return(pending_payment)


def sendMail(defaulter):
	
	


	

getDefaulters()



sendMail(getDefaulters())        			
