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
	
	smtpObj=smtplib.SMTP('smtp.gmail.com',587)
	smtpObj.starttls()
	
	user_email=input('Enter your email')
	
	
	for key in defaulter.keys():
		
		smtpObj.sendmail(user_email,key,'Subject:Urgent Reminder!!\n Dear vendor,It is to remind you that you not behind your payments from the month of '+defaulter[key][0]+'\n Kindly pay your dues in order to continue further services. Thanks and Regards,Kapil')
		

getDefaulters()
sendMail(getDefaulters())        			
