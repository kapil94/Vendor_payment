# Vendor_payment
To get the defaulters in payment and send an automated email. 

The program uses openpyxl module and smtplib module.

Algorithm:

1. Open the Workbook object for Vendor-payment excel.
2. Traverse the workbook object and check if the cell of excel file doesn't contain any data i.e None.
3. Store the email_id of defaulters as a key in pending_payment dictionary.
4. Traverse through the keys of pending_payment dictionary and store the months from which payment has been pending.
5. Now traverse through the pending_payment dictionary and send an automated mail to all keys of dictionary by stating that they have been missing payment from pending_payment[key][0] month.
