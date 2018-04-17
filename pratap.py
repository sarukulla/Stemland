# -*- coding: utf-8 -*-
"""
Created on Sun Mar 11 14:24:26 2018

@author: poovizhi
"""

 
 
More 
1 of 235
 
Formatting excel in python
Inbox
	x
prathap ganasan
	
Attachments8:42 PM (17 hours ago)
	
to Ranjith, Bala, me, S.logeshwari, Arun, naveen, Sundranandhan, Poovizhi
I have attached a python in which formatting of excel can be done in python.
Please go through it.
Attachments area
	
Click here to Reply, Reply to all, or Forward
0.82 GB (5%) of 15 GB used
Manage
Terms - Privacy
Last account activity: 7 hours ago
Details
	
	

## set formatting for xlxswriter
# The following line makes a cell Bold and  
fmt_bold = workbook.add_format({'bold': True})
# Example it makes first,first column Bold 
worksheet.write(0,0,'S.No',fmt_bold)

#The following line makes a cell Bold and Wrap the thext in a cell.
fmt_test_name = workbook.add_format({'bold': True,'text_wrap' : True})
# Its colours fills the cells with some colour.
worksheet.conditional_format(1,5,4*i+j+1,5, {'type': 'cell','criteria': '<>','value': 0,'format': format_red})
format_red = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

# Set titles
worksheet.write(0,0,'S.No',fmt_bold)
worksheet.write(0,1,'PROFILE',fmt_bold)
worksheet.write(0,2,'VARIABLE',fmt_bold)
worksheet.write(0,3,'LEGACY',fmt_bold)
worksheet.write(0,4,'PRO',fmt_bold)
worksheet.write(0,5,'DIFF',fmt_bold)
#Example for formatting the cell size
# Set formatting
inch_to_pxl = 12.5
worksheet.set_column(0,0,0.5*inch_to_pxl)
worksheet.set_column(1,1,2.5*inch_to_pxl)
worksheet.set_column(2,2,1.5*inch_to_pxl)
worksheet.set_column(3,5,2*inch_to_pxl)