'''
Created on Feb 9, 2021
@Description: Combines the water usage for all users between 2018-2020
@author: Sean
'''
import openpyxl



#open workbook
wbk = openpyxl.load_workbook("C:\\Users\\Sean\\Documents\\AllAccountReadsCPY.xlsx", read_only = True)
#Create sheet
wbk2 = openpyxl.load_workbook("C:\\Users\\Sean\\Documents\\test2.xlsx")
        
print("Combining Sheets...")

"""
Combine the data for all three years so that goes as follows
User 1: 2018 data
User 1: 2019 data
User 1: 2020 data

User 2: 2018 data
User 2: 2019 data
User 2: 2020 data
....etc
"""


i = 0 #superfluous counter variable
for (row1, row2, row3) in zip(wbk['2018'].rows, wbk['2019'].rows, wbk['2020'].rows):
    print(i)
    wbk2['2018'].append([j.value for j in row1])
    wbk2['2018'].append([j.value for j in row2])
    wbk2['2018'].append([j.value for j in row3])
    i += 1
 
wbk2.save("C:\\Users\\Sean\\Documents\\test2.xlsx")
wbk2.close()
wbk.close()
print("FINISHED")




