"""
### Kouvaya ###
The purpose of this python document is to parse through the daily sheet uploads
and format the data into the necessary format to be uploaded to a sql server


Sheet 1:
Sheet 2: Provider Summary
Sheet 3: Deposit Slip
Sheet 4: Credit Card Slips to deposit
"""

#import necessary read excel docs
import xlrd
import csv # not sure if I'll need this

'''
### TODO ###
when brough on the server,
check for what files need to parse and uploads
will probably have to also DL locally
https://developers.google.com/drive/v3/web/manage-downloads
'''



# open the workbook to read from
### TODO alterpath for VM
wkbk = xlrd.open_workbook('/Users/paulstavropoulos/Google Drive/kouvaya/reporting-stack/example_docs/11162016Daysheet.XLS')


'''
Work for sheet 1
Format for sheet 1:
rows 0-3 are headers, column 0 is trash
TODO --
'''
# for each person, create a dictionary. Thus, we have an array of dicts
provier_info = []
sheet1 = wkbk.sheet_by_index(0)

row = 4

# while loop --> end when column 6 is empty (Description)
# do i care that all rows are not unique to single patient? I don't think so
# in SQL we can do a search for a user and be fine (no duplicates either)
while(sheet1.cell(row,6) != ''):
    # init variables
    patient = {}
    hygiene_dep = 0

    patient['name'] = sheet1.cell(row,2).value
    patient['ID'] = sheet1.cell(row,1).value
    patient['description'] = sheet1.cell(row,6).value
    patient['prod'] = sheet1.cell(row,7).value

    # code --> used to determine hygiene field
    # assumes all we care about is overall count
    # TODO
    patients['code'] = int(sheet1.cell(row,5).value)
    if patients['code'] == 1110 or patients['code'] == 4341 or patients['code'] == 1206:
        hygiene_dep += 1

    #collection - sum of cash, check, Credit
    patient['collection'] = 0
    if(sheet1.cell(row,10) != ''):
        patient['collection'] += sheet1.cell(row,10).value

    if(sheet1.cell(row,11) != ''):
        patient['collection'] += sheet1.cell(row,11).value

    if(sheet1.cell(row,12) != ''):
        patient['collection'] += sheet1.cell(row,12).value












'''
Work for sheet 2
Looking for collection and production/provider,
patients/day,


rows 0-4 are trash,
Whenever a number appears in column 1, grab provider name and ID nubmer. This serves as dict key
'''

sheet2 = wkbk.sheet_by_index(1)

# create an array of dicts -->
provider_info = []

# while loop for [row, 3] != total
# this goes through all necesseary rows
row = 5
while (sheet2.cell(row,1) != "TOTALS"):
    currProvider = None
    if(sheet2.cell(row,0).value != "") :
        # create dictionary for provider
        provider_dict = {}
        provider_dict['id'] = str(int(sheet2.cell(row,0).value))
        provider_dict['name'] = sheet2.cell(row,0).value








######Work for sheet 3




######Work for sheet 4
