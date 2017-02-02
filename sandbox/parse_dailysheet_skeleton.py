"""
### Kouvaya ###
The purpose of this python document is to parse through the daily sheet uploads
and format the data into the necessary format to be uploaded to a google sheet where
ecery row represent 1 day's worth of activity.
This architecture will allow for much easier viewing and editing on the part of admins
of the PAAD.tech@gmail.com account


Sheet 1:
Sheet 2: Provider Summary
Sheet 3: Deposit Slip
Sheet 4: Credit Card Slips to deposit
"""


#import necessary read excel docs
import xlrd
import csv
from datetime import datetime
import os # used to access files
import glob # used to grab items


# create CSV
skeleton = open("skeleton_master_file.csv", "wb")
writer = csv.writer(skeleton, delimiter=",")

#create headline
### note: on the cloud, this skeleton will have already been created
writer.writerow(["Date", "Patients", "New Patients", "Production", "Production per Provider", "Cash Collections", "Check Collections", "Credit Card Collections", "Collections Per Provider", "Ortho Cases", "Hygiene Cases" ])

## get a list of files and loop through them!
os.chdir("/Users/paulstavropoulos/Google Drive/kouvaya/reporting-stack/jan_daily_sheets")
daysheets = glob.glob("*.xls*")

for item in daysheets:
    # open a dailysheet workbook to read from
    #this here is a test document -->
    # on the VM server, we will use the Google API to grab the last dailysheets that have not been uploaded
    # perform this code on each of them
    wkbk = xlrd.open_workbook(item)

    '''
    Work for sheet 1
    Format for sheet 1:
    rows 0-3 are headers, column 0 is trash
    '''
    sheet1 = wkbk.sheet_by_index(0)
    row = 6

    # define the dictionary to store all information to print
    # initialize items
    daily_info = {}
    daily_info["prod"] = daily_info["cash_collections"] = daily_info["check_collections"] =  daily_info["credit_collections"] = 0
    daily_info["ortho_cases"] = daily_info["hygiene_cases"] = 0
    daily_info["patients"] = daily_info["new_patients"] = 0
    daily_info["prod_per_provider"] = {}
    daily_info["collections_per_provider"] = {}

    # TODO: find out how to record billable insurance
    # TODO: find out how to record referrals


    # while loop --> end when column 6 is empty (Description)
    # do i care that all rows are not unique to single patient? I don't think so
    # in SQL we can do a search for a user and be fine (no duplicates either)
    # row 6 is never empty when there is data --> keep looping until it is empty
    while(sheet1.cell(row,6).value != ''):

        # production created
        if(sheet1.cell(row,7).value != ''):
            daily_info["prod"] += sheet1.cell(row,7).value

        #production created per provider/dentist
        # check if dr code in dictionary prod_per_provider
        # if it is, ad prod value. If it is not, initiate with prod value
            if sheet1.cell(row,3).value != '' and sheet1.cell(row,7).value != '':
                if sheet1.cell(row,3) in daily_info["prod_per_provider"].keys():
                    daily_info["prod_per_provider"][int(sheet1.cell(row,3).value)] += sheet1.cell(row,7).value
                else:
                    daily_info["prod_per_provider"][int(sheet1.cell(row,3).value)] = sheet1.cell(row,7).value

        # ortho cases and hygiene deposit
        # check description and code columns
        if 'invisalign' in sheet1.cell(row,6).value.lower():
            daily_info["ortho_cases"] += 1
        if sheet1.cell(row,5).value != '' and (int(sheet1.cell(row,5).value) == 1110 or int(sheet1.cell(row,5).value) == 4341 or int(sheet1.cell(row,5).value) == 1206):
            daily_info["hygiene_cases"] += 1

        # collections
        # cash, credit, check
        # also perform collections per provider work in this sections as well
        # use dr. code to find out
        if(sheet1.cell(row,10).value != ''):
            daily_info['cash_collections'] += int(sheet1.cell(row,10).value)

            # check to see if provider is in dictionary of collections_per_provider
            if sheet1.cell(row,3).value in daily_info["collections_per_provider"].keys():
                daily_info["collections_per_provider"][int(sheet1.cell(row,3).value)] += sheet1.cell(row,10).value
            else:
                daily_info["collections_per_provider"][int(sheet1.cell(row,3).value)] = sheet1.cell(row,10).value

        if(sheet1.cell(row,11).value != ''):
            daily_info['check_collections'] += int(sheet1.cell(row,11).value)

            # check to see if provider is in dictionary of collections_per_provider
            if sheet1.cell(row,3).value in daily_info["collections_per_provider"].keys():
                daily_info["collections_per_provider"][int(sheet1.cell(row,3).value)] += sheet1.cell(row,11).value
            else:
                daily_info["collections_per_provider"][int(sheet1.cell(row,3).value)] = sheet1.cell(row,11).value

        if(sheet1.cell(row,12).value != ''):
            daily_info['credit_collections'] += int(sheet1.cell(row,12).value)

            # check to see if provider is in dictionary of collections_per_provider
            if sheet1.cell(row,3).value in daily_info["collections_per_provider"].keys():
                daily_info["collections_per_provider"][int(sheet1.cell(row,3).value)] += sheet1.cell(row,12).value
            else:
                daily_info["collections_per_provider"][int(sheet1.cell(row,3).value)] = sheet1.cell(row,12).value

        # increment row!
        row += 1

    # once we've reached the end of the list, there are a static number of columns
    # between current row number and the information we want for summary metrics reported

    # record patients and new patients
    # look for cell that equals "patients" and "new patients"
    if sheet1.cell(row + 14, 6).value != '':
        daily_info["patients"] = int(sheet1.cell(row + 14, 6).value)
    if sheet1.cell(row + 15, 6).value != '':
        daily_info["new_patients"] = int(sheet1.cell(row + 15, 6).value)



    # write to row in skeletion
    #"Date", "Patients", "New Patients", "Production", "Production per Provider", "Cash Collections", "Check Collections",
    #"Credit Card Collections", "Collections Per Provider", "Ortho Cases", "Hygiene Cases"
    writer.writerow([datetime(*xlrd.xldate_as_tuple(sheet1.cell(2,0).value, wkbk.datemode)).date(), daily_info["patients"], daily_info["new_patients"], daily_info["prod"], daily_info["prod_per_provider"], daily_info["cash_collections"], daily_info["check_collections"], daily_info["credit_collections"], daily_info["collections_per_provider"], daily_info["ortho_cases"], daily_info["hygiene_cases"]
    ])
