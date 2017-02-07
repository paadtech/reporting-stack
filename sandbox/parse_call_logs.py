"""
### Kouvaya ###
The purpose of this python document is to parse through the call logs created by the PAAD office
and format the data into so as to enter row in the master file where 1 row will represent 1 day's worth of activity.
This architecture will allow for much easier viewing and editing on the part of admins
of the PAAD.tech@gmail.com account

"""

import xlrd
import csv
from datetime import datetime
import os # used to access files
import glob # used to grab items


# include methods here to call later in scripts
def removeFiles():
    files = [f for f in os.listdir('.') if os.path.isfile(f)]
        # for all the files in the folder that are an excel sheet
        for file in files:
            # If the files have a .xls
            if ".csv" in file and 'master' not in file:
                os.remove(file)




#### define needed function to initlize dictionary for every unique date found
def initializeDict():
    temp = {}
    temp["new_patients"] =0
    temp["number_answered_phone"] = 0
    temp["number_left_messages"] = 0
    temp["number_recall_emails"] = 0
    temp["appts_scheduled"] = 0
    temp["number_phones_disconnected"] = 0
    temp["number_people_moved"] = 0
    temp["number_local_office_changes"] = 0
    temp["reports_by"] = {}

    return temp


# open csv in append mode
skeleton = open("master_call_log_file.csv", "a")
writer = csv.writer(skeleton, delimiter=",")

#create headline
### note: on the cloud, this skeleton will have already been created
#writer.writerow(["date", "New Patients", "Number Answered Phone", "Number of Left Messages", "Number of Recall Emails Sent", "Appts Scheduled", "Number Phones Disconnected", "Number of People Moved", "Number of Local Office Changes", "Reports By"])


# Create arrays of Yes and no answers
yes_checks = ["y", "yes"]
#no_checks = ["n", "no", "na"]


# create path to file name
# os.chdir("/Users/paulstavropoulos/Google Drive/kouvaya/reporting-stack/example_docs")
#grab all .csvs, except if they have the string "master"
call_logs = []
allfiles = [f for f in os.listdir('.') if os.path.isfile(f)]
    # for all the files in the folder that are an excel sheet
    for file in allfiles:
        # If the files have a .xls
        if ".csv" in file and 'master' not in file:
            call_logs.push(file)




#loop through all items
for item in call_logs:
    try:
        # open item
        open_item = open(item, "rb")
        log = csv.reader(open_item)


        # create a dictionary that will store dictionaries for every unique date found
        dates = {}

        passed_header = False
        for row in log:
            #create boolean to check if we passed row where first item is "date"
            # we want to start parsing rows only when we've passed the header row
            # which we can't bet consistently occurs in the same spot
            try:
                if 'date' in row[0].lower():
                    passed_header = True
                    continue
            except:
                pass

            # once we've passed header, then start
            if passed_header == True:

                # check to see if first item has a date that is dictionary.
                # if not, initialize it and add
                try:
                    if row[0] not in dates.keys():
                        dictionary = initializeDict()
                        dates[row[0]] = dictionary
                except:
                    pass

                # for every column, perform a series of yes checks and appts_scheduled
                #new patient
                try:
                    if row[2].lower() in yes_checks:
                        dates[row[0]]["new_patients"] += 1
                except:
                    pass

                try:
                    #did answer
                    if row[3].lower() in yes_checks:
                        dates[row[0]]["number_answered_phone"] += 1
                except:
                    pass

                try:
                    #did leave Messages
                    if row[4].lower() in yes_checks:
                        dates[row[0]]["number_left_messages"] += 1
                except:
                    pass

                try:
                    #number recall emails
                    if row[5].lower() in yes_checks:
                        dates[row[0]]["number_recall_emails"] += 1
                except:
                    pass

                try:
                    # appts Scheduled
                    if row[6].lower() in yes_checks:
                        dates[row[0]]["appts_scheduled"] += 1
                except:
                    pass

                try:
                    #phone Disconnected
                    if row[7].lower() in yes_checks:
                        dates[row[0]]["number_phones_disconnected"] += 1
                except:
                    pass

                try:
                    # person moved
                    if row[8].lower() in yes_checks:
                        dates[row[0]]["number_people_moved"] += 1
                except:
                    pass

                try:
                    # changed local office
                    if row[9].lower() in yes_checks:
                        dates[row[0]]["number_local_office_changes"] += 1
                except:
                    pass

                try:
                    # add rows for every reporter
                    if row[10].lower().strip() in dates[row[0]]["reports_by"].keys():
                        dates[row[0]]["reports_by"][row[10].lower().strip()] += 1
                    if row[10].lower().strip() not in dates[row[0]]["reports_by"].keys():
                        dates[row[0]]["reports_by"][row[10].lower().strip()] = 1
                except:
                    pass

        for key in dates.keys():
            try:
                writer.writerow([key, dates[key]["new_patients"], dates[key]["number_answered_phone"], dates[key]["number_left_messages"], dates[key]["number_recall_emails"], dates[key]["appts_scheduled"], dates[key]["number_phones_disconnected"], dates[key]["number_people_moved"], dates[key]["number_local_office_changes"], dates[key]["reports_by"]])
            except:
                pass

        # close log we opened
        open_item.close()

    except:
        pass

#close writer
skeleton.close()

# remove the same files that we have just parseDaysheet
removeFiles()
