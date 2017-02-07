"""
## Kouvaya ##
## Feb 2017 ##

This script sends an email to necessary stakeholders of the performance of the PAAD office
It will send an email every Sudnay from aggregating data from call logs and dailysheet data
"""

# imports
import os
import csv
import datetime
import time
import dateutil.relativedelta

# global vars
# send_list, today, and trailing thirty day for monthly roll ups
send_list = ["paul@kouvaya.com", "sandrazbeautifulsmile@gmail.com"]
today = datetime.datetime.today()
ttd = today - dateutil.relativedelta.relativedelta(months=1)
t7d = today - dateutil.relativedelta.relativedelta(days=7)
# grab call log data for the last month
# returns a dictionary of data for every day after the TTD
def month_call_log_data():
    row_data = {}


    with open ("master_call_log_file.csv", "rb") as cl:
        # open as csv
        reader = csv.reader(cl)

        #skip header row and grab all necessary rows larger than TTD
        next(cl)
        for row in reader:
            # convert row 0 to a date time to compare. We want to grab all rows
            # that are after the trailing thirty date number
            date = time.strptime(row[0], "%m/%d/%Y")

            if date > ttd.timetuple():
                # put data into the dictionary
                #TODO: Protect against over-writing information that might exist.
                # this can also serve as a function to correct mistakes!!

                # data to grab from call logs:
                # "date", "New Patients", "Number Answered Phone", "Number of Left Messages", "Number of Recall Emails Sent",
                # "Appts Scheduled", "Number Phones Disconnected", "Number of People Moved", "Number of Local Office Changes", "Reports By"
                row_data[row[0]] = {"new_patients_logs":row[1], "number_answered_phone": row[2], "number_left_messages":row[3], "number_recall_emails":row[4],
                                    "appts_scheduled":row[5], "number_phones_disconnected":row[6], "number_people_moved":row[7], "number_local_office_changes":row[8],
                                    "reports_by":row[9]}

    return row_data


# grab dailysheet data for the last month
# returns a dictionary of data for every day after the TTD
def month_dailysheet_data():
    row_data = {}


    with open ("master_file.csv", "rb") as cl:
        # open as csv
        reader = csv.reader(cl)

        #skip header row and grab all necessary rows larger than TTD
        next(cl)
        for row in reader:
            # convert row 0 to a date time to compare. We want to grab all rows
            # that are after the trailing thirty date number
            date = time.strptime(row[0], "%Y-%m-%d")

            if date > ttd.timetuple():
                # put data into the dictionary
                #TODO: Protect against over-writing information that might exist.
                # this can also serve as a function to correct mistakes!!

                # data to grab from call logs:
                #"Date", "Patients", "New Patients", "Production", "Production per Provider", "Cash Collections", "Check Collections",
                #"Credit Card Collections", "Collections Per Provider", "Ortho Cases", "Hygiene Cases"
                row_data[row[0]] = {"patients":row[1], "new_patients": row[2], "prod":row[3], "prod_per_provider":row[4],
                                    "cash_collections":row[5], "check_collections":row[6], "credit_collections":row[7], "collections_per_provider":row[8],
                                    "ortho_cases":row[9], "hygiene_cases":row[9]}

    return row_data



# write email
def writeEmail():
    print 'hi'

# send email
def sendEmail():
    print 'hi'


def compute_total_metrics(log_data, dailysheet_data):
    # create rolled up dictionary for weekly and monthly metrics
    monthly_metrics = {}
    weekly_metrics = {}

    # loop through log data and dailysheet data and do comparison!
    #if within the right interval, then add accordingly
    for date in log_data.keys():
        if date.time.strptime(row[0], "%m/%d/%Y")



#### beginning of main ####
def main():

    # grab the right data from each sheet
    # grab data needed from the last thrity days
    call_log_data_month = month_call_log_data()
    dailysheet_data_month = month_dailysheet_data()

    # print call_log_data_month, dailysheet_data_month

    #compute weekly and monthly data metrics
    # this will return a list of 2 dictionaries
    compute_total_metrics(call_log_data_month, dailysheet_data_month)

    #write email

    #sendEmail


if __name__ == "__main__":
    main()
