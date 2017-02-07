"""
## Kouvaya ##
## Feb 2017 ##

This script sends an email to necessary stakeholders of the performance of the PAAD office
It will send an email every Sudnay from aggregating data from call logs and dailysheet data

TODO: modularize this -- this should be 2-3 different files
1 file should just have email methods in it
1 file should have master file parsing methods in it
1 file should be the main one that invokes those. 
"""

# imports
import os
import csv
import datetime
import time
import dateutil.relativedelta
import sendgrid
from sendgrid.helpers.mail import *

# global vars
# send_list, today, and trailing thirty day for monthly roll ups
today = datetime.datetime.today()
ttd = today - dateutil.relativedelta.relativedelta(months=1)
t7d = today - dateutil.relativedelta.relativedelta(days=7)
# grab call log data for the last month
# returns a dictionary of data for every day after the TTD
def month_call_log_data():
    row_data = {}

    try:
        with open ("master_call_log_file.csv", "rb") as cl:
            # open as csv
            reader = csv.reader(cl)

            #skip header row and grab all necessary rows larger than TTD
            next(cl)
            for row in reader:
                # convert row 0 to a date time to compare. We want to grab all rows
                # that are after the trailing thirty date number
                try:
                    date = time.strptime(row[0], "%m/%d/%Y")
                except:
                    date = today

                if date > ttd.timetuple():
                    # put data into the dictionary

                    # this can also serve as a function to correct mistakes!!
                    try:
                        # data to grab from call logs:
                        row_data[row[0]] = {"new_patients_logs":row[1], "number_answered_phone": row[2], "number_left_messages":row[3], "number_recall_emails":row[4],
                                        "appts_scheduled":row[5], "number_phones_disconnected":row[6], "number_people_moved":row[7], "number_local_office_changes":row[8],
                                        "reports_by":row[9]}
                    except:
                        pass
    except:
        pass
    return row_data


# grab dailysheet data for the last month
# returns a dictionary of data for every day after the TTD
def month_dailysheet_data():
    row_data = {}

    try:
        with open ("master_file.csv", "rb") as cl:
            # open as csv
            reader = csv.reader(cl)

            #skip header row and grab all necessary rows larger than TTD
            next(cl)
            for row in reader:
                # convert row 0 to a date time to compare. We want to grab all rows
                # that are after the trailing thirty date number
                try:
                    date = time.strptime(row[0], "%Y-%m-%d")
                except:
                    date = today

                if date > ttd.timetuple():
                    # put data into the dictionary
                    #TODO: Protect against over-writing information that might exist.
                    # this can also serve as a function to correct mistakes!!
                    try:
                    # data to grab from call logs:
                        row_data[row[0]] = {"patients":row[1], "new_patients": row[2], "prod":row[3], "prod_per_provider":row[4],
                                        "cash_collections":row[5], "check_collections":row[6], "credit_collections":row[7], "collections_per_provider":row[8],
                                        "ortho_cases":row[9], "hygiene_cases":row[9]}
                    except:
                        pass
    except:
        pass

    return row_data


def compute_total_metrics(log_data, dailysheet_data):
    # create rolled up dictionary for weekly and monthly metrics
    # and intialize

    monthly_metrics = {"ds_patients_from_dailysheet":0, "ds_new_patients_from_dailysheet":0, "ds_prod":0, "ds_prod_per_provider":{},
    "ds_cash_collections":0, "ds_check_collections":0, "ds_credit_collections":0, "ds_collections_per_provider":{},
    "ds_ortho_cases":0, "ds_hygiene_cases":0, "lg_new_patients_from_logs":0, "lg_number_answered_phone":0, "lg_number_left_messages":0, "lg_number_recall_emails":0,
    "lg_appts_scheduled":0, "lg_number_phones_disconnected":0, "lg_number_people_moved":0, "lg_number_local_office_changes":0,
    "lg_reports_by":{}}

    weekly_metrics = {"ds_patients_from_dailysheet":0, "ds_new_patients_from_dailysheet":0, "ds_prod":0, "ds_prod_per_provider":{},
    "ds_cash_collections":0, "ds_check_collections":0, "ds_credit_collections":0, "ds_collections_per_provider":{},
    "ds_ortho_cases":0, "ds_hygiene_cases":0, "lg_new_patients_from_logs":0, "lg_number_answered_phone":0, "lg_number_left_messages":0, "lg_number_recall_emails":0,
    "lg_appts_scheduled":0, "lg_number_phones_disconnected":0, "lg_number_people_moved":0, "lg_number_local_office_changes":0,
    "lg_reports_by":{}}


    # loop through log data and dailysheet data and do comparison!
    #if within the right interval, then add accordingly
    for date in log_data.keys():

        #change stringified dicts to dictionaries
        # this will make it possible to parse dictionaries --> pulling from CSV makes everything a string
        log_data_reports_by = eval(log_data[date]["reports_by"])


        # if the current day is larger than trailing thirty day
        try:
            if time.strptime(date, "%m/%d/%Y") > ttd.timetuple():
                # add data to monthly roll up
                # put each of these in a try-except

                try:
                    monthly_metrics["lg_new_patients_from_logs"] += int(log_data[date]["new_patients_logs"].strip())
                except:
                    pass
                try:
                    monthly_metrics["lg_number_answered_phone"] += int(log_data[date]["number_answered_phone"].strip())
                except:
                    pass
                try:
                    monthly_metrics["lg_number_left_messages"] += int(log_data[date]["number_left_messages"].strip())
                except:
                    pass
                try:
                    monthly_metrics["lg_number_recall_emails"] += int(log_data[date]["number_recall_emails"].strip())
                except:
                    pass
                try:
                    monthly_metrics["lg_number_phones_disconnected"] += int(log_data[date]["number_phones_disconnected"].strip())
                except:
                    pass
                try:
                    monthly_metrics["lg_number_people_moved"] += int(log_data[date]["number_people_moved"].strip())
                except:
                    pass
                try:
                    monthly_metrics["lg_number_local_office_changes"] += int(log_data[date]["number_local_office_changes"].strip())
                except:
                    pass

                for name in log_data_reports_by.keys():
                    # checkifitem is in the dictionary. If not, initiate
                    try:
                        if name in monthly_metrics["lg_reports_by"].keys():
                            monthly_metrics["lg_reports_by"][name] += log_data_reports_by[name]
                        else:
                            monthly_metrics["lg_reports_by"][name] = log_data_reports_by[name]
                    except:
                        pass


            # repeat this workflow for the weekly cut of data
            # this could be better --> in the future create a method to do this because it is essentially copy-pasted code from above
            if time.strptime(date, "%m/%d/%Y") > t7d.timetuple():
                #wrap these in try-excepts to prevent failure
                try:
                    weekly_metrics["lg_new_patients_from_logs"] += int(log_data[date]["new_patients_logs"].strip())
                except:
                    pass
                try:
                    weekly_metrics["lg_number_answered_phone"] += int(log_data[date]["number_answered_phone"].strip())
                except:
                    pass
                try:
                    weekly_metrics["lg_number_left_messages"] += int(log_data[date]["number_left_messages"].strip())
                except:
                    pass
                try:
                    weekly_metrics["lg_number_recall_emails"] += int(log_data[date]["number_recall_emails"].strip())
                except:
                    pass
                try:
                    weekly_metrics["lg_number_phones_disconnected"] += int(log_data[date]["number_phones_disconnected"].strip())
                except:
                    pass
                try:
                    weekly_metrics["lg_number_people_moved"] += int(log_data[date]["number_people_moved"].strip())
                except:
                    pass
                try:
                    weekly_metrics["lg_number_local_office_changes"] += int(log_data[date]["number_local_office_changes"].strip())
                except:
                    pass


                for name in log_data_reports_by.keys():
                    try:
                        # checkifitem is in the dictionary. If not, initiate
                        if name in weekly_metrics["lg_reports_by"].keys():
                            weekly_metrics["lg_reports_by"][name] += log_data_reports_by[name]
                        else:
                            weekly_metrics["lg_reports_by"][name] = log_data_reports_by[name]
                    except:
                        pass
        except:
            pass


    # for loop for dailysheet data
    # same workflow as above
    # For reference, this should be re-written to be more modular
    for date in dailysheet_data.keys():

        #change stringified dicts to dictionaries
        # this will make it easy for us to parse dictionaries
        try:
            dailysheet_data_prod_provider = eval(dailysheet_data[date]["prod_per_provider"])
            dailysheet_data_collections_provider = eval(dailysheet_data[date]["collections_per_provider"])
        except:
            dailysheet_data_prod_provider = dailysheet_data_collections_provider = {}

        try:
            # if the current day is larger than trailing thirty day
            #wrap all of these in try-excpet methods to prevent failure
            if time.strptime(date, "%Y-%m-%d") > ttd.timetuple():
                try:
                    monthly_metrics["ds_patients_from_dailysheet"] += int(dailysheet_data[date]["patients"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_new_patients_from_dailysheet"] += int(dailysheet_data[date]["new_patients"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_prod"] += float(dailysheet_data[date]["prod"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_cash_collections"] += int(dailysheet_data[date]["cash_collections"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_credit_collections"] += int(dailysheet_data[date]["credit_collections"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_check_collections"] += int(dailysheet_data[date]["check_collections"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_ortho_cases"] += int(dailysheet_data[date]["ortho_cases"].strip())
                except:
                    pass
                try:
                    monthly_metrics["ds_hygiene_cases"] += int(dailysheet_data[date]["hygiene_cases"].strip())
                except:
                    pass

                # populate prod per provider
                for item in dailysheet_data_prod_provider.keys():
                    try:
                        # check if item is in the dictionary. If not, initiate
                        if item in monthly_metrics["ds_prod_per_provider"].keys():
                            monthly_metrics["ds_prod_per_provider"][item] += dailysheet_data_prod_provider[item]
                        else:
                            monthly_metrics["ds_prod_per_provider"][item] = dailysheet_data_prod_provider[item]
                    except:
                        pass

                # populate collections per provide
                for item in dailysheet_data_collections_provider.keys():
                    try:
                        # check if item is in the dictionary. If not, initiate
                        if item in monthly_metrics["ds_prod_per_provider"].keys():
                            monthly_metrics["ds_collections_per_provider"][item] += dailysheet_data_collections_provider[item]
                        else:
                            monthly_metrics["ds_collections_per_provider"][item] = dailysheet_data_collections_provider[item]
                    except:
                        pass
        except:
            pass


        # repeat this workflow for the weekly cut of data
        # this needs improvement --> in the future create a method to do this
        if time.strptime(date, "%Y-%m-%d") > t7d.timetuple():
            try:
                weekly_metrics["ds_patients_from_dailysheet"] += int(dailysheet_data["patients"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_new_patients_from_dailysheet"] += int(dailysheet_data["new_patients"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_prod"] += int(dailysheet_data["prod"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_cash_collections"] += int(dailysheet_data["cash_collections"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_credit_collections"] += int(dailysheet_data["credit_collections"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_check_collections"] += int(dailysheet_data["check_collections"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_ortho_cases"] += int(dailysheet_data["ortho_cases"].strip())
            except:
                pass
            try:
                weekly_metrics["ds_hygiene_cases"] += int(dailysheet_data["hygiene_cases"].strip())
            except:
                pass

            # populate prod per provider
            for item in dailysheet_data_prod_provider.keys():
                try:
                    # check if item is in the dictionary. If not, initiate
                    if item in weekly_metrics["ds_prod_per_provider"].keys():
                        weekly_metrics["ds_prod_per_provider"][item] += dailysheet_data_prod_provider[item]
                    else:
                        weekly_metrics["ds_prod_per_provider"][item] = dailysheet_data_prod_provider[item]
                except:
                    pass

            # populate collections per provider
            for item in dailysheet_data_collections_provider.keys():
                try:
                    # check if item is in the dictionary. If not, initiate
                    if item in monthly_metrics["ds_prod_per_provider"].keys():
                        monthly_metrics["ds_collections_per_provider"][item] += dailysheet_data_collections_provider[item]
                    else:
                        monthly_metrics["ds_collections_per_provider"][item] = dailysheet_data_collections_provider[item]
                except:
                    pass



    # return two items
    return monthly_metrics, weekly_metrics


# write email
def writeEmail(monthly_data, weekly_data):
    # goal of the email is to clearly convey key metrics of on a weekly and monthly bases
    # will also convey date ranges considered to be in the last month and the last week.

    email_body = "<h2>PAAD office weekly and monthly email report</h2><br/><br/>This report shows weekly metrics from " + str(t7d.date()) + "<br/>and monthly metrics from " + str(ttd.date())
    email_body += "<p>These reports aggregate data from dailysheets and call logs and give a snapshot of the business</p><br/><br/>"
    email_body += "<h4>Weekly Metrics starting from " + str(t7d.date()) + ":</h4><br/>"

    # set up table
    email_body += "<table>"

    # loop through and print stat and output from week
    for item in weekly_data.keys():
        try:
            if item == "ds_prod" or item == "ds_check_collections" or item == "ds_cash_collections" or item == "ds_credit_collections":
                email_body += "<tr><td>" + item[3:] + "</td><td>$" + format(weekly_data[item], ",d") + "</td></tr>"
            else:
                email_body += "<tr><td>" + item[3:] + "</td><td>" + str(weekly_data[item]) + "</td></tr>"
        except:
            pass
    # close weekly table and open monthly table
    email_body+= "</table><br/><br/><h4>Monthly Metrics starting from " + str(ttd.date()) + ":</h4><br/>"
    email_body+= "<table>"

    for item in monthly_data.keys():
        try:
            if item == "ds_prod" or item == "ds_check_collections" or item == "ds_cash_collections" or item == "ds_credit_collections":
                email_body += "<tr><td>" + item[3:] + "</td><td>$" + format(monthly_data[item], ",d") + "</td></tr>"
            else:
                email_body += "<tr><td>" + item[3:] + "</td><td>" + str(monthly_data[item]) + "</td></tr>"
        except:
            pass

    #close email
    email_body += "<br/><br/>Cheers!<br/>The PAAD Team"
    return email_body


# send email
def sendEmail(email_text):
    # using sendgrid API, send an email to the people in the list defined above
    sg = sendgrid.SendGridAPIClient(apikey=os.environ.get('SENDGRID_PAAD_API_KEY'))

    # building email
    # this is taken from the sendgris API
    mail = Mail()
    mail.set_from(Email("PAAD.tech@gmail.com", "PAAD Office"))
    mail.set_subject("PAAD Weekly Summary Email")

    personalization = Personalization()
    personalization.add_to(Email("paul@kouvaya.com", "Paul Stav"))
    personalization.add_to(Email("sandrazbeautifulsmile@gmail.com", "Sandra Zeledon"))
    mail.add_personalization(personalization)
    mail.add_content(Content("text/html", email_text))
    response = sg.client.mail.send.post(request_body=mail.get())

    print(response.status_code)
    print(response.body)
    print(response.headers)


#### beginning of main ####
def main():

    # grab the right data from each sheet
    # grab data needed from the last thrity days
    call_log_data_month = month_call_log_data()
    dailysheet_data_month = month_dailysheet_data()

    # print call_log_data_month, dailysheet_data_month

    #compute weekly and monthly data metrics
    # this will return a list of 2 dictionaries
    monthly_data, weekly_data = compute_total_metrics(call_log_data_month, dailysheet_data_month)

    # for debugging purposes
    """
    print monthly_data
    print ' '
    print ' '
    print weekly_data
    """

    #write email
    email_text = writeEmail(monthly_data, weekly_data)
    # print email_text

    #sendEmail
    sendEmail(email_text)



if __name__ == "__main__":
    main()
