# -*- coding: utf-8 -*-
"""
@Author: Joel Feddes
@Date Modified: 12/2/2021
@Purpose:
    This program is designed to seek out and store monthly scorecards into assigned folders.
    This will greatly reduce the manual work on going into scorecards, checking where the distributor belongs,
    and manually moving the scorecards into correct locations. It also helps to eliminate human error.
"""

import win32com.client as client
import os
import pandas as pd
import re

'''
distyFileName strips off the part of the file name that corresponds to the distributors name
'''
def distyFileName(filename):
    filename = str(filename)
    disty_name = re.match("(.*?) (grow|incentive)",filename)
    return disty_name.groups()[0]

'''
fileTypeCheck checks to see if the attachment is a pdf file and returns True or False
'''
def fileTypeCheck(filename):
    filename = str(filename)
    if filename[-3:] == "pdf":
        return True
    else:
        return False
'''
dertermineRegion returns the region of the distributor OR returns the disty name if a region is not found.
'''
def determineRegion(disty_name,dictionary):
    if disty_region_dict.get(attachment_disty_name) != None:
        region = disty_region_dict.get(attachment_disty_name)
    else:
        region = disty_name
    return region

'''
fileAttachment takes in an attachment, a date, and the distributors region
then, the function files the pdf attachment into the appropriate folders
'''
def fileAttachment(attachment,date,region):
    date = date[0] + "-" + date[1]

    if region in "central southwest southeast northwest northeast west east south":            
        attachment.SaveASFile(os.path.join(usa_directory_dict.get(date), attachment.FileName))
        print(f"attachment {attachment.FileName} from {s} saved to {usa_directory_dict.get(date)}")
        
    elif region == "mexico":
        attachment.SaveASFile(os.path.join(mexico_directory_dict.get(date), attachment.FileName))
        print(f"attachment {attachment.FileName} from {s} saved to {mexico_directory_dict.get(date)}") 
        
    elif region == "canada":
        attachment.SaveASFile(os.path.join(can_directory_dict.get(date), attachment.FileName))
        print(f"attachment {attachment.FileName} from {s} saved to {can_directory_dict.get(date)}") 
    
    elif region == "rola":
        attachment.SaveASFile(os.path.join(rola_directory_dict.get(date), attachment.FileName))
        print(f"attachment {attachment.FileName} from {s} saved to {rola_directory_dict.get(date)}")

    else:
        print("This Disty's name isn't being picked up %s" % attachment_name)

'''
calcMonth takes in the date an email was recieved and returns the month it was recieved.
'''
def calcMonth(email_date):
    email_date_str = str(email_date)
    email_date_list = email_date_str.split("-")
    email_date_month = email_date_list[0:2]
    return email_date_month

        
disty_region_spreadsheet_path = r"C:\master_lists\2021\gwp-master-2021.csv"
df = pd.read_csv(disty_region_spreadsheet_path) #grab dataframe of master spreadsheet 2021
clean_df = df.dropna() #ditch rows with empty values. Useless

#clean up columns to remove any casing and to strip off any extra spaces at the end of names
clean_df["Regions"] = clean_df["Regions"].str.lower()
clean_df["Regions"] = clean_df["Regions"].str.strip()
clean_df["Distributors"] = clean_df["Distributors"].str.lower()
clean_df["Distributors"] = clean_df["Distributors"].str.strip()

disty_region_dict = dict(zip(clean_df["Distributors"], clean_df["Regions"])) #move items into a dictionary
outlook = client.Dispatch("Outlook.Application").GetNameSpace("MAPI")
account = outlook.Folders['joel.feddes@panduit.com'] #provide outlook credentials
gwp_folder = outlook.Folders.Item("Grow with Panduit")
score_cards_inbox = gwp_folder.Folders["Inbox"]

#The below dictionaries are file paths where the scorecards will be stored.
canada_directory_dict = {
                      "2021-02":r"C:\Users\JOEF\scorecards\can\2021\january",
                      "2021-03":r"C:\Users\JOEF\scorecards\can\2021\february",
                      "2021-04":r"C:\Users\JOEF\scorecards\can\2021\march",
                      "2021-05":r"C:\Users\JOEF\scorecards\can\2021\april",
                      "2021-06":r"C:\Users\JOEF\scorecards\can\2021\may",
                      "2021-07":r"C:\Users\JOEF\scorecards\can\2021\june",
                      "2021-08":r"C:\Users\JOEF\scorecards\can\2021\july",
                      "2021-09":r"C:\Users\JOEF\scorecards\can\2021\august",
                      "2021-10":r"C:\Users\JOEF\scorecards\can\2021\september",
                      "2021-11":r"C:\Users\JOEF\scorecards\can\2021\october",
                      "2021-12":r"C:\Users\JOEF\scorecards\can\2021\november",
                      "2022-01":r"C:\Users\JOEF\scorecards\can\2021\december",
    
                      "2022-02":r"C:\Users\JOEF\scorecards\can\2022\january",
                      "2022-03":r"C:\Users\JOEF\scorecards\can\2022\february",
                      "2022-04":r"C:\Users\JOEF\scorecards\can\2022\march",
                      "2022-05":r"C:\Users\JOEF\scorecards\can\2022\april",
                      "2022-06":r"C:\Users\JOEF\scorecards\can\2022\may",
                      "2022-07":r"C:\Users\JOEF\scorecards\can\2022\june",
                      "2022-08":r"C:\Users\JOEF\scorecards\can\2022\july",
                      "2022-09":r"C:\Users\JOEF\scorecards\can\2022\august",
                      "2022-10":r"C:\Users\JOEF\scorecards\can\2022\september",
                      "2022-11":r"C:\Users\JOEF\scorecards\can\2022\october",
                      "2022-12":r"C:\Users\JOEF\scorecards\can\2022\november",
                      "2023-01":r"C:\Users\JOEF\scorecards\can\2022\december"          
}

rola_directory_dict = {
                      "2021-02":r"C:\Users\JOEF\scorecards\latam\rola\2021\january",
                      "2021-03":r"C:\Users\JOEF\scorecards\latam\rola\2021\february",
                      "2021-04":r"C:\Users\JOEF\scorecards\latam\rola\2021\march",
                      "2021-05":r"C:\Users\JOEF\scorecards\latam\rola\2021\april",
                      "2021-06":r"C:\Users\JOEF\scorecards\latam\rola\2021\may",
                      "2021-07":r"C:\Users\JOEF\scorecards\latam\rola\2021\june",
                      "2021-08":r"C:\Users\JOEF\scorecards\latam\rola\2021\july",
                      "2021-09":r"C:\Users\JOEF\scorecards\latam\rola\2021\august",
                      "2021-10": r"C:\Users\JOEF\scorecards\latam\rola\2021\september",
                      "2021-11":r"C:\Users\JOEF\scorecards\latam\rola\2021\october",
                      "2021-12":r"C:\Users\JOEF\scorecards\latam\rola\2021\november",
                      "2022-01":r"C:\Users\JOEF\scorecards\latam\rola\2021\december",
    
                      "2022-02":r"C:\Users\JOEF\scorecards\latam\rola\2022\january",
                      "2022-03":r"C:\Users\JOEF\scorecards\latam\rola\2022\february",
                      "2022-04":r"C:\Users\JOEF\scorecards\latam\rola\2022\march",
                      "2022-05":r"C:\Users\JOEF\scorecards\latam\rola\2022\april",
                      "2022-06":r"C:\Users\JOEF\scorecards\latam\rola\2022\may",
                      "2022-07":r"C:\Users\JOEF\scorecards\latam\rola\2022\june",
                      "2022-08":r"C:\Users\JOEF\scorecards\latam\rola\2022\july",
                      "2022-09":r"C:\Users\JOEF\scorecards\latam\rola\2022\august",
                      "2022-10": r"C:\Users\JOEF\scorecards\latam\rola\2022\september",
                      "2022-11":r"C:\Users\JOEF\scorecards\latam\rola\2022\october",
                      "2022-12":r"C:\Users\JOEF\scorecards\latam\rola\2022\november",
                      "2023-01":r"C:\Users\JOEF\scorecards\latam\rola\2022\december" 
}


mexico_directory_dict = {
                      "2021-02":r"C:\Users\JOEF\scorecards\latam\mexico\2021\january",
                      "2021-03":r"C:\Users\JOEF\scorecards\latam\mexico\2021\february",
                      "2021-04":r"C:\Users\JOEF\scorecards\latam\mexico\2021\march",
                      "2021-05":r"C:\Users\JOEF\scorecards\latam\mexico\2021\april",
                      "2021-06":r"C:\Users\JOEF\scorecards\latam\mexico\2021\may",
                      "2021-07":r"C:\Users\JOEF\scorecards\latam\mexico\2021\june",
                      "2021-08":r"C:\Users\JOEF\scorecards\latam\mexico\2021\july",
                      "2021-09":r"C:\Users\JOEF\scorecards\latam\mexico\2021\august",
                      "2021-10": r"C:\Users\JOEF\scorecards\latam\mexico\2021\september",
                      "2021-11":r"C:\Users\JOEF\scorecards\latam\mexico\2021\october",
                      "2021-12":r"C:\Users\JOEF\scorecards\latam\mexico\2021\november",
                      "2022-01":r"C:\Users\JOEF\scorecards\latam\mexico\2021\december",
    
                      "2022-02":r"C:\Users\JOEF\scorecards\latam\mexico\2022\january",
                      "2022-03":r"C:\Users\JOEF\scorecards\latam\mexico\2022\february",
                      "2022-04":r"C:\Users\JOEF\scorecards\latam\mexico\2022\march",
                      "2022-05":r"C:\Users\JOEF\scorecards\latam\mexico\2022\april",
                      "2022-06":r"C:\Users\JOEF\scorecards\latam\mexico\2022\may",
                      "2022-07":r"C:\Users\JOEF\scorecards\latam\mexico\2022\june",
                      "2022-08":r"C:\Users\JOEF\scorecards\latam\mexico\2022\july",
                      "2022-09":r"C:\Users\JOEF\scorecards\latam\mexico\2022\august",
                      "2022-10": r"C:\Users\JOEF\scorecards\latam\mexico\2022\september",
                      "2022-11":r"C:\Users\JOEF\scorecards\latam\mexico\2022\october",
                      "2022-12":r"C:\Users\JOEF\scorecards\latam\mexico\2022\november",
                      "2023-01":r"C:\Users\JOEF\scorecards\latam\mexico\2022\december"
}

can_directory_dict = {
                      "2021-02":r"C:\Users\JOEF\scorecards\usa\2021\january",
                      "2021-03":r"C:\Users\JOEF\scorecards\can\2021\february",
                      "2021-04":r"C:\Users\JOEF\scorecards\can\2021\march",
                      "2021-05":r"C:\Users\JOEF\scorecards\can\2021\april",
                      "2021-06":r"C:\Users\JOEF\scorecards\can\2021\may",
                      "2021-07":r"C:\Users\JOEF\scorecards\can\2021\june",
                      "2021-08":r"C:\Users\JOEF\scorecards\can\2021\july",
                      "2021-09":r"C:\Users\JOEF\scorecards\can\2021\august",
                      "2021-10":r"C:\Users\JOEF\scorecards\can\2021\september",
                      "2021-11":r"C:\Users\JOEF\scorecards\can\2021\october",
                      "2021-12":r"C:\Users\JOEF\scorecards\can\2021\november",
                      "2022-01":r"C:\Users\JOEF\scorecards\can\2021\december",
    
                      "2022-02":r"C:\Users\JOEF\scorecards\can\2022\january",
                      "2022-03":r"C:\Users\JOEF\scorecards\can\2022\february",
                      "2022-04":r"C:\Users\JOEF\scorecards\can\2022\march",
                      "2022-05":r"C:\Users\JOEF\scorecards\can\2022\april",
                      "2022-06":r"C:\Users\JOEF\scorecards\can\2022\may",
                      "2022-07":r"C:\Users\JOEF\scorecards\can\2022\june",
                      "2022-08":r"C:\Users\JOEF\scorecards\can\2022\july",
                      "2022-09":r"C:\Users\JOEF\scorecards\can\2022\august",
                      "2022-10":r"C:\Users\JOEF\scorecards\can\2022\september",
                      "2022-11":r"C:\Users\JOEF\scorecards\can\2022\october",
                      "2022-12":r"C:\Users\JOEF\scorecards\can\2022\november",
                      "2023-01":r"C:\Users\JOEF\scorecards\can\2022\december"
}

usa_directory_dict = {
                      "2021-02":r"C:\Users\JOEF\scorecards\usa\2021\january",
                      "2021-03":r"C:\Users\JOEF\scorecards\usa\2021\february",
                      "2021-04":r"C:\Users\JOEF\scorecards\usa\2021\march",
                      "2021-05":r"C:\Users\JOEF\scorecards\usa\2021\april",
                      "2021-06":r"C:\Users\JOEF\scorecards\usa\2021\may",
                      "2021-07":r"C:\Users\JOEF\scorecards\usa\2021\june",
                      "2021-08":r"C:\Users\JOEF\scorecards\usa\2021\july",
                      "2021-09":r"C:\Users\JOEF\scorecards\usa\2021\august",
                      "2021-10":r"C:\Users\JOEF\scorecards\usa\2021\september",
                      "2021-11":r"C:\Users\JOEF\scorecards\usa\2021\october",
                      "2021-12":r"C:\Users\JOEF\scorecards\usa\2021\november",
                      "2022-01":r"C:\Users\JOEF\scorecards\usa\2021\december",
    
                      "2022-02":r"C:\Users\JOEF\scorecards\usa\2022\january",
                      "2022-03":r"C:\Users\JOEF\scorecards\usa\2022\february",
                      "2022-04":r"C:\Users\JOEF\scorecards\usa\2022\march",
                      "2022-05":r"C:\Users\JOEF\scorecards\usa\2022\april",
                      "2022-06":r"C:\Users\JOEF\scorecards\usa\2022\may",
                      "2022-07":r"C:\Users\JOEF\scorecards\usa\2022\june",
                      "2022-08":r"C:\Users\JOEF\scorecards\usa\2022\july",
                      "2022-09":r"C:\Users\JOEF\scorecards\usa\2022\august",
                      "2022-10":r"C:\Users\JOEF\scorecards\usa\2022\september",
                      "2022-11":r"C:\Users\JOEF\scorecards\usa\2022\october",
                      "2022-12":r"C:\Users\JOEF\scorecards\usa\2022\november",
                      "2023-01":r"C:\Users\JOEF\scorecards\usa\2022\december"
}

disty_names_to_fix = []  #list of distributor names that were found in an scorecard, but not within the excel list
attachment_name_list = [] 
try:
    for email in score_cards_inbox.Items: #emails in gwp inbox
        try:
            s = email.sender
            for attachment in email.Attachments:
                attachment_name = attachment.FileName.lower()
                #check if the attachment_name is infact a scorecard and that it is a pdf file
                if "scorecard" in attachment_name and fileTypeCheck(attachment_name) == True:
                    email_date = email.senton.date()
                    email_month = calcMonth(email_date)              
                    if distyFileName(attachment_name) != None:
                        attachment_disty_name = distyFileName(attachment_name)
                        region = determineRegion(attachment_disty_name,disty_region_dict)
                        attachment_name_list.append(distyFileName(attachment_name))                        
                        #if the determineRegion function couldn't find the region
                        if region == attachment_disty_name:
                            disty_names_to_fix.append(region)
                            print("This Disty's name isn't being picked up: %s" % region)
                        else:
                            #region found
                            fileAttachment(attachment,email_month,region)
                                
        except Exception as e:
            print("error when saving the attachment." + str(e))
except Exception as e:
    print("error when processing emails messages:" + str(e))
    
if (len(disty_names_to_fix) > 0):
    print("within the file stored at %s, you will need to add the distributor names and regions that are listed below: \n" %disty_region_spreadsheet_path)
    print(disty_names_to_fix)
