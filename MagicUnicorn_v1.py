"""
MagicUnicorn_v1.py - Parse and return data from Microsoft Office 365 Activities API reports
Copyright (C) <2018>  <LMG Security>

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program.  If not, see <https://www.gnu.org/licenses/>.

"""

import argparse
import string
import csv
import os
import sys
from ast import literal_eval

parser = argparse.ArgumentParser(description='Parse Office 365 Outlook Activities')
parser.add_argument('-i', metavar='<input_file>', help='Input text file', required=True)
parser.add_argument('-t', metavar='<title>', help='Main report title', required=True)
parser.add_argument('-o', metavar='<output>', help='Output target directory', required=True)
args = parser.parse_args()

in_file = open(args.i, "r",encoding='utf-8')
rep_title = args.t
out_file_dir = args.o

##in_file = open(args.i, "r")
##out_file = open(args.o, "w+")

typesIWant = ['Web', 'Exchange', 'Outlook','Mobile']

message_ids = {}
reads_by_id = {}
reads_by_time = ["Date\tTime\tType\tPlatform\tSubject\tFrom\tReceived Date\tLast Login IP\n"]
attachments_by_time = ["Date\tTime\tPlatform\tSubject\tFrom\tReceived Date\tAction\tLast Login IP\tActivity Item ID\n"]
ips_found = ["No Data"]
message_ids_found = []
search_by_time = ["Date\tTime\tPlatform\tQuery\tLast Login ip\n"] 
line_num = 0 

print("Building mail dictionary")
## First pass to build the mail dictionary
for line in in_file:
    line_num+=1
    split_line = line.split(',')
    time_and_date = split_line[0].split("T")
    date_stamp = time_and_date[0]
    time_stamp = time_and_date[1]        
    app_type = split_line[1]
    activity_type = split_line[2]
    activity_item_id = split_line[3] ##blank for some items
    activity_creation_time = split_line[4]
    session_id = split_line[5]
    custom_properties = {}
    if(app_type in typesIWant):

        if(activity_type=='MessageDelivered' and app_type=='Exchange'):            
            
            try:
                message_ids_found.append(activity_item_id)
                dict_string = ','.join(split_line[6:])          
                dict_string = dict_string.translate({ord(c): None for c in '\";'})                         
                custom_properties = literal_eval(dict_string)
                message_ids[activity_item_id] = {'ConversationId': custom_properties['ConversationId'], 'InternetMessageId': custom_properties['InternetMessageId'],'Subject': custom_properties['Subject'], 'From': custom_properties['SenderSmtpAddress'], 'Received': custom_properties[ 'ReceivedTime'], 'Folder': custom_properties['DeliveredFolderType']}
                
            except Exception as e:
                temp_dict_string = dict_string.split(',')                
                message_ids[activity_item_id] = {'ConversationId': 'No Data', 'InternetMessageId': 'No Data','Subject': 'No Data', 'From': 'No Data', 'Received': 'No Data', 'Folder': 'No Data'}
                for item in temp_dict_string:
                    item = item.translate({ord(c): None for c in '\";'})
                    dict_items = item.split(":")
                    for thing in dict_items:
                        if('Subject' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['Subject'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                            except:
                                pass

                        if('Received' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['Received'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                            except:
                                pass
                                
                        if('Conversation' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['ConversationId'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                            except:
                                pass

                        if('SenderSmtpAddress' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['From'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                            except:
                                pass

                        if('InternetMessageId' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['InternetMessageId'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                            except:
                                pass
                        
                        if('DeliveredFolderType' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['Folder'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                            except:
                                pass

        if(activity_type=='MarkAsRead' and (app_type=='Mobile' or app_type=='Outlook')):
            message_ids_found.append(activity_item_id)
            dict_string = ','.join(split_line[6:])          
            dict_string = dict_string.translate({ord(c): None for c in '\";'})

            if(activity_item_id in message_ids.keys()):
                pass

            else:
                if('IPM.Schedule.Meeting.Request' in dict_string):
                    pass
                else:
                    message_ids[activity_item_id] = {'ConversationId': 'No Data', 'InternetMessageId': 'No Data', 'Subject': 'No Data', 'From': 'No Data', 'Received': 'No Data', 'Folder': 'No Data'}
                try:                                              
                    custom_properties = literal_eval(dict_string)
                    message_ids[activity_item_id]['Received'] = custom_properties['ReceivedTime']
                    message_ids[activity_item_id]['From'] = custom_properties['SenderAddress']
                    message_ids[activity_item_id]['InternetMessageId'] = custom_properties['InternetMessageId']
                    message_ids[activity_item_id]['Folder'] = custom_properties['SourceFolder']
                except:
                    temp_dict_string = dict_string.split(',')
                    for item in temp_dict_string:
                        item = item.translate({ord(c): None for c in '\";'})
                        dict_items = item.split(":")
                        for thing in dict_items:
                            if('ReceivedTime' in str(thing[1:])):
                                try:
                                    output_dict = ':'.join(dict_items[1:])                            
                                    message_ids[activity_item_id]['Received'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                                except:
                                    pass
                            if('InternetMessageId' in str(thing[1:])):
                                try:
                                    output_dict = ':'.join(dict_items[1:])                            
                                    message_ids[activity_item_id]['InternetMessageId'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                                except:
                                    pass
                            if('SourceFolder' in str(thing[1:])):
                                try:
                                    output_dict = ':'.join(dict_items[1:])                            
                                    message_ids[activity_item_id]['Folder'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                                except:
                                    pass                            
                            if('SenderAddress' in str(thing[1:])):
                                try:
                                    output_dict = ':'.join(dict_items[1:])                            
                                    message_ids[activity_item_id]['From'] = output_dict[1:].translate({ord(c): None for c in '\"\';'})
                                except:
                                    pass

        if(activity_type=='ReplyAll'):
            message_ids_found.append(activity_item_id)
            dict_string = ','.join(split_line[6:])          
            dict_string = dict_string.translate({ord(c): None for c in '\";'})
            if(activity_item_id in message_ids.keys()):
                pass
            else:
                message_ids[activity_item_id] = {'ConversationId': 'No Data', 'InternetMessageId': 'No Data', 'Subject': 'No Data', 'From': 'No Data', 'Received': 'No Data', 'Folder': 'No Data'}
            try:
                custom_properties = literal_eval(dict_string)
                message_ids[activity_item_id]['Received'] = custom_properties['ReceivedTime']
                message_ids[activity_item_id]['From'] = custom_properties['SenderAddress']
                message_ids[activity_item_id]['InternetMessageId'] = custom_properties['InternetMessageId']
                message_ids[activity_item_id]['Folder'] = custom_properties['SourceDefaultFolderType']
            except:
                temp_dict_string = dict_string.split(',')
                for item in temp_dict_string:
                    item = item.translate({ord(c): None for c in '\";'})
                    dict_items = item.split(":")
                    for thing in dict_items:
                        if('ReceivedTime' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['Received'] = output_dict[1:].translate({ord(c): None for c in '}{\"\';'}).rstrip()
                            except:
                                pass
                        
                        if('InternetMessageId' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['InternetMessageId'] = output_dict[1:].translate({ord(c): None for c in '}{\"\';'}).rstrip()
                            except:
                                pass
                        
                        if('SourceDefaultFolderType' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['Folder'] = output_dict[1:].translate({ord(c): None for c in '}{\"\';'}).rstrip()
                            except:
                                pass
                        
                        if('SenderAddress' in str(thing[1:])):
                            try:
                                output_dict = ':'.join(dict_items[1:])                            
                                message_ids[activity_item_id]['From'] = output_dict[1:].translate({ord(c): None for c in '}{\"\';'}).rstrip()
                            except:
                                pass


in_file.close()
line_num = 0  
logins_found = ["Date\tTime\tSource IP\tLogon Type\tApp Type\tUser\tClient"]
in_file = open(args.i, "r",encoding='utf-8')
print("Parsing activities")
for line in in_file:    
    line_num+=1
    split_line = line.split(',')
    time_and_date = split_line[0].split("T")
    date_stamp = time_and_date[0]
    time_stamp = time_and_date[1]        
    app_type = split_line[1]
    activity_type = split_line[2]
    activity_item_id = split_line[3] ##blank for some items
    activity_creation_time = split_line[4]
    session_id = split_line[5]
    custom_properties = {}

    if(app_type in typesIWant):

        if(activity_type=='ServerLogon'):
            dict_string = ','.join(split_line[6:])          
            dict_string = dict_string.translate({ord(c): None for c in '\";'})
            custom_properties = literal_eval(dict_string)
            logins_found.append(date_stamp + "\t" + time_stamp + "\t" + custom_properties['ClientIP'] + "\t" + activity_type + "\t" + app_type + "\t" + custom_properties['UserName'] + "\t" + custom_properties['UserAgent'] + "\n")
            if(custom_properties['ClientIP'] not in ips_found):
                ips_found.append(custom_properties['ClientIP'])
        
        if(activity_type=='Logon'):
            dict_string = ','.join(split_line[6:])          
            dict_string = dict_string.translate({ord(c): None for c in '\";'})
            custom_properties = literal_eval(dict_string)
            logins_found.append(date_stamp + "\t" + time_stamp + "\t" + custom_properties['IPAddress'] + "\t" + activity_type + "\t" + app_type + "\t" + " " + "\t" + custom_properties['Browser'] + "\n")
            if(custom_properties['IPAddress'] not in ips_found):
                ips_found.append(custom_properties['IPAddress'])
            
        if('ReadingPane' in activity_type):
            if(activity_item_id not in message_ids.keys()):
                message_ids[activity_item_id] = {'ConversationId': 'No Data', 'InternetMessageId': 'No Data', 'Subject': 'No Data', 'From': 'No Data', 'Received': 'No Data', 'Folder': 'No Data'}

            if(activity_item_id in reads_by_id.keys()):
                logged_reads = reads_by_id[activity_item_id]
                reads_by_id[activity_item_id] = logged_reads + (date_stamp + "\t" + time_stamp + "\t" + activity_type +"\t" + app_type + "\t" + message_ids[activity_item_id]['Subject'] + "\t" + message_ids[activity_item_id]['From'] + "\t" + message_ids[activity_item_id]['Received'] + ips_found[len(ips_found)-1] + "\n")
            else:
                reads_by_id[activity_item_id] =  "Date\tTime\tType\tPlatform\tSubject\tFrom\tReceived Date\tLast Login IP\n" + (date_stamp + "\t" + time_stamp + "\t" + activity_type + "\t" + app_type + "\t" + message_ids[activity_item_id]['Subject'] + "\t" + message_ids[activity_item_id]['From'] + "\t" + message_ids[activity_item_id]['Received'] + ips_found[len(ips_found)-1] + "\n")
            reads_by_time.append(date_stamp + "\t" + time_stamp + "\t" + activity_type +"\t" + app_type + "\t" + message_ids[activity_item_id]['Subject'] + "\t" + message_ids[activity_item_id]['From'] + "\t" + message_ids[activity_item_id]['Received'] + "\t" + ips_found[len(ips_found)-1] + "\n")

        if(activity_type=='SearchResult'):
            try:
                dict_string = ','.join(split_line[6:])          
                dict_string = dict_string.translate({ord(c): None for c in '\";'})
                custom_properties = literal_eval(dict_string)
                search_by_time.append(date_stamp + "\t" + time_stamp + "\t" + app_type + "\t" + custom_properties['Query'] + "\t" + ips_found[len(ips_found) -1] + "\n")
            except:
                temp_dict_string = dict_string.split(',')
                for item in temp_dict_string:
                    item = item.translate({ord(c): None for c in '\";'})
                    dict_items = item.split(":")
                    for thing in dict_items:
                        if('Query' in str(thing[1])):
                            try:
                                output_dict = ':'.join(dict_items[1:])
                                search_by_time.append(date_stamp + "\t" + time_stamp + "\t" + app_type + "\t" + output_dict[1:].translate({ord(c): None for c in '}{\"\';'}).rstrip() + "\t" + ips_found[len(ips_found) -1] + "\n")                                                            
                            except:
                                pass                            

        if(activity_type=='SearchSuggestionsDisplay'):
            try:
                dict_string = ','.join(split_line[6:])          
                dict_string = dict_string.translate({ord(c): None for c in '\";'})
                custom_properties = literal_eval(dict_string)
                search_by_time.append(date_stamp + "\t" + time_stamp + "\t" + app_type + "\t" + custom_properties['SuggestionStimulus'] + "\t" + ips_found[len(ips_found) -1] + "\n")
            except:
                temp_dict_string = dict_string.split(',')
                for item in temp_dict_string:
                    item = item.translate({ord(c): None for c in '\";'})
                    dict_items = item.split(":")
                    for thing in dict_items:
                        if('SuggestionStimulus' in str(thing[1])):
                            try:
                                output_dict = ':'.join(dict_items[1:])
                                search_by_time.append(date_stamp + "\t" + time_stamp + "\t" + app_type + "\t" + output_dict[1:].translate({ord(c): None for c in '}{\"\';'}).rstrip() + "\t" + ips_found[len(ips_found) -1] + "\n")                                                            
                            except:
                                pass

        if(activity_type=='OpenedAnAttachment'):
            if(activity_item_id not in message_ids.keys()):
                message_ids[activity_item_id] = {'ConversationId': 'No Data', 'InternetMessageId': 'No Data', 'Subject': 'No Data', 'From': 'No Data', 'Received': 'No Data', 'Folder': 'No Data'}
            dict_string = ','.join(split_line[6:])          
            dict_string = dict_string.translate({ord(c): None for c in '\";'})
            custom_properties = literal_eval(dict_string)    
            attachments_by_time.append(date_stamp + "\t" + time_stamp + "\t" + app_type + "\t" + message_ids[activity_item_id]['Subject'] + "\t" + message_ids[activity_item_id]['From'] + "\t" + message_ids[activity_item_id]['Received'] + "\t" + custom_properties['AttachmentAction'] + "\t" + ips_found[len(ips_found)-1] + "\t" + activity_item_id + "\n")
            


##Print Reports
print("Generating reports")
file_name = args.o + args.t + "-attachments-activity.tsv"
out_file= open(file_name, "w+",encoding='utf-8')
for item in attachments_by_time:
    out_file.write(item)
out_file.close()

file_name = args.o + args.t + "-search-activity.tsv"
out_file= open(file_name, "w+",encoding='utf-8')
for item in search_by_time:
    out_file.write(item)
out_file.close()

file_name = args.o + args.t + "-read-activity-by-time.tsv"
out_file= open(file_name, "w+",encoding='utf-8')
for item in reads_by_time:
    out_file.write(item)
out_file.close()

file_name = args.o + args.t + "-read-activity-by-item.tsv"
out_file= open(file_name, "w+",encoding='utf-8')
for keys in reads_by_id.keys():
    out_file.write(keys + "\n")
    out_file.write(reads_by_id[keys])
out_file.close()

file_name = args.o + args.t + "-logon-activity.tsv"
out_file= open(file_name, "w+",encoding='utf-8')
for entry in logins_found:
    out_file.write(entry)
out_file.close()

print("Completed")