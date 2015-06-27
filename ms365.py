#!/usr/local/bin/python

import urllib2
import json
import base64
import re
import datetime
import argparse


#vars
parser = argparse.ArgumentParser()
parser.add_argument("-u", "--Username", help="State your username(your email)")
parser.add_argument("-p", "--Password", help="State your password")
parser.add_argument("-t", "--Time", help="Time frame(today,tommorow) - Defualt: today")
parser.add_argument("-tz", "--TimeZone", help="Time Zone EST OR EDT")
args = parser.parse_args()
username= args.Username
password= args.Password
time= args.Time
tz =args.TimeZone

today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
starthr="T07:00:00Z"
add="&endDateTime="
filename ="meetings.txt"
file = open(filename, 'w')
subject = "Subject"
values = "value"
start ="Start"
offset =""
baseurl="https://outlook.office365.com/api/v1.0/Me/CalendarView/?startDateTime="
today = datetime.date.today()
tomorrow = today + datetime.timedelta(days=1)
dayaftertomorrow = today + datetime.timedelta(days=2)
urltoday = baseurl+str(today)+add+str(tomorrow)
urltomorrow = baseurl+str(tomorrow)+add+str(dayaftertomorrow)

def usage():
    print "Usage:"
    print "365cal.py [-u] usearname@domain.com [-p] password [-t] time"
    print "Example:"
    print "365cal.py -u jhon@office365.com -p password -t today"
    print " "
    print "Currently the following time supported:"
    print "1.today"
    print "2.tomorrow"

#REGEX
re1='.*?'    # Non-greedy match on filler
re2='((?:(?:[0-1][0-9])|(?:[2][0-3])|(?:[0-9])):(?:[0-5][0-9])(?::[0-5][0-9])?(?:\\s?(?:am|AM|pm|PM))?)'    # HourMinuteSec 1
rg = re.compile(re1+re2,re.IGNORECASE|re.DOTALL)
hoursedt = {"00:00:00": "8 PM",
         "00:15:00": "8 fifteen PM",
         "00:30:00": "8 thirty PM",
         "00:45:00": "8 forty five PM",
                 "01:00:00": "9 PM",
                 "01:15:00": "9 fifteen PM",
                 "01:30:00": "9 thirty PM",
                 "01:45:00": "9 forty five PM",
                 "02:00:00": "10 PM",
                 "02:15:00": "10 fifteen PM",
                 "02:30:00": "10 thirty PM",
                 "02:45:00": "10 forty five PM",                 
                 "03:00:00": "11 PM",
                 "03:15:00": "11 fifteen PM",
                 "03:30:00": "11 thirty PM",
                 "03:45:00": "11 forty five PM",
                 "04:00:00": "12 PM",
                 "04:15:00": "12 fifteen PM",
                 "04:30:00": "12 thirty PM",
                 "04:45:00": "12 forty five PM",                 
                 "05:00:00": "1 AM",
                 "05:15:00": "1 fifteen AM",
                 "05:30:00": "1 thirty AM",
                 "05:45:00": "1 forty five AM",
                 "06:00:00": "2 AM",
                 "06:15:00": "2 fifteen AM",
                 "06:30:00": "2 thirty AM",
                 "06:45:00": "2 forty five AM",                 
                 "07:00:00": "3 AM",
                 "07:15:00": "3 fifteen AM",
                 "07:30:00": "3 thirty AM",
                 "07:45:00": "3 forty five AM",
                 "08:00:00": "4 AM",
                 "08:15:00": "4 fifteen AM",
                 "08:30:00": "4 thirty AM",
                 "08:45:00": "4 forty five AM",                 
                 "09:00:00": "5 AM",
                 "09:15:00": "5 fifteen AM",
                 "09:30:00": "5 thirty AM",
                 "09:45:00": "5 forty five AM",
                 "10:00:00": "6 AM",
                 "10:15:00": "6 fifteen AM",
                 "10:30:00": "6 thirty AM",
                 "10:45:00": "6 forty five AM",                 
                 "11:00:00": "7 AM",
                 "11:15:00": "7 fifteen AM",
                 "11:30:00": "7 thirty AM",
                 "11:45:00": "7 forty five AM",
                 "12:00:00": "8 AM",
                 "12:15:00": "8 fifteen AM",
                 "12:30:00": "8 thirty AM",
                 "12:45:00": "8 forty five AM",                 
                 "13:00:00": "9 AM",
                 "13:15:00": "9 fifteen AM",
                 "13:30:00": "9 thirty AM",
                 "13:45:00": "9 forty five AM",
                 "14:00:00": "10 AM",
                 "14:15:00": "10 fifteen AM",
                 "14:30:00": "10 thirty AM",
                 "14:45:00": "10 forty five AM",                 
                 "15:00:00": "11 AM",
                 "15:15:00": "11 fifteen AM",
                 "15:30:00": "11 thirty AM",
                 "15:45:00": "11 forty five AM",
                 "16:00:00": "12 AM",
                 "16:15:00": "12 fifteen AM",
                 "16:30:00": "12 thirty AM",
                 "16:45:00": "12 forty five AM",                 
                 "17:00:00": "1 PM",
                 "17:15:00": "1 fifteen PM",
                 "17:30:00": "1 thirty PM",
                 "17:45:00": "1 forty five PM",
                 "18:00:00": "2 PM",
                 "18:15:00": "2 fifteen PM",
                 "18:30:00": "2 thirty PM",
                 "18:45:00": "2 forty five PM",                 
                 "19:00:00": "3 PM",
                 "19:15:00": "3 fifteen PM",
                 "19:30:00": "3 thirty PM",
                 "19:45:00": "3 forty five PM",
                 "20:00:00": "4 PM",
                 "20:15:00": "4 fifteen PM",
                 "20:30:00": "4 thirty PM",
                 "20:45:00": "4 forty five PM",                 
                 "21:00:00": "5 PM",
                 "21:15:00": "5 fifteen PM",
                 "21:30:00": "5 thirty PM",
                 "21:45:00": "5 forty five PM",
                 "22:00:00": "6 PM",
                 "22:15:00": "6 fifteen PM",
                 "22:30:00": "6 thirty PM",
                 "22:45:00": "6 forty five PM",                 
                 "23:00:00": "7 PM",
                 "23:15:00": "7 fifteen PM",
                 "23:30:00": "7 thirty PM",
                 "23:45:00": "7 forty five PM",                 
                 }

hoursest = {"00:00:00": "7 PM",
         "00:15:00": "7 fifteen PM",
         "00:30:00": "7 thirty PM",
         "00:45:00": "7 forty five PM",
                 "01:00:00": "8 PM",
                 "01:15:00": "8 fifteen PM",
                 "01:30:00": "8 thirty PM",
                 "01:45:00": "8 forty five PM",
                 "02:00:00": "9 PM",
                 "02:15:00": "9 fifteen PM",
                 "02:30:00": "9 thirty PM",
                 "02:45:00": "9 forty five PM",                 
                 "03:00:00": "10 PM",
                 "03:15:00": "10 fifteen PM",
                 "03:30:00": "10 thirty PM",
                 "03:45:00": "10 forty five PM",
                 "04:00:00": "11 PM",
                 "04:15:00": "11 fifteen PM",
                 "04:30:00": "11 thirty PM",
                 "04:45:00": "11 forty five PM",                 
                 "05:00:00": "12 AM",
                 "05:15:00": "12 fifteen AM",
                 "05:30:00": "12 thirty AM",
                 "05:45:00": "12 forty five AM",
                 "06:00:00": "1 AM",
                 "06:15:00": "1 fifteen AM",
                 "06:30:00": "1 thirty AM",
                 "06:45:00": "1 forty five AM",                 
                 "07:00:00": "2 AM",
                 "07:15:00": "2 fifteen AM",
                 "07:30:00": "2 thirty AM",
                 "07:45:00": "2 forty five AM",
                 "08:00:00": "3 AM",
                 "08:15:00": "3 fifteen AM",
                 "08:30:00": "3 thirty AM",
                 "08:45:00": "3 forty five AM",                 
                 "09:00:00": "4 AM",
                 "09:15:00": "4 fifteen AM",
                 "09:30:00": "4 thirty AM",
                 "09:45:00": "4 forty five AM",
                 "10:00:00": "5 AM",
                 "10:15:00": "5 fifteen AM",
                 "10:30:00": "5 thirty AM",
                 "10:45:00": "5 forty five AM",                 
                 "11:00:00": "6 AM",
                 "11:15:00": "6 fifteen AM",
                 "11:30:00": "6 thirty AM",
                 "11:45:00": "6 forty five AM",
                 "12:00:00": "7 AM",
                 "12:15:00": "7 fifteen AM",
                 "12:30:00": "7 thirty AM",
                 "12:45:00": "7 forty five AM",                 
                 "13:00:00": "8 AM",
                 "13:15:00": "8 fifteen AM",
                 "13:30:00": "8 thirty AM",
                 "13:45:00": "8 forty five AM",
                 "14:00:00": "9 AM",
                 "14:15:00": "9 fifteen AM",
                 "14:30:00": "9 thirty AM",
                 "14:45:00": "9 forty five AM",                 
                 "15:00:00": "10 AM",
                 "15:15:00": "10 fifteen AM",
                 "15:30:00": "10 thirty AM",
                 "15:45:00": "10 forty five AM",
                 "16:00:00": "11 AM",
                 "16:15:00": "11 fifteen AM",
                 "16:30:00": "11 thirty AM",
                 "16:45:00": "11 forty five AM",                 
                 "17:00:00": "12 PM",
                 "17:15:00": "12 fifteen PM",
                 "17:30:00": "12 thirty PM",
                 "17:45:00": "12 forty five PM",
                 "18:00:00": "1 PM",
                 "18:15:00": "1 fifteen PM",
                 "18:30:00": "1 thirty PM",
                 "18:45:00": "1 forty five PM",                 
                 "19:00:00": "2 PM",
                 "19:15:00": "2 fifteen PM",
                 "19:30:00": "2 thirty PM",
                 "19:45:00": "2 forty five PM",
                 "20:00:00": "3 PM",
                 "20:15:00": "3 fifteen PM",
                 "20:30:00": "3 thirty PM",
                 "20:45:00": "3 forty five PM",                 
                 "21:00:00": "4 PM",
                 "21:15:00": "4 fifteen PM",
                 "21:30:00": "4 thirty PM",
                 "21:45:00": "4 forty five PM",
                 "22:00:00": "5 PM",
                 "22:15:00": "5 fifteen PM",
                 "22:30:00": "5 thirty PM",
                 "22:45:00": "5 forty five PM",                 
                 "23:00:00": "6 PM",
                 "23:15:00": "6 fifteen PM",
                 "23:30:00": "6 thirty PM",
                 "23:45:00": "6 forty five PM",                 
                 }

def get_cal(url,tz):    
    request = urllib2.Request(url)
    base64string = base64.encodestring('%s:%s' % (username, password)).replace('\n', '')
    request.add_header("Authorization", "Basic %s" % base64string)   
    result = urllib2.urlopen(request)
    data =  json.load(result)
    if data[values]:
       file.write("Here is the meetings informatio for ")
       file.write(time)
       file.write(",")
       file.write("\n")
       for value in data[values]:
           m = rg.search((value[start]))
           time1 = m.group(1)
           humentime = hoursest[time1]
           meetings = value[subject]
           if meetings is not None:          
               meetings = meetings + "," +" " + "Starting at"  + ","+ " " + str(humentime) #  str(value["Start"])
               #print value["Subject"], "Starting at", value["Start"] 
               file.write(meetings)
               file.write("\n")               
               file.write("Next,")
               file.write("\n")
           else:
               file.write("It seems that you don't have mettings")
               file.write(time)
               file.write(",")
               file.write("\n")
               file.write("\n")
       file.write("Thats it for")
       file.write(time)
    else:
         file.write("It seems that you don't have mettings")
         file.write(time)
         file.write(",")
         file.write("\n")
         file.write("\n")

if args.Username is None:
    usage()
    print "Exit reason, -u is empty"
elif args.Password is None:
    usage()
    print "Exit reason, -p is empty"
elif args.Time is None:
    usage()
    print "Exit reason, -t is empty"
elif args.TimeZone is None:
    usage()
    print "Exit reason, -tz is empty"    
else:
    if tz.lower() == "est":
        timezone ="hoursest"
    elif tz.lower() == "edt":
        timezone ="hoursedt"
    else:
        print "-tz is not recognise, please set it to est or edt"
        exit()
    if time.lower() == "today":
        get_cal(urltoday,timezone)
    elif time.lower() == "tomorrow":
        timepiriod = args.Time
        #global offset      
        get_cal(urltomorrow,timezone)
    else:
        print "Error can't process the 'time' variable, please make sure you state it correct(today OR tomorrow)"   