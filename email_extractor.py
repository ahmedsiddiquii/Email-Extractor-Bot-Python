import imaplib
import email
import time
import datetime
import pandas as pd
import os
from datetime import date, timedelta
import threading
import os.path
from os import path
from time import sleep
#==================================================Find Date Ranges and Give Values==================================================

EMAIL ="email"
PASSWORD= "password"


startdate = input("Enter start date (mm/dd/yy) : ")
enddate = input("Enter End date (mm/dd/yy) : ")
str_A = input("Enter Start Time (hh:mm AM) : ")
str_B = input("Enter End Time (hh:mm AM) : ")
output_path=input("Enter Folder Path : ")

exec_time = time.time()

def startdate_split_fun(startdate):
    startdate_split = startdate.split('/')
    s_date = int(startdate_split[1])
    s_month = int(startdate_split[0])
    s_year = int(startdate_split[2])    
    return s_year,s_month,s_date

def enddate_split_fun(enddate):
    enddate_split = enddate.split('/')
    e_date = int(enddate_split[1])
    e_month = int(enddate_split[0])
    e_year = int(enddate_split[2])
    return e_year,e_month,e_date

all_dates = []
def date_range(sdate,edate):
    delta = edate - sdate
    for i in range(delta.days + 1):
        day = sdate + timedelta(days=i)
        day_split = str(day).split('-')
        dates = day_split[2]
        month = day_split[1]
        year = day_split[0]                        
        month = datetime.date(1900, int(month[1:2]), 1).strftime('%B')                                
        complete_date = month[0:3]+"/"+dates+"/"+year
        all_dates.append(complete_date)
        
sdate = date(startdate_split_fun(startdate)[0],startdate_split_fun(startdate)[1],startdate_split_fun(startdate)[2])

edate = date(enddate_split_fun(enddate)[0],enddate_split_fun(enddate)[1],enddate_split_fun(enddate)[2])
edate2 = edate + datetime.timedelta(days=1)
sdate2 = sdate + datetime.timedelta(days=-1)

sdate_str = sdate.strftime("%d-%b-%Y")
edate_str = edate2.strftime("%d-%b-%Y")

date_range(sdate,edate)

#=================================================Find time ranges==================================================#
time_ranges = []
def getUser_time(str_A,str_B):    
    # Create our datetime objects
    A = datetime.datetime.strptime(str_A,"%I:%M %p")
    B = datetime.datetime.strptime(str_B,"%I:%M %p")
    tmp = A
    count=0
    if str(tmp.time()) <=str(B.time()):
        while tmp <=B:
            x = str(tmp.time())     
            time_ranges.append(x[:5])        
            tmp = tmp + timedelta(minutes=1)
    else:
        while True :
            x = str(tmp.time())     
            time_ranges.append(x[:5])        
            tmp = tmp + timedelta(minutes=1)
            
            if str(tmp.time()) ==str(B.time()):
                x = str(B.time())     
                time_ranges.append(x[:5]) 
                break
getUser_time(str_A,str_B)

# ==================================================Email Retriver==================================================#

all_emails = []

SERVER = 'pop3.live.com'
mail = imaplib.IMAP4_SSL(SERVER)
mail.login(EMAIL, PASSWORD)
mail.select('inbox')
status, data = mail.search(None, 'since '+sdate_str+" before "+edate_str)
mail_ids = []
delete_emails=[]
def add_mail_ids(block):
    global mail_ids
    mail_ids += block.split()
    mail_id=block.split()
    #m2=threading.Thread(target=run,args=(mail_id,))
    #m2.start()
def run(status,data,emaill):
    global mail
    global output_path
    global startdate
    global enddate
    global all_dates
    global time_ranges
    global all_emails
    global delete_emails
       
    for response_part in data:    
        if isinstance(response_part, tuple):    
            message = email.message_from_bytes(response_part[1])                        
            mail_from = message['from']
            mail_subject = message['subject']                                                
            mail_date = message['date'].split()    
            
           
            
            
            
            date = mail_date[1]
            month = mail_date[2]
            year = mail_date[3]
            timee = (mail_date[4])[0:5]
           
            complete_date = month+"/"+date+"/"+year
            try:
                filePath=filePath
                pass
            except:
                filePath="No File Found"
            
                    
            if (complete_date in all_dates) and (timee in time_ranges):
                print(complete_date)
                print(timee)
                print('\n')
                data = {
                                'Subject':mail_subject,
                                'Date': complete_date,
                                'Time': timee,
                }
                all_emails.append(data)
                delete_emails.append(emaill)
                     
            else:
                pass
thread1=[]
thread2=[]
print("Please Wait..")  
for block in data:
    m=threading.Thread(target=add_mail_ids,args=(block,))
    thread1.append(m)
    m.start()
for x in thread1:
    x.join()
    
for i in mail_ids:
    status, data = mail.fetch(i, '(BODY.PEEK[HEADER.FIELDS (FROM TO CC SUBJECT DATE)])') 
    m2=threading.Thread(target=run,args=(status,data,i,))
    thread2.append(m2)
    m2.start()

for x2 in thread2:
    x2.join()

for i in delete_emails:
    
    mail.store(i, "+FLAGS","\\Deleted")
mail.expunge()
print("Email Moved !")  

import pandas as pd
df = pd.DataFrame(all_emails)
PATH=output_path+"/"+str(startdate).replace("/","-")+"--"+str(enddate).replace("/","-")+'--Output.csv'
PATH_D=PATH
st=path.exists(PATH)
if st == True:
    c=1
    while True:
        PATH=PATH_D.replace(".csv","")
        PATH+=str('(')+str(c)+").csv"
        st=path.exists(PATH)
        if st==False:
            break
        if st==True:
            c+=1
try:
    df.to_csv(PATH)
except Exception as e:
    print(e)
    df.to_csv('output.csv')
print('saved\n')

print("Execution Time is %s seconds ---" % round((time.time() - exec_time)))