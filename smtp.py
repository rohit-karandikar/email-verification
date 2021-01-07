import pandas as pd
import dns.resolver
import socket
import smtpd
import smtplib
import os
import time
import re
import ssl
from dns.resolver import NXDOMAIN
from smtplib import SMTPServerDisconnected
start = time.time()
def checkmx(a):
    try:
        dns.resolver.resolve(a, 'MX')
    except NXDOMAIN:
        mxRecord = "1"
    else :
        records = dns.resolver.resolve(a, 'MX')
        mxRecord = str(records[0].exchange)


start = time.time()
a = pd.read_excel("test.xlsx")
b = pd.DataFrame()
e = []
final_frame = pd.DataFrame()
email_exist = []
b["finalemails"] = list(dict.fromkeys(a["emails"]))
for i in b["finalemails"]:
    addressToVerify = i
    match = re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', addressToVerify)
    if match != None :
        e.append(i)    
c = []
d = []
rolebased = []
mx_check = []
temp2 = []
defaulters = []
temp1 = pd.read_excel("rolebased.xlsx")
temp1 = list(temp1["roles"])
for i in e:
    c.append((i.split("@"))[0])
    d.append((i.split("@"))[1])
for i in range(len(c)):
    if c[i] in temp1:
        rolebased.append("True")
    else:
        rolebased.append("False")
d, c = zip(*sorted(zip(d, c)))
records = dns.resolver.resolve(d[0], 'MX')
server = smtplib.SMTP("smtp.gmail.com",587)
server.login("rohitk.techfest","homibhabha")
for i in range(1,len(d)):
    mxRecord = str(records[0].exchange)
    server.set_debuglevel(1)
    server.connect(mxRecord)
    (code,resp) = server.helo(server.local_hostname)
    if code == 250:
        (code,resp) = server.mail("rohit.techfest@gmail.com")
        if code == 250:
            (code, resp) = server.rcpt(c[i-1] + "@" + d[i-1])
            if code == 250:
                email_exist.append("exists")
            else:
                email_exist.append("does not exist")
        else:
            email_exist.append("exists")
            defaulters.append(c[i-1] + "@" + d[i-1])
    else:
        email_exist.append("does not exist")
    if d[i] != d[i-1]:
        records = dns.resolver.resolve(d[i], 'MX')           
mxRecord = str(records[0].exchange)
server = smtplib.SMTP()
server.set_debuglevel(0)
server.connect(mxRecord)
(code,resp) = server.helo(server.local_hostname)
if code == 250:
    (code,resp) = server.mail("rohit.techfest@gmail.com")
    if code == 250:
        (code, resp) = server.rcpt(c[-1] + "@" + d[-1])
        if code == 250:
            email_exist.append("exists")
        else:
            email_exist.append("does not exist")
    else:
        email_exist.append("exists")
        defaulters.append(c[-1] + "@" + d[-1])
else:
    email_exist.append("does not exist")
final_frame["name"] = c 
final_frame["domain"] = d 
final_frame["rolebased"] = rolebased
final_frame["DNS check"] = email_exist
print(final_frame)
final_frame.to_excel(test_final +".xlsx")
print(time.time()-start)
print(defaulters)