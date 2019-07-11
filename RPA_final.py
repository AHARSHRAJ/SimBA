import imaplib
import email
import pandas as pd

from flask import Flask
from flask import render_template, request, send_file
app = Flask(__name__)


server = "imap.gmail.com"
outputdir = '/home/thejas/RPA'


#function to search from a particular value
def search(key, value, con):
    result, data = con.search(None, key, '"{}"'.format(value))
    return data

# Connect to an IMAP server
def connect(server, user, password):
    m = imaplib.IMAP4_SSL(server)
    m.login(user, password)
    print("Logging in")
    m.select('INBOX')
    return m


# Download all attachment files for a given email
def downloadAttachments(m, emailid, outputdir):
    #m = connect(server, user, password)
    resp, data = m.fetch(emailid, "(BODY.PEEK[])")
    email_body = data[0][1]
    mail = email.message_from_bytes(email_body)
    if mail.get_content_maintype() != 'multipart':
        return
    for part in mail.walk():
        if part.get_content_maintype() != 'multipart' and part.get('Content-Disposition') is not None:
            open(outputdir + '/' + part.get_filename(), 'wb').write(part.get_payload(decode=True))


def download_from(server, user, password, date):
    
    m = connect(server, user, password)
    raw = search('SINCE', date, m)
    for emailid in raw[0].split():
         downloadAttachments(m, emailid, outputdir)    

    append()       
    



    


def append():
    

    df = pd.read_excel("list.xlsx")
    of = pd.read_excel("offers.xlsx")

    df['Offers'] = ""
    df['Zone'] = ""
    df['Month'] = ""
    temp_df = df

    for i in range (0, len(of)):

        for j in range (0, len(temp_df)):
            if of.BRAND[i].lower() == temp_df.Signature[j].lower():
                #update(i, j)
                temp_df.Offers[j] = of.OFFERS[i]
                temp_df.Zone[j] = of.ZONE[i]
                temp_df.Month[j] = of.MONTH[i]

        df = df.append(temp_df, ignore_index = True)
    
    df.to_excel("output.xlsx")

def map_server(server):
    if(server == 'Gmail'):
        server = 'imap.gmail.com'
    if(server == 'Outlook'):
        server = 'imap.outlook.com'
        
    return server

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/', methods = ['POST'])
def start():
    server = 'imap.gmail.com'
    #date = '4-Jul-2019'
    
    #form inputs
    user = request.form['user']
    password = request.form['password']
    #date input
    day = request.form['day']
    month = request.form['month']
    year = request.form['year']
    server = request.form['server']
    server = map_server(server)
    date = day+month+year
    download_from(server, user, password, date)
    
    return send_file("output.xlsx")

if __name__ == "__main__":
    app.run(debug = True)

