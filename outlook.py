#!/usr/bin/python
# -*- coding: latin-1 -*-

import imaplib,email
import gspread ,json, time, pytz, re
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials


user = 'davspencer30@gmail.com'
password = 'Mastershifu30'
imap_url = 'imap.gmail.com'
subject = 'Quad Bike'

con = imaplib.IMAP4_SSL(imap_url)
con.login(user, password)
con.select('tests')       # select by label eg inbox,

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(creds)

wks = client.open('kgm_quad_calculator_201912 (1)').sheet1


result, items = con.search(None,'UNSEEN', 'FROM', 'patrickdeclan@btinternet.com')
# result, items = con.search(None,'(UNSEEN)', '(SUBJECT "%s")' % subjectt

def get_mails():
    global text_data
    item = items[0].split()
    if len(items) > 0:
        for emailid in items[-1]:   #use here for string slicing or selecting the latest emails or oldests smth.
            resp, data = con.fetch(emailid, '(RFC822)')
            raw_email = data[0][1]
            resp, data = con.store(emailid, '-FLAGS', 'Seen')
            email_body = email.message_from_string(raw_email)
            if email_body.is_multipart():
                for part in email_body.walk():
                    ctype = part.get_content_type()
                    cdispo = str(part.get('content-Disposition'))

                    if ctype == 'text/plain' and 'attachment' not in cdispo:
                        text_data = part.get_payload(decode=True)
                        print(text_data)
                        file = open('wks.txt', 'w')
                        file.write(text_data)
                        break
            else:
                body = email_body.get_payload(decode=True)
    else:
        print('no new mails mate...')

def calculate():
    global pcode
    global esize
    global _value
    global age
    global y_of_make
    global use
    global bonus_claim
    global ncb
    global mage
    global tp3
    global con_load
    global cload
    global own_yr
    global p_psg
    global stg4
    global ganchor
    global scrty
    global riders
    global val

    p1 = text_data.find('POSTCODE')
    p2 = text_data.find('DATE OF BIRTH')
    try:
        p3 = re.sub(r'>','', text_data[p1:p2].split('\n')[1])
        p4 = p3.split()
        pcode = p4[0]+ ' ' +p4[1]
        print pcode

    except:
        p3 = text_data[p1:p2].split('\n')[1].split()
        pcode = p3[0]+ ' ' +p3[1]
        print pcode

    e1 = text_data.find('ENGINE SIZE')
    e2 = text_data.find('REGISTRATION')
    esize = re.findall(r'\d+',text_data[e1:e2].split('\n')[1])[0]
    print esize

    v1 = text_data.find('VEHICLE VALUE')
    v2 = text_data.find('YEAR OF MANUFACTURE')
    _value = re.findall(r'\d+',text_data[v1:v2].split('\n')[1])[0]
    print _value

    ag1 = text_data.find('DATE OF BIRTH')
    ag2 = text_data.find('INSURANCE START DATE')
    ag3 = text_data[ag1:ag2].split('\n')[1].split('/')[-1]
    print ag3
    tz_London = pytz.timezone('Europe/London')
    datetime_London = datetime.now(tz_London).strftime("%Y")
    age = int(datetime_London) - int(ag3)
    print age

    y1 = text_data.find('YEAR OF MANUFACTURE')
    y2 = text_data.find('  IS THE VEHICLE MODIFIED?')
    y_of_make = re.findall(r'\d+',text_data[y1:y2].split('\n')[1])[0]
    print y_of_make

    use = 'SD&P Only'
    print use
    # SD&P Only       donno how to get this from the table data tho...
    # SD&P + C
    # business Use

    bonus_claim = '0%'
    print bonus_claim

    ncb = 'No'
    print ncb

    mage = '3000'
    print mage

    try:
        typ1 = text_data[text_data.find('VEHICLE USE'):text_data.find('HOW LONG HAVE YOU OWNED THE VEHICLE?')]
        tp2 = typ1.split('\n')[1]
        tp3 = re.sub(r'\s','',tp2)
        print tp3
        if tp3 == 'AgriculturalFarming' or '>AgriculturalFarming':
            tp3 = 'Sports'
            print tp3

        elif tp3 == 'OffRoadOnly' or '>OffRoadOnly':
            tp3 = 'Sports'
            print tp3
                                                      #utility,leisure,sports
        elif tp3 == 'RoadUseOnly' or '>RoadUseOnly':
            tp3 = 'Sports'
            print tp3


        elif tp3 == 'BothOnAndOffRoad' or '>BothOnAndOffRoad':
            tp3 = 'Sports'
            print tp3

        elif tp3 == '>Racing' or 'Racing':
            tp3 = 'Sports'
            print tp3

        elif tp3 == 'Other' or 'Other':
            tp3 = 'Sports'
            print tp3
    except Exception as err:
        print err


    con_load = '0%'
    print con_load

    cload = '0%'
    print cload

    ow1 = text_data.find('HOW LONG HAVE YOU OWNED THE VEHICLE?')
    ow2 = text_data.find('WHERE IS THE VEHICLE STORED')
    ow3 = text_data[ow1:ow2].split('\n')[1]
    own_yr = re.findall(r'\d+', ow3)[0]
    print(own_yr)

    p_psg = 'No'
    print p_psg

    stg1 = text_data.find('WHERE IS THE VEHICLE STORED OVERNIGHT?')
    stg2 = text_data.find('PERSONAL DETAILS')

    try:
        stg3 = re.sub(r'>','',text_data[stg1:stg2].split('\n')[1])
        stg4 = re.sub(r'\s','',stg3)
        print stg4
    except Exception as e:
        print e

    # stg = re.sub(r'\s','',text_data[stg1:stg2].split('\n')[1])
    # print stg

    ganchor = 'No'
    print ganchor

    scrty = 'Thatcham1'
    print scrty

    riders = 'insured only'
    print riders

    fn1 = text_data.find('NAME')
    fn2 = text_data.find('PREFERRED TELEPHONE NUMBER')
    try:
        fn3 =text_data[fn1:fn2].split('\n')[1].split()[1]
        titles = ['Mr', 'Mrs', 'Ms', 'Miss']
        for title in titles:
          if fn3 == title:
            x = 1
            first_name = text_data[fn1:fn2].split('\n')[1].split()[1+x]
            print first_name
        else:
            if fn3 not in titles:
                first_name = fn3
                print first_name
    except:
        first_name = text_data[fn1:fn2].split()[2]
        print first_name

    m1 = text_data[text_data.find('VEHICLE MAKE'):text_data.find('VEHICLE MODEL')]
    try:
        make = re.sub(r'>','',m1.split()[2] + m1.split()[3])
        print make

    except Exception as ex:
        n = 3
        if m1.split()[2] == m1.split()[n-1]:
            make = re.sub(r'>','',m1.split()[2])
            print make

    v1 = text_data[text_data.find('VEHICLE MODEL (if known)'):text_data.find('ENGINE SIZE')]
    try:
        _model = re.sub(r'>','',v1.split()[4] + v1.split()[5])
        print _model
    except Exception as e:
        n2 = 5
        if v1.split()[4] == v1.split()[n2 - 1]:
            _model = re.sub(r'>','',v1.split()[4])
            print _model

def parse_to_sheets():


    cover = wks.update_acell('E4','TPO')
    postcode = wks.update_acell('E6', pcode)
    engine_size = wks.update_acell('E8',esize)
    vehicle_value = wks.update_acell('E14', _value)
    age_of_rider = wks.update_acell('E16', age)
    year_of_manufacture = wks.update_acell('E18',y_of_make)
    vehicle_use = wks.update_acell('E20', use)
    claim_bonus = wks.update_acell('E22', bonus_claim)
    protected_NCB = wks.update_acell('E24', ncb)
    mileage = wks.update_acell('E26', mage)
    typ_of_quad = wks.update_acell('E28', tp3)
    conviction_load = wks.update_acell('E30',con_load)
    claim_load = wks.update_acell('E32',cload)
    ownership = wks.update_acell('E34', own_yr)
    pillion_passenger = wks.update_acell('E36',p_psg)
    storage = wks.update_acell('E38',stg4)
    ground_anchor = wks.update_acell('E40',ganchor)
    security = wks.update_acell('E42', scrty)
    additional_riders = wks.update_acell('E44',riders)

    val = wks.acell('H51').value
    print val

get_mails()
calculate()
parse_to_sheets()




