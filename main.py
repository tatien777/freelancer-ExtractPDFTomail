from PyPDF2 import PdfFileReader, PdfFileWriter,PdfWriter
from PyPDF2.generic import BooleanObject, NameObject, IndirectObject
from info import FROM_EMAIL,FROM_PWD,SMTP_SERVER,TO_EMAIL,SMTP_PORT
import info


import smtplib
import imaplib
import ssl
import email
from email.message import EmailMessage
import traceback
import re
import datetime
import json
def read_email_from_gmail():
    try:

        # Log in and send the email
            mail = imaplib.IMAP4_SSL("imap.gmail.com" )
            mail.login(FROM_EMAIL,FROM_PWD)
            mail.select('inbox')

            data = mail.search(None, '(SUBJECT "Test")' )
            mail_ids = data[1]
            id_list = mail_ids[0].split()   
            print(id_list)

            latest_email_id = int(id_list[-1])
            
            data = mail.fetch(str(latest_email_id), '(RFC822)' )
            for response_part in data:
                arr = response_part[0]
                if isinstance(arr, tuple):
                    msg = email.message_from_string(str(arr[1],'utf-8'))
                    
                    # with the content we can extract the info about
                    # who sent the message and its subject
                    mail_from = msg['from']
                    mail_subject = msg['subject']
                   

                    # then for the text we have a little more work to do
                    # because it can be in plain text or multipart
                    # if its not plain text we need to separate the message
                    # from its annexes to get the text
                    if msg.is_multipart():
                        mail_content = ''
                        # on multipart we have the text message and
                        # another things like annex, and html version
                        # of the message, in that case we loop through
                        # the email payload
                        for part in msg.get_payload():
                            # if the content type is text/plain
                            # we extract it
                            if part.get_content_type() == 'text/plain':
                                mail_content += part.get_payload()
                    else:
                        # if the message isn't multipart, just extract it
                        
                        mail_content = msg.get_payload()
                    # and then let's show its result
                    # print(f'From: {mail_from}')
                    # print(f'Subject: {mail_subject}')
                    # print(f'Content: {mail_content}')
                    return mail_content
    except Exception as e:
        traceback.print_exc() 
        print(str(e))



def write_email_from_gmail(subject,body,filename):
    try:
#    Add SSL (layer of security)
        context = ssl.create_default_context()
        em = EmailMessage()
        em['From'] = FROM_EMAIL
        em['To'] = TO_EMAIL
        em['Subject'] = subject
        em.set_content(body) 
        


        with open(filename, 'rb') as content_file:
            content = content_file.read()
            em.add_attachment(content, maintype='application/pdf', subtype='pdf', filename=filename)
        

        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as smtp:
                smtp.login(FROM_EMAIL, FROM_PWD)       
                smtp.sendmail(FROM_EMAIL, TO_EMAIL, em.as_string())
                print("message: SENT")

    except Exception as e:
        traceback.print_exc() 
        print(str(e))

def process_content(content):
    # reg1 = re.compile("\n.*") # ([A-Za-z0-9_])*
    reg2 = re.compile("\n(.*)\r")
    all_content = re.findall(reg2,content)
    good_content = [x for x in all_content if x]
    dict_content = {"Supaloc Job Number": "",
    "Address (i.e. number/street)":"",
    "Suburb":"",
    "Postcode":"",
    "Council":"",
    "Windspeed":"",
    "Type of construction":"",
    "Job Information (i.e. 1-3. include general layout)":"",
    "Steel Wall Framing (i.e. 7, 12-17. include beam layout)":"",
    "Steel roof Framing (i.e. 4-6, 8-11)":"",
    "Date of Supaloc Inspection":"",
    }
    # print(a[a.index("Supaloc Job Numbe")+1])
    # print(good_content)
    for key in dict_content: 
        # print(key, "And values: ",good_content[good_content.index(key)+1])
        dict_content[key] = good_content[good_content.index(key)+1]

    dict_content[("Date of Supaloc Inspection")] = datetime.datetime.strptime(dict_content.get("Date of Supaloc Inspection"), "%d/%m/%y").strftime('%m/%d/%Y')
    
    # print(dict_content)    
    return dict_content
    



def set_need_appearances_writer(writer):
    # See 12.7.2 and 7.7.2 for more information:
    # http://www.adobe.com/content/dam/acom/en/devnet/acrobat/
    #     pdfs/PDF32000_2008.pdf
    try:
        catalog = writer._root_object
        # get the AcroForm tree and add "/NeedAppearances attribute
        if "/AcroForm" not in catalog:
            writer._root_object.update(
                {
                    NameObject("/AcroForm"): IndirectObject(
                        len(writer._objects), 0, writer
                    )
                }
            )

        need_appearances = NameObject("/NeedAppearances")
        writer._root_object["/AcroForm"][need_appearances] = BooleanObject(True)
        return writer

    except Exception as e:
        print("set_need_appearances_writer() catch : ", repr(e))
        return writer

def printPDF(dict_map,open_file):
# Open the pdf file 
    PDF_read = PdfFileReader(open_file)
    # page = PDF_read.pages[0]

    if "/AcroForm" in PDF_read.trailer["/Root"]:
        PDF_read.trailer["/Root"]["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )

    writer = PdfFileWriter()
    # print(page)

    # first_page = PDF_read.getPage(0)
    # print(first_page.extractText())\

    set_need_appearances_writer(writer)
    if "/AcroForm" in writer._root_object:
        writer._root_object["/AcroForm"].update(
            {NameObject("/NeedAppearances"): BooleanObject(True)}
        )
    ######## Page 1 #############
    field_dictionary = { 
        'Street address 1': dict_map.get('Street address 1'),
        'Street address 2' : dict_map.get('Street address 2'),
        'Postcode 1' : dict_map.get('Postcode 1'),
        'Lot and plan details 1': dict_map.get('Lot and plan details 1') ,
        'Local government area 1': dict_map.get('Local government area 1') ,
        'Building/structure description': dict_map.get('Building/structure description') ,
        # 'Reference documentation': dict_map.get('Reference documentation'),
    }

    # print(PDF_read.getFormTextFields() )

    writer.addPage(PDF_read.getPage(0))
    writer.updatePageFormFieldValues(writer.getPage(0), field_dictionary)


    ##### Page 2 #######
    field_dictionary = { 
        'Basis of certification': dict_map.get('Basis of certification'),
        'Reference documentation': dict_map.get('Reference documentation'),
        'Date 9':  dict_map.get('Date 9'),
    }

    # print(PDF_read.getFormTextFields() )

    writer.addPage(PDF_read.getPage(1))
    writer.updatePageFormFieldValues(writer.getPage(1), field_dictionary)
    writer.addPage(PDF_read.getPage(2))
    writer.addPage(PDF_read.getPage(3))

    with open("all.pdf", "wb") as fp:
        writer.write(fp)


## TEST ### 
def main_task():
    mail_content = read_email_from_gmail()
    dict_content = process_content(mail_content)
    print(dict_content)
    subject = "Test program"
    body = ( f"""
        Content log: 
        {json.dumps(dict_content)}
        """) 
    
    # data_map = ['71158','Lot 20, 21 Ruchi Place','Wynnum West','4178','BRISBANE CITY COUNCIL'
    # ,'33m/s (N2)','Single Storey Dwelling','1-2','11-18','3-10','07/07/2022']

    dict_map = {
        'Street address 1' :dict_content.get("Address (i.e. number/street)"),
        'Street address 2' : dict_content.get("Suburb"),
        'Postcode 1' :dict_content.get("Postcode"),
        'Lot and plan details 1': dict_content.get("Address (i.e. number/street)") ,
        'Local government area 1': dict_content.get("Council") ,
        'Basis of certification': f"Certification based on Steel Building Systems (SBS) International test results, standard section data and check calculation and compliance with the requirements of the Building Code of Australia.\r\rDesign Wind Rating: Vhp = {dict_content.get('Type of construction')}" ,
        'Reference documentation': f"""Steel Building Systems job number: {dict_content.get('Supaloc Job Number')} \rDrawings: FD {dict_content.get('Job Information (i.e. 1-3. include general layout)')} Job Information, 
        FD {dict_content.get('Steel Wall Framing (i.e. 7, 12-17. include beam layout)')} Steel Wall Framing, 
        FD {dict_content.get('Steel roof Framing (i.e. 4-6, 8-11)')}\r Steel Roof Framing \rSBSI standard detail sheets: SBS 2004 WF01-WF013 & WF13-WF17\r                                                   SBS 2004 RF01-RF18, SBS 2004 BJ01-BJ04\r                                                   SBS 2004 PB pages 1 to 4, SBS 2004 SF01-SF05\r\r""" ,
        'Date 9':  dict_content.get('Date of Supaloc Inspection')
    }
    openfile = 'test.pdf'
    outputfile = 'all.pdf'
    printPDF(dict_map,openfile)

    print(body)
    write_email_from_gmail(subject,body,outputfile)

main_task()
        