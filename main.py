import datetime
from pathlib import Path
import win32com.client
from win32com.client import Dispatch, constants
import time
import PySimpleGUI as sg
from docxtpl import DocxTemplate
import os
from datetime import date
from datetime import datetime
from docx2pdf import convert


sg.theme('Dark Black 1')


doc_ta = DocxTemplate('models/Transit Agreement - 2022.docx')
doc_co = DocxTemplate('models/CANCELLED OFFER.docx')
doc_ar = DocxTemplate('models/Arrival Room Check Form.docx')
doc_tw = DocxTemplate('models/TO WHOM IT MAY CONCERN.docx')
doc_bc = DocxTemplate('models/BOOKING CONFIRMATION.docx')
doc_pr = DocxTemplate('models/PAYMENT RECEIVED.docx')
doc_taxin= DocxTemplate('models/Tax Invoice.docx')
doc_pb = DocxTemplate('models/Provisional Booking.docx')
doc_cbc = DocxTemplate('models/Campbed Booking Confirmation.docx')


date = date.today()
today = date.strftime("%d %B %Y")
print(today)

font = ('Arial', 11)
sg.set_options(font=font)
menu_def = [['File', ['Arrival Room Check Form', 'Booking Confirmation','Cancelled Offer','Campbed Booking Confirmation','Payment Received', 'Provisional Booking','Tax Invoice', 'To Whom It May Concern','Transit Agreement']]]


layout1 = [
    [sg.Text('Transit Agreement', font=('Arial 14'))],     
    [sg.Text("Guest Name:"), sg.Input(key="GUEST", do_not_clear=False)],
    [sg.Text("Date of Arrival:"), sg.Input(key="D_ARRIVAL", do_not_clear=False)],
    [sg.Text("Date of Departure"), sg.Input(key="D_DEPARTURE",do_not_clear=False)],
    [sg.Text ("Room Number:"), sg.Input(key="ROOM", do_not_clear=False)],
    [sg.Text("Type of Room:"), sg.OptionMenu(key="T_ROOM",values=['Campbed','Single (Non- Ensuite)', 'Single (Ensuite)','Twin (Non- Ensuite)','Twin (Ensuite)','Triple (Non- Ensuite)','Triple (Ensuite)'],default_value='Twin (Ensuite)')],
    [sg.Text("Payment Information:"), sg.OptionMenu(key="PAY_INFO",values=['Airbnb','Booking.com','Business'],default_value='Booking.com')],
    [sg.Text ("Amount Paid:"), sg.Input(key="AMOUNT", do_not_clear=False)],
    [sg.Text ("Description:"), sg.Input(key="DESCRIPTION", do_not_clear=False)],
    [sg.Button("Create",key='CreateTA'), sg.Exit()],
   # [sg.CalendarButton('date')],
    
     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout2 = [   
    [sg.Text('Cancelled Offer', font=('Arial 14'))], 
    [sg.Text("First Name:"), sg.Input(key="FIRSTNAME", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="SURNAME", do_not_clear=False)],
    [sg.Text("Receptionist"), sg.Input(key="RECEPTIONIST",do_not_clear=False)],
    [sg.Text("Email"), sg.Input(key="EMAIL",do_not_clear=False)],
    [sg.Button("Create",key='CreateCO')],
    

     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout3 = [   
    [sg.Text('Arrival Room Check Form', font=('Arial 14'))], 
    [sg.Text("First Name:"), sg.Input(key="FIRSTNAMEAR", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="SURNAMEAR", do_not_clear=False)],
    [sg.Text ("Room Number:"), sg.Input(key="ROOMAR", do_not_clear=False)],
    [sg.Text("Date of Arrival:"), sg.Input(key="D_ARRIVALAR", do_not_clear=False)],
    [sg.Button("Create",key='CreateAR')],


     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout4=[
    [sg.Text('Payment Received', font=('Arial 14'))],
    [sg.Text("First Name:"), sg.Input(key="FIRSTNAMEPR", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="SURNAMEPR", do_not_clear=False)],
    [sg.Text("Email"), sg.Input(key="EMAILPR",do_not_clear=False)],
    [sg.Text ("Amount Paid:"), sg.Input(key="AMOUNTPR", do_not_clear=False)],
    [sg.Text("Reference:"), sg.OptionMenu(key="STAY",values=['1st instalment','2nd instalment','3rd instalment', 'stay','Summer Fees'],default_value='1st instalment')],
    [sg.Text("First Day:"), sg.Input(key="FDAY", do_not_clear=False)],
    [sg.Text("Last Day:"), sg.Input(key="LDAY", do_not_clear=False)],
    [sg.Text("Receptionist"), sg.Input(key="RECEPTIONISTPR",do_not_clear=False)],
    [sg.Button("Create",key='CreatePR')],

     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout5 = [
    [sg.Text('To Whom It May Concern', font=('Arial 14'))],
    [sg.Text ("Title:"), sg.OptionMenu(key="TITLETW",values=['Mr', 'Ms','Miss','Mrs'],default_value='Mr')],    
    [sg.Text("First Name:"), sg.Input(key="FNTW", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="SNTW", do_not_clear=False)],
    [sg.Button("Create",key='CreateTW')],


     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout6 = [ 
    [sg.Text('Booking Confirmation', font=('Arial 14'))],
    [sg.Text("First Name:"), sg.Input(key="FIRSTNAMEBC", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="SURNAMEBC", do_not_clear=False)],
    [sg.Text("Email"), sg.Input(key="EMAILBC",do_not_clear=False)],
    [sg.Text ("Amount Paid:"), sg.Input(key="AMOUNTBC", do_not_clear=False),],
    [sg.Text("Number of Guests:"), sg.Input(key="NGUESTBC", do_not_clear=False)],
    [sg.Text("Type of Room:"), sg.OptionMenu(key="TROOMBC",values=['Single (Non- Ensuite)', 'Single (Ensuite)','Twin (Non- Ensuite)','Twin (Ensuite)','Triple (Non- Ensuite)','Triple (Ensuite)'],default_value='Single (Non- Ensuite)')],
    [sg.Text("Date of Arrival:"), sg.Input(key="FDAYBC", do_not_clear=False)],
    [sg.Text("Date of Departure"), sg.Input(key="LDAYBC",do_not_clear=False)],
    [sg.Text("Receptionist"), sg.Input(key="RECEPTIONISTBC",do_not_clear=False)],
    [sg.Button("Create",key='CreateBC'), sg.Exit()],
    
     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout7 = [ 
    [sg.Text('Tax Invoice', font=('Arial 14'))],
    [sg.Text("First Name:"), sg.Input(key="TIFN", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="TISN", do_not_clear=False)],
    [sg.Text("Email:"), sg.Input(key="TIEMAIL",do_not_clear=False)],
    [sg.Text("Invoice Number:"), sg.Input(key="TIIN",do_not_clear=False)],
    [sg.Text ("Amount per Week:"), sg.Input(key="TIGW", do_not_clear=False)],
    [sg.Text("Type of Room:"), sg.OptionMenu(key="TIRT",values=['Campbed','Single (Non- Ensuite)', 'Single (Ensuite)','Twin (Non-Ensuite)','Twin (Ensuite)','Triple (Non-Ensuite)','Triple (Ensuite)'],default_value='Single (Non-Ensuite)')],
    [sg.Text("Date of Arrival:"), sg.Input(key="TIFD", do_not_clear=False)],
    [sg.Text("Date of Departure"), sg.Input(key="TILD",do_not_clear=False)],
    [sg.Text ("Receptionist:"), sg.Input(key="TIRECEPTIONIST", do_not_clear=False)],
    [sg.Button("Create",key='CreateTI'), sg.Exit()],
    
     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout8 = [ 
    [sg.Text('Provisional Booking', font=('Arial 14'))],
    [sg.Text("First Name:"), sg.Input(key="FIRSTNAMEPB", do_not_clear=False)],
    [sg.Text("Surname:"), sg.Input(key="SURNAMEPB", do_not_clear=False)],
    [sg.Text("Email"), sg.Input(key="EMAILPB",do_not_clear=False)],
    [sg.Text("Number of Guests:"), sg.Input(key="NGUESTPB", do_not_clear=False)],
    [sg.Text("Type of Room:"), sg.OptionMenu(key="TROOMPB",values=['Single (Non- Ensuite)', 'Single (Ensuite)','Twin (Non- Ensuite)','Twin (Ensuite)','Triple (Non- Ensuite)','Triple (Ensuite)'],default_value='Single (Non- Ensuite)')],
    [sg.Text ("Rate per Person per Night:"), sg.Input(key="AMOUNTPB", do_not_clear=False)],
    [sg.Text("Number of Nights:"), sg.Input(key="NNIGHTPB", do_not_clear=False)],
    [sg.Text("Date of Arrival:"), sg.Input(key="FDAYPB", do_not_clear=False)],
    [sg.Text("Date of Departure"), sg.Input(key="LDAYPB",do_not_clear=False)],
    [sg.Text("Receptionist"), sg.Input(key="RECEPTIONISTPB",do_not_clear=False)],
    [sg.Button("Create",key='CreatePB'), sg.Exit()],
    
     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]

layout9 = [ 
    [sg.Text('Campbed Booking Confirmation', font=('Arial 14'))],
    [sg.Text("Resident First Name:"), sg.Input(key="CBCFN", do_not_clear=False)],
    [sg.Text("Resident Surname:"), sg.Input(key="CBCSN", do_not_clear=False)],
    [sg.Text("Guest First Name:"), sg.Input(key="CBCGFN", do_not_clear=False)],
    [sg.Text("Guest Surname:"), sg.Input(key="CBCGSN", do_not_clear=False)],
    [sg.Text("Resident Email"), sg.Input(key="CBCEMAIL",do_not_clear=False)],
    [sg.Text("Campbed for Room:"), sg.Input(key="CBCCR",do_not_clear=False)],  
    [sg.Text ("Rate per Person per Night:"), sg.Input(key="CBCRATE", do_not_clear=False)],
    [sg.Text("Date of Arrival:"), sg.Input(key="CBCFDAY", do_not_clear=False)],
    [sg.Text("Date of Departure"), sg.Input(key="CBCLDAY",do_not_clear=False)],
    [sg.Text("Receptionist"), sg.Input(key="CBCRECEPTIONIST",do_not_clear=False)],
    [sg.Button("Create",key='CreateCBC'), sg.Exit()],
    
     [sg.Text("")],

    [sg.Text("Developed by: Max Fideles")]

]


layout = [
    [sg.Menu(menu_def, key='MENU')],
    [sg.Column(layout1, key='COL1'),
     sg.Column(layout2, key='COL2', visible=False),
     sg.Column(layout3, key='COL3',visible=False),
     sg.Column(layout4,key='COL4',visible=False),
     sg.Column(layout5,key='COL5',visible=False),
     sg.Column(layout6,key='COL6',visible=False),
     sg.Column(layout7,key='COL7',visible=False),
     sg.Column(layout8,key='COL8',visible=False),
     sg.Column(layout9,key='COL9',visible=False)]
]


window = sg.Window("Form Generator", layout, element_justification="right")


col, col1, col2, col3, col4, col5, col6, col7, col8, col9 = 1, window['COL1'], window['COL2'], window['COL3'],window['COL4'],window['COL5'], window['COL6'], window['COL7'], window['COL8'], window['COL9']

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
      break

    if event == 'Transit Agreement' and col != 1:
        col = 1
        col1.update(visible=True)
        col2.update(visible=False)
        col3.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=False)

    if event == 'Cancelled Offer' and col != 2:
        col = 2
        col2.update(visible=True)
        col1.update(visible=False)
        col3.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=False)

    if event == 'Arrival Room Check Form' and col != 3:
        col = 3
        col3.update(visible=True)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=False)

    if event == 'Payment Received' and col != 4:
        col = 4
        col3.update(visible=False)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=True)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=False)

    if event == 'To Whom It May Concern' and col != 5:
        col = 5
        col3.update(visible=False)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=True)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=False)

    if event == 'Booking Confirmation' and col != 6:
        col = 6
        col3.update(visible=False)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=True)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=False)
    
    if event == 'Tax Invoice' and col != 7:
        col = 7
        col3.update(visible=False)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=True)
        col8.update(visible=False)
        col9.update(visible=False)

    if event == 'Provisional Booking' and col != 8:
        col = 8
        col3.update(visible=False)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=True)
        col9.update(visible=False)

    if event == 'Campbed Booking Confirmation' and col != 9:
        col = 9
        col3.update(visible=False)
        col1.update(visible=False)
        col2.update(visible=False)
        col4.update(visible=False)
        col5.update(visible=False)
        col6.update(visible=False)
        col7.update(visible=False)
        col8.update(visible=False)
        col9.update(visible=True)
    
    if event == "CreateTA":
       
        doc_ta.render(values)
        output_path ="new\%s"% f"{values['GUEST']}-Transit Agreement.docx"
        doc_ta.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
   
    if event == "CreateCO":
       
        values["todayco"]=today
        doc_co.render(values)
        output_path = "new\%s"% f"{values['FIRSTNAME']}-CANCELLED OFFER.docx"
        doc_co.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")

        convert("new\%s"% f"{values['FIRSTNAME']}-CANCELLED OFFER.docx","new\%s"% f"{values['FIRSTNAME']}-CANCELLED OFFER.pdf")
         
        #Opening Outlook 
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = "Cancelled Offer"
        newMail.Body = "Hello " f"{values['FIRSTNAME']}"',\n \nWe are sending this email to inform you that your offer of accommodation has now been cancelled. \n \n Yours sincerely,\n \n 'f"{values['RECEPTIONIST']}"' \n The Hotel'
        os.startfile("new\%s"% f"{values['FIRSTNAME']}-CANCELLED OFFER.pdf", "print")
        newMail.BodyFormat = 2 
        newMail.To = f"{values['EMAIL']}"
        attachment1 = r"%s\new\%s" %(os.getcwd(),f"{values['FIRSTNAME']}-CANCELLED OFFER.pdf")
        print(attachment1)
        newMail.Attachments.Add(attachment1)
        newMail.display(True)

    if event == "CreatePR":
       
        values["todaypr"]=today
        doc_pr.render(values)
        output_path = "new\%s"% f"{values['FIRSTNAMEPR']}-Payment Received {values['STAY']}.docx"
        doc_pr.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        convert("new\%s"% f"{values['FIRSTNAMEPR']}-Payment Received {values['STAY']}.docx","new\%s"% f"{values['FIRSTNAMEPR']}-Payment Received {values['STAY']}.pdf")
        
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = f"{values['STAY']}-Payment Received"
        newMail.Body = "Hello " f"{values['FIRSTNAMEPR']}"',\n \nWe are sending this email to inform we have received your payment. \n \n Yours sincerely,\n \n 'f"{values['RECEPTIONISTPR']}"' \n The Hotel'
        newMail.BodyFormat = 2 
        newMail.To = f"{values['EMAILPR']}"
        attachment1 = r"%s\new\%s" %(os.getcwd(),f"{values['FIRSTNAMEPR']}-Payment Received {values['STAY']}.pdf")
        print(attachment1)
        newMail.Attachments.Add(attachment1)
        newMail.display(True)

    if event == "CreateAR":
       
        doc_ar.render(values)
        output_path ="new\%s"% f"{values['FIRSTNAMEAR']}-Arrival Room Check Form.docx"
        doc_ar.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
   
    if event == "CreateTW":
        if values["TITLETW"] == 'Mr':
            values["GEN"] = "his"
        else :
            values["GEN"] = "her"
        values["todaytw"]=today
        doc_tw.render(values)
        output_path ="new\%s"% f"{values['FNTW']}-To Whom It May Concern.docx"
        doc_tw.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")   

    if event == "CreateBC":


        if values["NGUESTBC"] == '1':
            values["NPBC"] = "person"
        else :
            values["NPBC"] = "people"
        values["todaybc"]=today
        
        doc_bc.render(values)
        output_path ="new\%s"% f"{values['FIRSTNAMEBC']}-Booking Confirmation.docx"
        doc_bc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        sg.popup_auto_close("Converting your file to pdf")
        convert("new\%s"% f"{values['FIRSTNAMEBC']}-Booking Confirmation.docx","new\%s"% f"{values['FIRSTNAMEBC']}-Booking Confirmation.pdf")

        #Opening Outlook
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = f"{values['FIRSTNAMEBC']}- Booking Confirmation"
        newMail.Body = "Hello " f"{values['FIRSTNAMEBC']}"',\n \nWe are sending attached on this email your Booking Confirmation letter . \n \n Yours sincerely,\n \n 'f"{values['RECEPTIONISTBC']}"' \n The Hotel'
        newMail.BodyFormat = 2
        newMail.To = f"{values['EMAILBC']}"
        attachment1 = r"%s\new\%s" %(os.getcwd(),f"{values['FIRSTNAMEBC']}-Booking Confirmation.pdf")
        newMail.Attachments.Add(attachment1)
        newMail.display(True)




    if event == "CreateTI":
       values["todayti"]=today
       
       TIFD = datetime.strptime(values["TIFD"], "%d %B %Y")
       TILD = datetime.strptime(values["TILD"], "%d %B %Y")
       tindays = (TILD - TIFD).days
       tinweeks = (tindays-1)/7
       print (tindays)
       print(tinweeks)

              
       values["TINETW"] =  float("{:0.2f}".format((float(values["TIGW"]))/1.2))    #calculating Net per week
       values["TITAXW"] =  float("{:0.2f}".format((float(values["TIGW"]))-(values["TINETW"]))) #calculating Tax per week

       values["NW"] =  tinweeks

       tigw = (float(values["TIGW"]))
       values["TIG"] = f'{(tinweeks*tigw):,.2f}'
       
       tinet = (tinweeks*tigw)/1.2
       values["TINET"] = f'{tinet:,.2f}'   #calculating Net Total

       titax = (tinweeks*tigw)-tinet
       values["TITAX"] = f'{titax:,.2f}' #calculating Tax Total

       values["NW"] =  f'{tinweeks:,.2f}' 

       doc_taxin.render(values)
       output_path ="new\%s"% f"{values['TIFN']}-Tax Invoice.docx"
       doc_taxin.save(output_path)
       sg.popup("File saved", f"File has been saved here: {output_path}")
       sg.popup_auto_close("Converting your file to pdf")
       convert("new\%s"% f"{values['TIFN']}-Tax Invoice.docx","new\%s"% f"{values['TIFN']}-Tax Invoice.pdf")

       const=win32com.client.constants
       olMailItem = 0x0
       obj = win32com.client.Dispatch("Outlook.Application")
       newMail = obj.CreateItem(olMailItem)
       newMail.Subject = f"{values['TIFN']}- Invoice"
       newMail.Body = "Hello " f"{values['TIFN']}"',\n \nWe are sending attached on this email your invoice. \n \n Yours sincerely,\n \n 'f"{values['TIRECEPTIONIST']}"' \n The Hotel'
       newMail.BodyFormat = 2
       newMail.To = f"{values['TIEMAIL']}"
       attachment1 = r"%s\new\%s" %(os.getcwd(),f"{values['TIFN']}-Tax Invoice.pdf")
       newMail.Attachments.Add(attachment1)
       newMail.display(True)

    if event == "CreatePB":


        if values["NGUESTPB"] == '1':
            values["NPPB"] = "person"
        else :
            values["NPPB"] = "people"
        

        nofguests = float(values["NGUESTPB"])  #Number of Guests
        nofnights = float(values["NNIGHTPB"]) #Number of Nights
        rpnight = float(values["AMOUNTPB"]) #Rate per Night

        values["TAMOUNTPB"] = f'{(nofnights*nofguests*rpnight):,.2f}' #Total Amount

        values["todaypb"]=today
        
        doc_pb.render(values)
        output_path ="new\%s"% f"{values['FIRSTNAMEPB']}-Provisional Booking.docx"
        doc_pb.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        sg.popup_auto_close("Converting your file to pdf")
        convert("new\%s"% f"{values['FIRSTNAMEPB']}-Provisional Booking.docx","new\%s"% f"{values['FIRSTNAMEPB']}-Provisional Booking.pdf")
        
        #Opening Outlook
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = f"{values['FIRSTNAMEPB']}- Provisional Booking"
        newMail.Body = "Hello " f"{values['FIRSTNAMEPB']}"',\n \nWe are sending attached on this email your Provisional Booking letter . \n \n Yours sincerely,\n \n 'f"{values['RECEPTIONISTPB']}"' \n The Hotel'
        newMail.BodyFormat = 2
        newMail.To = f"{values['EMAILPB']}"
        attachment1 = r"%s\new\%s" %(os.getcwd(),f"{values['FIRSTNAMEPB']}-Provisional Booking.pdf")
        newMail.Attachments.Add(attachment1)
        newMail.display(True)


    if event == "CreateCBC":

        values["cbctoday"]=today

       
        CBCFDAY = datetime.strptime(values["CBCFDAY"], "%d %B %Y")
        CBCLDAY = datetime.strptime(values["CBCLDAY"], "%d %B %Y")
        
        cbcndays = (CBCLDAY - CBCFDAY).days # diference of days
        
    
        cbctamount = cbcndays*float(values["CBCRATE"])
        print(cbctamount)

        values["CBCTAMOUNT"] = f'{(cbctamount):,.2f}'



        doc_cbc.render(values)
        output_path ="new\%s"% f"{values['CBCFN']}-Campbed Booking Confirmation.docx"
        doc_cbc.save(output_path)
        sg.popup("File saved", f"File has been saved here: {output_path}")
        sg.popup_auto_close("Converting your file to pdf")
        convert("new\%s"% f"{values['CBCFN']}-Campbed Booking Confirmation.docx","new\%s"% f"{values['CBCFN']}-Campbed Booking Confirmation.pdf")

        #Opening Outlook
        const=win32com.client.constants
        olMailItem = 0x0
        obj = win32com.client.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        newMail.Subject = f"{values['CBCFN']}-Campbed Booking Confirmation"
        newMail.Body = "Hello " f"{values['CBCFN']}"',\n \nWe are sending attached on this email your Campbed Booking Confirmation letter . \n \n Yours sincerely,\n \n 'f"{values['CBCRECEPTIONIST']}"' \n The Hotel'
        newMail.BodyFormat = 2
        newMail.To = f"{values['CBCEMAIL']}"
        attachment1 = r"%s\new\%s" %(os.getcwd(),f"{values['CBCFN']}-Campbed Booking Confirmation.pdf")
        newMail.Attachments.Add(attachment1)
        newMail.display(True)
           

window.close()