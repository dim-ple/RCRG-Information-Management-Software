import tkinter as tk
from tkinter import font as tkfont
from tkinter import StringVar, BooleanVar, Label, Entry, OptionMenu, Radiobutton, Button, Toplevel

import shutil

import os
import win32com.client as win32

from Contact_Dictionaries import rcrg, lender, attorney

from DateAndTime import Time
from DateAndTime import datetime

from fillpdf import fillpdfs 

import sqlite3

# **Note for Later** Possibly nest within MainFrame class
try:
    conn = sqlite3.connect('rcrg.db')
    c = conn.cursor()
    print("Successfully Connected to Database!")

except:
    pass

agents = []
lenders = []
attorneys =[]

for agent in rcrg:
    agents.append(agent)

for lend in lender:
    lenders.append(lend)

for office_name in attorney:
    attorneys.append(office_name)

commissions = [
    "6% Total, 3/3",
    "5.5% Total, 2.75/2.75",
    "5% Total, 2.5/2.5",
    "5% Total, 2/3",
    "Other"
    ]

admin_fees = [
    395.0,
    495.0,
    0.0
]

teams = [
    "Alpha",
    "Bravo"
    ]

loanType = [
    "Conventional",
    "FHA",
    "Cash",
    "VHDA",
    "USDA",
]


class MainFrame(tk.Tk):

    def __init__(self, *args, **kwargs):
        tk.Tk.__init__(self, *args, **kwargs)
        self.titlefont = tkfont.Font(family = 'Verdana', size = 12,
                                     weight = "bold", slant = 'roman')
        
        container = tk.Frame()
        container.grid(row=0, column=0, sticky='nesw')

        self.geometry('1000x800')
        self.id = tk.StringVar()
        self.id.set("RCRG Admin")

        self.listing = {}
        
        for p in (WelcomePage, BuyerTran, SellerTran, TeamMeeting, ZillowTeam, BuyerZillow, SellerZillow, NewListing):
            page_name = p.__name__
            frame = p(parent = container, controller = self)
            frame.grid(row=0, column=0, sticky='nsew')
            self.listing[page_name] = frame
        
        self.up_frame('WelcomePage')


    def up_frame(self, page_name):
        page = self.listing[page_name]
        page.tkraise()


class WelcomePage(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        
        label = tk.Label(self, text = 'Welcome Page \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "New Buyer Transaction", 
                        command = lambda: controller.up_frame("BuyerTran"))
        bou1.grid(column=2, row=1)

        bou2 = tk.Button(self, text = "New Seller Transaction",
                         command = lambda: controller.up_frame("SellerTran"))
        bou2.grid(column=2, row=2)

        bou3 = tk.Button(self, text = "Team Meeting E-mail",
                         command = lambda: controller.up_frame("TeamMeeting"))
        bou3.grid(column=2, row=3)

        bou4 = tk.Button(self, text = "Zillow Team E-mail",
                         command = lambda: controller.up_frame("ZillowTeam"))
        bou4.grid(column=2, row=4)

        bou5 = tk.Button(self, text = "Buyer Zillow Review",
                         command = lambda: controller.up_frame("BuyerZillow"))
        bou5.grid(column=2, row=5)

        bou6 = tk.Button(self, text = "Seller Zillow Review",
                         command = lambda: controller.up_frame("SellerZillow"))
        bou6.grid(column=2, row=6)

        bou7 = tk.Button(self, text = "New Listing Folder",
                         command = lambda: controller.up_frame("NewListing"))
        bou7.grid(column=2, row=7)

        bou8 = tk.Button(self, text = "Close the Window",
                         command= controller.destroy)
        bou8.grid(column=2, row=8)


class BuyerTran(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        label = tk.Label(self, text = 'New Buyer Transaction \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "Back to Main", 
                        command = lambda: controller.up_frame("WelcomePage"))
        bou1.grid(column=1, row=1)

        clicked_agents = StringVar()
        clicked_agents.set("Agents")

        clicked_lenders = StringVar()
        clicked_lenders.set("Lenders")

        clicked_boolean = BooleanVar()

        clicked_admin_fee = StringVar()

        clicked_attorneys = StringVar()
        clicked_attorneys.set("Attorneys")

        #1st Q & A - Property Address
        lbl1 = Label(self, text = "What is the Property Address?")
        lbl1.grid(column = 2, row = 0)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 0)

        #City
        lbl1 = Label(self, text = "What is the Property City?")
        lbl1.grid(column = 2, row = 1)
        txt1 = Entry(self, width=20)
        txt1.grid(column = 3, row = 1)

        #Zip
        lbl1 = Label(self, text = "What is the Property Zip?")
        lbl1.grid(column = 2, row = 2)
        txt1 = Entry(self, width=8)
        txt1.grid(column = 3, row = 2)

        #County
        lbl1 = Label(self, text = "What is the Property County?")
        lbl1.grid(column = 2, row = 3)
        txt1 = Entry(self, width=20)
        txt1.grid(column = 3, row = 3)

        #MLS Number
        lbl1 = Label(self, text = "What is the MLS Number?")
        lbl1.grid(column = 2, row = 4)
        txt1 = Entry(self, width=10)
        txt1.grid(column = 3, row = 4)

        #Sales Price
        lbl1 = Label(self, text = "What is the Sales Price?")
        lbl1.grid(column = 2, row = 5)
        txt1 = Entry(self, width=20)
        txt1.grid(column = 3, row = 5)

        #List Price
        lbl1 = Label(self, text = "What was the List Price?")
        lbl1.grid(column = 2, row = 6)
        txt1 = Entry(self, width=20)
        txt1.grid(column = 3, row = 6)

        '''#Offer Date
        lbl1 = Label(self, text = "What was the Offer Date?")
        lbl1.grid(column = 2, row = 7)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 7)

        #Ratification Date
        lbl1 = Label(self, text = "What was the Date of Ratification?")
        lbl1.grid(column = 2, row = 8)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 8)

        #Close Date
        lbl1 = Label(self, text = "What is the Closing Date?")
        lbl1.grid(column = 2, row = 9)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 9)

        #Seller Paid Closing Costs
        lbl1 = Label(self, text = "What is the Property City?")
        lbl1.grid(column = 2, row = 10)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 10)

        #Seller Name
        lbl4 = Label(self, text = "What is the Seller(s) Full Name? For Multiple Names, separate with a ';'")
        lbl4.grid(column = 2, row = 11)
        txt3 = Entry(self, width=38)
        txt3.grid(column = 3, row = 11)'''

        #2nd Q & A - Agent
        lbl2 = Label(self, text = "Who is the Selling Agent?")
        lbl2.grid(column = 2, row = 7)
        drop1 = OptionMenu(self, clicked_agents, *agents)
        drop1.grid(column = 3, row = 7)

        #3rd Q & A - Commission
        lbl3 = Label(self, text = "What is the Selling Agent's Commission")
        lbl3.grid(column = 2, row = 8)
        txt2 = Entry(self, width=8)
        txt2.grid(column = 3, row = 8)

        #Transaction Fee (Radio 3 option - 395, 495, 0)
        lbl4 = Label(self, text="What is the Admin Fee?")
        lbl4.grid(column=2, row=9)
        radio3 = Radiobutton(self, text="N/A", variable = clicked_admin_fee,
                            value="0")
        radio3.grid(column=3, row=9)
        radio4 = Radiobutton(self, text="$495", variable = clicked_admin_fee,
                            value="495")
        radio4.grid(column=4, row=9)
        radio5 = Radiobutton(self, text="$395", variable = clicked_admin_fee,
                             value="395")
        radio5.grid(column=5, row=9)
        clicked_admin_fee.set("395")

        #5th Q & A - Client Name
        lbl5 = Label(self, text = "What is the Client's Full Name? For Multiple Names, separate with a ';'")
        lbl5.grid(column=2, row=10)
        txt3 = Entry(self, width=38)
        txt3.grid(column=3, row=10)

        #Client Phone Number(s)

        #6th Q & A - Lender
        lbl6 = Label(self, text = "Who is the Lender?")
        lbl6.grid(column = 2, row = 11)
        drop2 = OptionMenu(self, clicked_lenders, *lenders)
        drop2.grid(column = 3, row = 11)

        #7th Q & A - EMD
        lbl7 = Label(self, text="Do we have the EMD?")
        lbl7.grid(column = 2, row = 12)
        radio1 = Radiobutton(self, text = "Yes", variable = clicked_boolean,
                            value=True)
        radio1.grid(column = 3, row = 12)
        radio2 = Radiobutton(self, text = "No", variable = clicked_boolean,
                            value=False)
        radio2.grid(column = 4, row = 12)

        #8th Q & A - Attorney Contact
        lbl8 = Label(self, text = "Who is the Attorney?")
        lbl8.grid(column = 2, row = 13)
        drop3 = OptionMenu(self, clicked_attorneys, *attorneys)
        drop3.grid(column = 3, row = 13)

        #9th Q & A - Client E-mail
        lbl9 = Label(self, text = "What is the Client's E-mail?")
        lbl9.grid(column = 2, row = 14)
        txt4 = Entry(self, width=38)
        txt4.grid(column = 3, row = 14)

        #Listing Agent Company
        
        #10th Q & A - Listing Agent Name
        lbl10 = Label(self, text = "What is the Listing Agent's Name")
        lbl10.grid(column = 2, row = 15)
        txt5 = Entry(self, width=38)
        txt5.grid(column = 3, row = 15)

        #11th Q & A - Listing Agent E-mail
        lbl11 = Label(self, text = "What is the Listing Agent's E-mail")
        lbl11.grid(column = 2, row = 16)
        txt6 = Entry(self, width=38)
        txt6.grid(column = 3, row = 16)


        def buyer_folder():
            if os.getcwd() != 'C:\\Users\\rcrgr\\Desktop\\E-mail Programs':
                os.chdir('C:\\Users\\rcrgr\\Desktop\\E-mail Programs')
        
            property_address = txt1.get()
            selling_agent = clicked_agents.get()
            listing_agent = txt5.get()
            listing_email = txt6.get()
            commission = txt2.get()
            client1 = txt3.get()
            client2 = ' '
            client_email1 = txt4.get()
            client_email2 = ' '
            attorney_contact = clicked_attorneys.get()
            lender_contact = clicked_lenders.get()
            admin_fee = clicked_admin_fee.get()

            

            if ";" in client1:
                i = client1.find(";")
                client2 = client1[(i+2):]
                client1 = client1[0:i]

            if ";" in client_email1:
                i = client_email1.find(";")
                client_email2 = client_email1[(i+2):]
                client_email1 = client_email1[0:i]

            fillpdfs.get_form_fields("Transaction Info Sheet(Fillable).pdf")


            data_dict = {'Property Address': property_address, 'City': '', 'State': 'VA', 'Zip': '', 'County': '',
                        'CVRMLS': '', 'Sales Price': '', 'Offer Date_af_date': '', 'Date2_af_date': '',
                        'Rat-Date_af_date': '', 'Closing Date_af_date': '', 'List Price': '', 'Closing Costs Paid by Seller': '',
                        'Seller': '', 'Purchaser': 'Yes', 'Seller 1': '', 'Seller 2': '', 'Seller Email 1': '', 'Seller Email 2': '',
                        'Seller Cell': '', 'Seller Work': '', 'Seller Home': '', 'Seller Fax': '', 'Seller Forwarding Address': '',
                        'Seller City': '', 'Seller State': '', 'Seller Zip': '', 'Buyer 1': client1, 'Buyer 2': client2,
                        'Buyer Email': client_email1, 'Buyer Email 2': client_email2, 'Buyer Cell': '', 'Buyer Work': '', 'Buyer Home': '',
                        'Buyer Fax': '', 'Home Warranty': '', 'Home Inspec\x98on Co': '', 'Termite Co': '', 'FuelOil Co': '',
                        'Well  Sep\x98c Co': '', 'Lender': lender_contact, 'Loan Officer Name': lender[lender_contact][2], 'Loan Officer Phone': lender[lender_contact][3], 'Loan Officer Email': lender[lender_contact][0],
                        'Seller Attorney Firm': '', 'Seller Attorney Contact': '', 'Seller Office Phone': '', 'Seller Attorney Fax': '',
                        'Seller Attorney Email': '', 'Buyer Attorney Firm': attorney[attorney_contact][2], 'Buyer Attorney Contact': attorney[attorney_contact][3], 'Buyer Attorney Office Phone': '',
                        'Buyer Attorney Fax': '', 'Buyer Attorney Email': attorney[attorney_contact][0], 'HOA Name': '', 'HOA Mgmt Co': '', 'HOA Phone': '', 'HOA Email': '',
                        'Listing Company Name': '', 'Listing Agent Name': listing_agent, 'Transaction Coordinator': '', 'Listing Agent Phone': '',
                        'Listing Agent E-mail': listing_email, 'Selling Company Name': 'The Rick Cox Realty Group', 'Selling Agent Name': rcrg[selling_agent][1], 'Selling Agent TC': 'Harrison Goehring - harrison@rickcoxrealty.com',
                        'Selling Agent Phone': rcrg[selling_agent][4], 'Selling Agent Email': rcrg[selling_agent][0], 'Escrow Deposit': '', 'Held by': '', 'Commission': commission + ' to Selling Agent',
                        'Transac\x98on Fee': admin_fee, 'Referral Fee': '', 'Paid to': '', 'Referral Address': '', 'Reset': ''}
            
            fillpdfs.write_fillable_pdf('Transaction Info Sheet(Fillable).pdf', 'Transaction Info Sheet(f).pdf', data_dict)
            
            if selling_agent == "Other":
                path = " "
            else:
                path = rcrg[selling_agent][3]
                os.chdir(path)

            os.mkdir(property_address)

            os.chdir(f"{path}\\{property_address}")

            os.mkdir("Contract-Addenda")
            os.mkdir("Invoices-Inspections")

            shutil.copy('C:\\Users\\rcrgr\\Desktop\\E-mail Programs\\Transaction Info Sheet(f).pdf', f'{path}\\{property_address}\\Contract-Addenda')

        def buyer_email():
            property_address = txt1.get()
            selling_agent = clicked_agents.get()
            client_name1 = txt3.get()
            client_email = txt4.get()
            Address_To_Client = f"{client_name1}"

            if ";" in client_name1:
                i = client_name1.find(";")
                client_name2 = client_name1[(i+2):]
                client_name1 = client_name1[0:i]
                Address_To_Client = f"{client_name1} & {client_name2}"

            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Your New Purchase of ' + property_address
            mailItem.BodyFormat = 1

            if selling_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[selling_agent][1]
                mailItem.CC = rcrg[selling_agent][0] + " amy@rickcoxrealty.com;"

            html_body = f"""
                <p class=MsoNormal>Good {Time}, {Address_To_Client}!<br><br></p>
                <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I work with {agent_name} and will be assisting with your purchase of {property_address}. Attached, you will find copies of the fully-executed contract and any addenda or disclosures in conjunction with your closing.<br><br></p>
                <p class=MsoNormal>Should you have any questions regarding closing or any aspect of the transaction leading up to that point, please feel free to reach out me. My congratulations to you on your upcoming home purchase!<br><br></p>
                <p class=MsoNormal>CC: Your agent, {agent_name}; Team Administrator, Amy Foldes; <br><br></p>
                <p class=MsoNormal>Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
            
            mailItem.HTMLBody = html_body
            mailItem.To = client_email
            

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        def attorney_email():
            property_address = txt1.get()
            selling_agent = clicked_agents.get()
            commission = txt2.get()
            attorney_contact = clicked_attorneys.get()

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'New Purchase-Side Transaction - ' + property_address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Attorney E-mail'

            #To: Operating Logic - Dictionary Call
            if attorney_contact == "Other":
                Attorney_Name = " "
                mailItem.To = " "
            else:
                Attorney_Name = attorney[attorney_contact][1]
                mailItem.To = attorney[attorney_contact][0]

            #CC: Operating Logic - Dictionary Call
            if selling_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[selling_agent][1]
                mailItem.CC = rcrg[selling_agent][0] + " amy@rickcoxrealty.com;"
                

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Attorney_Name}!<br><br></p>
                <p class=MsoNormal>{agent_name}'s client would like to use your office for the title and settlement work needed for their purchase of {property_address}. Please find the ratified contract, transaction information sheet and tax record attached!<br><br></p>
                <p class=MsoNormal> Please note that the selling agent's commission for this transaction will be {commission}. Additionally, our brokerage will charge a $395.00 Administrative Fee to the purchaser at closing. Please overnight both checks to our office at <b> 2913 Fox Chase Lane, Midlothian, VA 23112. </b> Thank you! <br><br></p>
                <p class=MsoNormal>CC: {agent_name}, Selling Agent; Team Administrator, Amy Foldes;<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        def lender_email():
            lender_contact = clicked_lenders.get()
            EMD_Status = clicked_boolean.get()
            property_address = txt1.get()
            selling_agent = clicked_agents.get()
            client_name1 = txt3.get()
            client_name2 = ' '
            client_email_Message = f"client, {client_name1}"
            Client_Subject_Line = f"{client_name1}"
    
            if ";" in client_name1:
                i = client_name1.find(";")
                client_name2 = client_name1[(i+2):]
                client_name1 = client_name1[0:i]
                client_email_Message = f"clients, {client_name1} & {client_name2}"
                Client_Subject_Line = f"{client_name1} & {client_name2}"
                
            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = f'New Purchase Contract - {property_address} for ({Client_Subject_Line})'
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Lender E-mail'

            #Addressee Operating Logic - Database
            if lender_contact == "Other":
                Lender_Name = " "
                mailItem.To = " "
            else:
                Lender_Name = lender[lender_contact][1]
                mailItem.To = lender[lender_contact][0]
                
            if selling_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[selling_agent][1]
                mailItem.CC = rcrg[selling_agent][0] + " amy@rickcoxrealty.com;"

            #EMD Logic
            if EMD_Status == True:
                EMD = "We have received the earnest money deposit, please find a copy of the check attached."
            elif EMD_Status == False:
                EMD = "We have not yet received the earnest money deposit. Once received, we will forward a copy of the check to you!"
            else:
                EMD = ""
                
            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Lender_Name}!<br><br></p>
                <p class=MsoNormal>Please find a ratified contract attached for {agent_name}'s {client_email_Message}! {EMD}<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        def listing_agent_email():
            property_address = txt1.get()
            selling_agent = clicked_agents.get()
            attorney_contact = clicked_attorneys.get()
            listing_agent = txt5.get()
            listing_email = txt6.get()
            
            
            if clicked_attorneys.get() == "Other":
                attorney_msg = f"Our purchaser will be using {attorney[attorney_contact][2]} for their title and settlement needs. The primary contact will be {attorney[attorney_contact][3]}, their e-mail is {attorney[attorney_contact][0]}."
            else:
                attorney_msg = "Our purchaser has not yet decided on who they will be using for their title and settlement needs. Once they have decided, I will let you know!"

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Coordinator Introduction - ' + property_address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Coordinator Introduction'

            #To: Operating Logic - Dictionary Call
            if listing_email == "":
                mailItem.To = " "
            else:
                mailItem.To = listing_email

            #CC: Operating Logic - Dictionary Call
            if selling_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[selling_agent][1]
                mailItem.CC = rcrg[selling_agent][0] + " amy@rickcoxrealty.com;"
                

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {listing_agent}!<br><br></p>
                <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I will be assisting {agent_name} and their client on the purchase of {property_address}. I look forward to working with you!<br><br></p>
                <p class=MsoNormal>{attorney_msg} Would you mind providing me with the contact for the Seller's Attorney or Title Company who will be handling the deed preparation for the Seller once that information becomes available?<br><br></p>
                <p class=MsoNormal>Additionally, would your seller be willing to share who their current utility providers for Electricity, Water/Sewer, Internet, Trash and Gas are?<br><br></p>
                <p class=MsoNormal>CC: {agent_name}, Selling Agent; Team Administrator, Amy Foldes;<br><br></p>
                <p class=MsoNormal>Kind regards & thanks,<br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()
        
        def clear_fields():
            clicked_agents.set("Agents")
            clicked_lenders.set("Lenders")
            clicked_attorneys.set("Attorneys")
            txt1.delete("0", "end")
            txt2.delete("0", "end")
            txt3.delete("0", "end")
            txt4.delete("0", "end")
            txt5.delete("0", "end")
            txt6.delete("0", "end")
            clicked_boolean.set(False)
            clicked_admin_fee.set("395")

        def data_submit(table_name, first, last, cell, email, agent_type, dpor, broker):
            
            c.execute(f"""
            
                INSERT INTO {table_name} 
                (agentfirst, agentlast, agentphone, agentemail, agenttype, agentlicensenum, agentbroker) 
            
                VALUES 
                ({first}, {last}, {cell}, 
                {email}, {agent_type}, {dpor}, 
                {broker})
            
                """)
            
            
        def new_agent_info():
            
            agent_table = "agents"
            clicked_agent_type = StringVar()


            top = Toplevel(parent)
            top.geometry("450x175")
            top.title("New Agent Info - Input Form")

            agent_first_lbl = Label(top, text = "Agent First Name:")
            agent_first_lbl.grid(column = 2, row = 0)
            agent_first_ent = Entry(top, width=20)
            agent_first_ent.grid(column = 3, row = 0)

            agent_last_lbl = Label(top, text = "Agent Last Name:")
            agent_last_lbl.grid(column = 2, row = 1)
            agent_last_ent = Entry(top, width=20)
            agent_last_ent.grid(column = 3, row = 1)

            agent_cell_lbl = Label(top, text = "Agent Cell:")
            agent_cell_lbl.grid(column = 2, row = 2)
            agent_cell_ent = Entry(top, width=20)
            agent_cell_ent.grid(column = 3, row = 2)
           
            agent_email_lbl = Label(top, text = "Agent E-mail:")
            agent_email_lbl.grid(column = 2, row = 3)
            agent_email_ent = Entry(top, width=38)
            agent_email_ent.grid(column = 3, row = 3)

            # Add Agent Type field (Dropdown selection, default to 'Salesperson')
            agent_type_lbl = Label(top, text = "Agent Type:")
            agent_type_lbl.grid(column = 2, row = 4)
            agent_type_select = OptionMenu(top, clicked_agent_type, "Salesperson", "Principal Broker")
            agent_type_select.grid(column = 3, row = 4)

            agent_dpor_lbl = Label(top, text = "Agent DPOR License Number:")
            agent_dpor_lbl.grid(column = 2, row = 5)
            agent_dpor_ent = Entry(top, width=17)
            agent_dpor_ent.grid(column = 3, row = 5)

            agent_broker_lbl = Label(top, text = "Agent Brokerage Name:")
            agent_broker_lbl.grid(column = 2, row = 6)
            agent_broker_ent = Entry(top, width=30)
            agent_broker_ent.grid(column = 3, row = 6)

            #Need to See if we can use lambda to close the pop-up window when the data is successfully passed. This would allow us to remove the close button.
            pass_data_button = Button(top, text = "Submit Data",
                                      command = lambda:[data_submit(agent_table, first=agent_first_ent.get(), last=agent_last_ent.get(), cell=agent_cell_ent.get(), 
                                                        email=agent_email_ent.get(), agent_type=agent_type_select.get(), dpor=agent_dpor_ent.get(), broker=agent_broker_ent.get())])
            pass_data_button.grid(column=3, row=6)

            close_button = Button(top, text = "Close the Window",
                              command= top.destroy)
            close_button.grid(column=3, row=7)
            




        #Execute Button
        submit_button = Button(self, text = "Submit",
                               command = lambda:[buyer_email(), attorney_email(), listing_agent_email(), lender_email()])
        submit_button.grid(column = 3, row = 17)

        new_folder_button = Button(self, text = "Create New Folder",
                                   command = lambda:[buyer_folder()])
        new_folder_button.grid(column=3, row=18)

        clear_fields_button = Button(self, text = "Reset Fields",
                                     command = lambda:[clear_fields()])
        clear_fields_button.grid(column=3, row=19)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=3, row=20)

        new_agent_button = Button(self, text="New Agent",
                                  command = lambda: new_agent_info())
        new_agent_button.grid(column=4, row=15)


class SellerTran(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        label = tk.Label(self, text = 'New Seller Transaction \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "Back to Main", 
                        command = lambda: controller.up_frame("WelcomePage"))
        bou1.grid(column=1, row=1)
        
        clicked_agents = StringVar()
        clicked_agents.set("Agents")

        clicked_attorneys = StringVar()
        clicked_attorneys.set("Attorneys")

        clicked_commissions = StringVar()
        clicked_commissions.set("Splits")

        #1st Q & A - Property Address
        lbl1 = Label(self, text = "What is the Property Address?")
        lbl1.grid(column = 2, row = 0)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 0)

        #2nd Q & A - Agent
        lbl2 = Label(self, text = "Who is the Listing Agent?")
        lbl2.grid(column = 2, row = 1)
        drop1 = OptionMenu(self, clicked_agents, *agents)
        drop1.grid(column = 3, row = 1)

        #3rd Q & A - Commission
        lbl3 = Label(self, text = "What is the Total Commission & Split?")
        lbl3.grid(column = 2, row = 2)
        drop2 = OptionMenu(self, clicked_commissions, *commissions)
        drop2.grid(column = 3, row = 2)

        #4th Q & A - Client Name
        lbl4 = Label(self, text = "What is the Client's Full Name?")
        lbl4.grid(column = 2, row = 3)
        txt3 = Entry(self, width=38)
        txt3.grid(column = 3, row = 3)

        #5th Q & A - Client E-mail
        lbl5 = Label(self, text = "What is the Client's E-mail?")
        lbl5.grid(column = 2, row = 4)
        txt4 = Entry(self, width=38)
        txt4.grid(column = 3, row = 4)

        #6th Q & A - Attorney Contact
        lbl6 = Label(self, text = "Who is the Attorney?")
        lbl6.grid(column = 2, row = 5)
        drop3 = OptionMenu(self, clicked_attorneys, *attorneys)
        drop3.grid(column = 3, row = 5)


        def seller_email():
            property_address = txt1.get()
            listing_agent = clicked_agents.get()
            client_name = txt3.get()
            client_email = txt4.get()
            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Your Sale of ' + property_address
            mailItem.BodyFormat = 1

            if listing_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[listing_agent][1]
                mailItem.CC = rcrg[listing_agent][0] + " amy@rickcoxrealty.com;"

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {client_name}!<br><br></p>
                <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I work with {agent_name} and will be assisting with your sale of {property_address}. Attached, you will find copies of the fully-executed contract and any addenda or disclosures in conjunction with your sale.<br><br></p>
                <p class=MsoNormal>It should be noted that as a part of your real estate transaction, we will need to have a Termite inspection done at your property within 30 days of closing. Either myself or our Team Administrator, Amy Foldes (CC’d on this e-mail), will reach out to schedule a convenient time and date to complete this inspection!<br><br></p>
                <p class=MsoNormal>Should you have any questions regarding closing or any aspect of the sale leading up to that point, please feel free to reach out me. My congratulations to you on your upcoming home sale!<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
            
            mailItem.HTMLBody = html_body
            mailItem.To = client_email
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        def attorney_email():
            property_address = txt1.get()
            listing_agent = clicked_agents.get()
            commission = clicked_commissions.get()
            attorney_contact = clicked_attorneys.get()

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'New Seller-Side Transaction - ' + property_address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Attorney E-mail'

            #To: Operating Logic - Dictionary Call
            if attorney_contact == "Other":
                Attorney_Name = " "
                mailItem.To = " "
            else:
                Attorney_Name = attorney[attorney_contact][1]
                mailItem.To = attorney[attorney_contact][0]

            #CC: Operating Logic - Dictionary Call
            if listing_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[listing_agent][1]
                mailItem.CC = rcrg[listing_agent][0] + " amy@rickcoxrealty.com;"
            
            #Operation Logic - Commission String based on Option Menu choice
            if commission == "Other":
                commission_split = "*ENTER COMMISSION HERE*"
            elif commission == "6% Total, 3/3":
                commission_split = "6% total, split 3% to the Listing Agent and 3% to the Selling Agent"
            elif commission == "5.5% Total, 2.75/2.75":
                commission_split = "5.5% total, split 2.75% to the Listing Agent and 2.75% to the Selling Agent"
            else:
                commission_split = "5% total, split 2.5% to the Listing Agent and 2.5% to the Selling Agent"    

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Attorney_Name}!<br><br></p>
                <p class=MsoNormal>{agent_name}'s client would like to use your office for the deed preparation necessary for their sale of {property_address}. Please find the ratified contract, transaction information sheet and tax record attached!<br><br></p>
                <p class=MsoNormal> Please note that the commission for this transaction will be {commission_split}. Additionally, our brokerage will charge a $395.00 Administrative Fee to the seller at closing. Should the purchaser's attorney ask, we would like both checks mailed to our office at <b> 2913 Fox Chase Lane, Midlothian, VA 23112. </b> Thank you! <br><br></p>
                <p class=MsoNormal>CC: {agent_name}, Listing Agent; Amy Foldes, Team Administrator<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()


        #Execute Button
        submit_button = Button(self, text = 'Submit', command = lambda:[seller_email(), attorney_email()])
        submit_button.grid(column = 3, row = 6)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=3, row=7)


class TeamMeeting(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        label = tk.Label(self, text = 'Team Meeting \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "Back to Main", 
                         command = lambda: controller.up_frame("WelcomePage"))
        bou1.grid(column=1, row=1)

        def team_meeting_email():
            #Variables for Subject & Body of E-mail
            #If-then Logic to help determine proceeding Thursday.
            #Weekdays are notated by Integers 0 = Monday, 1 = Tuesday, 2 = Thursday, etc.
            e = datetime.datetime.now()
            if e.weekday() == 0:
                e += datetime.timedelta(days=2)
            elif e.weekday() == 1:
                e += datetime.timedelta(days=1)
            elif e.weekday() == 2:
                e += datetime.timedelta(days=0)
            elif e.weekday() == 3:
                e += datetime.timedelta(days=6)
                
            #e += datetime.timedelta(days=1) #Adds 1 Day to the Current Date called in Variable 'e' using datetime.now()
            Date = "%s/%s" % (e.month, e.day) #Establishes Formatting for Date Noted in E-mail
            Time = "9:30AM"

            #Morning or Afternoon Variable Determination Logic
            if e.strftime("%p") == "AM": 
                AM_PM = "Morning"
            elif e.strftime("%p") == "PM":
                AM_PM = "Afternoon"

            #E-mail Dispatch Code
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = f'RCRG Team Meeting - {Date} @ {Time}'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {AM_PM}, RCRG Team!<br><br></p>
                <p class=MsoNormal>Friendly reminder that we will be conducting our meeting this Wednesday, {Date}, at {Time}. Hope to see you all then!<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            #BCC List for Distribution of E-mail
            mailItem.BCC = "eleni@findhomerva.com; melanies1274@yahoo.com; rick@rickcoxrealty.com; brettmlynes@gmail.com; gregsellsva@gmail.com; mbarlowrvahomes@gmail.com; kathyhole1@gmail.com; tundehasthekey@gmail.com; amy@rickcoxrealty.com; morgan@byrdpm.com; soldbygizzirva@gmail.com; benny@richmondwithbenny.com"
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        #Submit Button to Execute E-mail Function
        submit_button = Button(self, text = 'Submit', command = team_meeting_email)
        submit_button.grid(column =2, row =2)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=2, row=3)


class ZillowTeam(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        label = tk.Label(self, text = 'Active Zillow Team \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        back_button = tk.Button(self, text = "Back to Main", 
                        command = lambda: controller.up_frame("WelcomePage"))
        back_button.grid(column=1, row=1)


        clicked_team = StringVar()
        clicked_team.set("Teams")


        lbl1 = Label(self, text = "Which Team is ON to Receive Zillow Leads")
        lbl1.grid(column = 2, row = 3)
        drop1 = OptionMenu(self, clicked_team, *teams)
        drop1.grid(column =3, row=3)


        def team_meeting_email():
            team_on = clicked_team.get()
            if team_on == "Alpha":
                team_off = "Bravo"
            elif team_on == "Bravo":
                team_off = "Alpha"

            e = datetime.datetime.now()
            if e.strftime("%p") == "AM": 
                AM_PM = "Morning"
            elif e.strftime("%p") == "PM":
                AM_PM = "Afternoon"

            #E-mail Dispatch Code
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = f'{team_on} Team Active for Zillow Leads'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {AM_PM}, {team_on} Team!<br><br></p>
                <p class=MsoNormal>This is a friendly reminder that you are now on to receive Zillow leads for the week.<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            #Addressee Operating Logic - Database
            if team_on == "Alpha":
                mailItem.To = ""
                mailItem.CC = "eleni@findhomerva.com; melanies1274@yahoo.com; brettmlynes@gmail.com; benny@richmondwithbenny.com"
                mailItem.BCC = "rick@rickcoxrealty.com"
            elif team_on == "Bravo":
                mailItem.To = ""
                mailItem.CC = "GregSellsVA@Gmail.com; tundehasthekey@gmail.com; kathyhole1@gmail.com; soldbygizzirva@gmail.com"
                mailItem.BCC = "Rick@RickCoxRealty.com"
            else:
                mailItem.To= ""
                mailItem.CC = ""
                mailItem.BCC = ""


            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

            #E-mail Dispatch Code - Team OFF
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = f'{team_off} Team Paused for Zillow Leads'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {AM_PM}, {team_off} Team! <br><br></p>
                <p class=MsoNormal>This is a friendly reminder that you are now paused for Zillow leads for the week. <br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            #Addressee Operating Logic - Database
            if team_off == "Alpha":
                mailItem.To = ""
                mailItem.CC = "eleni@findhomerva.com; melanies1274@yahoo.com; brettmlynes@gmail.com;"
                mailItem.BCC = "rick@rickcoxrealty.com"
            elif team_off == "Bravo":
                mailItem.To = ""
                mailItem.CC = "GregSellsVA@Gmail.com; tundehasthekey@gmail.com; kathyhole1@gmail.com;"
                mailItem.BCC = "Rick@RickCoxRealty.com"
            else:
                mailItem.To= ""
                mailItem.CC = ""
                mailItem.BCC = ""


            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        submit_button = Button(self, text = 'Submit', command = team_meeting_email)
        submit_button.grid(column =2, row =4)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=2, row=5)

        
class BuyerZillow(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        label = tk.Label(self, text = 'Buyer Zillow Review \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "Back to Main", 
                        command = lambda: controller.up_frame("WelcomePage"))
        bou1.grid(column=1, row=1)

        clicked_agents = StringVar()
        clicked_agents.set("Agents")

        #1st Q & A - Property Address
        lbl1 = Label(self, text = "What is the Property Address?")
        lbl1.grid(column = 2, row = 0)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 0)

        #2nd Q & A - Agent
        lbl2 = Label(self, text = "Who was the Selling Agent?")
        lbl2.grid(column = 2, row = 1)
        drop1 = OptionMenu(self, clicked_agents, *agents)
        drop1.grid(column = 3, row = 1)

        #4th Q & A - Client Name
        lbl4 = Label(self, text = "What is the Client's Full Name?")
        lbl4.grid(column = 2, row = 2)
        txt2 = Entry(self, width=38)
        txt2.grid(column = 3, row = 2)

        #5th Q & A - Client E-mail
        lbl5 = Label(self, text = "What is the Client's E-mail?")
        lbl5.grid(column = 2, row = 3)
        txt3 = Entry(self, width=38)
        txt3.grid(column = 3, row = 3)


        def buyer_zillow_email():
            property_address = txt1.get()
            selling_agent = clicked_agents.get()
            client_name = txt2.get()
            client_email = txt3.get()
            
            if selling_agent == "Melanie":    
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzhc49pnkdwcp_31kcj'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Eleni":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzptgpd1ekdu1_5886z'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Rick":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzagn27cmg8i1_1y8fb'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Greg":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUy0ye6f74j2tl_9hc6g'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Christine":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5sfhm4mnvgp_9y704'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Tunde":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUxqyuo617aoeh_1zmzs'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Matt":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUt5yuy1qjw3k9_8jmj4'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Brett":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5xj0syh2ivd_wb2k'>Link to Zillow Review for {selling_agent}</a>"
            elif selling_agent == "Kathy":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZU10zkkrzdo9csp_4ok8d'>Link to Zillow Review for {selling_agent}</a>"
            else:
                zillow_link = "**PUT ZILLOW LINK HERE**"



            olApp = win32.Dispatch('Outlook.Application')
            
            olNS = olApp.GetNameSpace('MAPI')
            
            mailItem = olApp.CreateItem(0)
            
            if selling_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[selling_agent][1]
                mailItem.CC = rcrg[selling_agent][0] + " amy@rickcoxrealty.com;"
            
            mailItem.To = client_email
            mailItem.Subject = f'Zillow Review - Your Sale of {property_address}'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {client_name}! <br><br></p>
                <p class=MsoNormal>Congratulations on your purchase of {property_address}! Thank you for choosing to work with {agent_name} and our real estate group. <br><br></p>
                <p class=MsoNormal>The purpose of this e-mail and the reason I am reaching out to you today is to ask for your assistance in determining how our team and team members preformed for you. If you have some time and felt like our team provided exceptional service to you, would you be able to give {selling_agent} a 5-star review on Zillow via the link below? We would love to hear from you! <br><br></p>
                <p class=MsoNormal>{zillow_link} <br><br></p>
                <p class=MsoNormal>We sincerely appreciate your business! <br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
            
            mailItem.HTMLBody = html_body
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        #Execute Button
        submit_button = Button(self, text = 'Submit', command = buyer_zillow_email)
        submit_button.grid(column = 3, row = 4)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=3, row=5)


class SellerZillow(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        label = tk.Label(self, text = 'Seller Zillow Review \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "Back to Main", 
                        command = lambda: controller.up_frame("WelcomePage"))
        bou1.grid(column=1, row=1)

        clicked_agents = StringVar()
        clicked_agents.set("Agents")

        #1st Q & A - Property Address
        lbl1 = Label(self, text = "What is the Property Address?")
        lbl1.grid(column = 2, row = 0)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 0)

        #2nd Q & A - Agent
        lbl2 = Label(self, text = "Who was the Listing Agent?")
        lbl2.grid(column = 2, row = 1)
        drop1 = OptionMenu(self, clicked_agents, *agents)
        drop1.grid(column = 3, row = 1)

        #4th Q & A - Client Name
        lbl4 = Label(self, text = "What is the Client's Full Name?")
        lbl4.grid(column = 2, row = 2)
        txt2 = Entry(self, width=38)
        txt2.grid(column = 3, row = 2)

        #5th Q & A - Client E-mail
        lbl5 = Label(self, text = "What is the Client's E-mail?")
        lbl5.grid(column = 2, row = 3)
        txt3 = Entry(self, width=38)
        txt3.grid(column = 3, row = 3)

        def seller_zillow_email():
            property_address = txt1.get()
            listing_agent = clicked_agents.get()
            client_name = txt2.get()
            client_email = txt3.get()
            
            if listing_agent == "Melanie":    
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzhc49pnkdwcp_31kcj'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Eleni":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzptgpd1ekdu1_5886z'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Rick":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzagn27cmg8i1_1y8fb'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Greg":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUy0ye6f74j2tl_9hc6g'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Christine":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5sfhm4mnvgp_9y704'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Tunde":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUxqyuo617aoeh_1zmzs'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Matt":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUt5yuy1qjw3k9_8jmj4'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Brett":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5xj0syh2ivd_wb2k'>Link to Zillow Review for {listing_agent}</a>"
            elif listing_agent == "Kathy":
                zillow_link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZU10zkkrzdo9csp_4ok8d'>Link to Zillow Review for {listing_agent}</a>"
            else:
                zillow_link = "**PUT ZILLOW LINK HERE**"



            olApp = win32.Dispatch('Outlook.Application')
            
            olNS = olApp.GetNameSpace('MAPI')
            
            mailItem = olApp.CreateItem(0)
            
            if listing_agent == "Other":
                agent_name = " "
                mailItem.CC = " "
            else:
                agent_name = rcrg[listing_agent][1]
                mailItem.CC = rcrg[listing_agent][0] + " amy@rickcoxrealty.com;"
            
            mailItem.To = client_email
            mailItem.Subject = f'Zillow Review - Your Sale of {property_address}'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {client_name}! <br><br></p>
                <p class=MsoNormal>Congratulations on your sale of {property_address}! Thank you for choosing to work with {agent_name} and our real estate group. <br><br></p>
                <p class=MsoNormal>The purpose of this e-mail and the reason I am reaching out to you today is to ask for your assistance in determining how our team and team members preformed for you. If you have some time and felt like our team provided exceptional service to you, would you be able to give {listing_agent} a 5-star review on Zillow via the link below? We would love to hear from you!<br><br></p>
                <p class=MsoNormal>{zillow_link} <br><br></p>
                <p class=MsoNormal>We sincerely appreciate your business! <br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
            
            mailItem.HTMLBody = html_body
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        #Execute Button
        submit_button = Button(self, text = 'Submit', command = seller_zillow_email)
        submit_button.grid(column = 3, row = 4)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=3, row=5)


class NewListing(tk.Frame):


    def __init__(self, parent, controller):
            tk.Frame.__init__(self, parent)
            self.controller = controller
            self.id = controller.id

            label = tk.Label(self, text = 'New Listing Folder Creation \n' + controller.id.get(), font = controller.titlefont)
            label.grid(column=1, row=0)

            bou1 = tk.Button(self, text = "Back to Main", 
                            command = lambda: controller.up_frame("WelcomePage"))
            bou1.grid(column=1, row=1)

            clicked_agents = StringVar()
            clicked_agents.set("Agents")

            lbl1 = Label(self, text = "What is the Property Address?")
            lbl1.grid(column = 2, row = 0)
            txt1 = Entry(self, width=38)
            txt1.grid(column = 3, row = 0)

            #2nd Q & A - Agent
            lbl2 = Label(self, text = "Who is the Listing Agent?")
            lbl2.grid(column = 2, row = 1)
            drop1 = OptionMenu(self, clicked_agents, *agents)
            drop1.grid(column = 3, row = 1)

            def seller_folder():
                
                if os.getcwd() != 'C:\\Users\\rcrgr\\Desktop\\E-mail Programs':
                    os.chdir('C:\\Users\\rcrgr\\Desktop\\E-mail Programs')
                
                property_address = txt1.get()
                listing_agent = clicked_agents.get()

                fillpdfs.get_form_fields("Transaction Info Sheet(Fillable).pdf")

                data_dict = {'Property Address': property_address, 'City': '', 'State': '', 'Zip': '', 'County': '',
                        'CVRMLS': '', 'Sales Price': '', 'Offer Date_af_date': '', 'Date2_af_date': '',
                        'Rat-Date_af_date': '', 'Closing Date_af_date': '', 'List Price': '', 'Closing Costs Paid by Seller': '',
                        'Seller': '', 'Purchaser': '', 'Seller 1': '', 'Seller 2': '', 'Seller Email 1': '', 'Seller Email 2': '',
                        'Seller Cell': '', 'Seller Work': '', 'Seller Home': '', 'Seller Fax': '', 'Seller Forwarding Address': '',
                        'Seller City': '', 'Seller State': '', 'Seller Zip': '', 'Buyer 1': '', 'Buyer 2': '',
                        'Buyer Email': '', 'Buyer Email 2': '', 'Buyer Cell': '', 'Buyer Work': '', 'Buyer Home': '',
                        'Buyer Fax': '', 'Home Warranty': '', 'Home Inspec\x98on Co': '', 'Termite Co': '', 'FuelOil Co': '',
                        'Well  Sep\x98c Co': '', 'Lender': '', 'Loan Officer Name': '', 'Loan Officer Phone': '', 'Loan Officer Email': '',
                        'Seller Attorney Firm': '', 'Seller Attorney Contact': '', 'Seller Office Phone': '', 'Seller Attorney Fax': '',
                        'Seller Attorney Email': '', 'Buyer Attorney Firm': '', 'Buyer Attorney Contact': '', 'Buyer Attorney Office Phone': '',
                        'Buyer Attorney Fax': '', 'Buyer Attorney Email': '', 'HOA Name': '', 'HOA Mgmt Co': '', 'HOA Phone': '', 'HOA Email': '',
                        'Listing Company Name': 'The Rick Cox Realty Group', 'Listing Agent Name': rcrg[listing_agent][1], 'Transaction Coordinator': 'Harrison Goehring - harrison@rickcoxrealty.com', 'Listing Agent Phone': rcrg[listing_agent][4],
                        'Listing Agent E-mail': rcrg[listing_agent][0], 'Selling Company Name': '', 'Selling Agent Name': '', 'Selling Agent TC': '',
                        'Selling Agent Phone': '', 'Selling Agent Email': '', 'Escrow Deposit': '', 'Held by': '', 'Commission': '',
                        'Transac\x98on Fee': '395.00', 'Referral Fee': '', 'Paid to': '', 'Referral Address': '', 'Reset': ''}
            
                fillpdfs.write_fillable_pdf('Transaction Info Sheet(Fillable).pdf', 'Transaction Info Sheet(f).pdf', data_dict)

                if listing_agent == "Other":
                    path = " "
                else:
                    path = rcrg[listing_agent][3]
                    os.chdir(path)

                os.mkdir(property_address)

                os.chdir(f"{path}\\{property_address}")

                os.mkdir("Contract-Addenda")
                os.mkdir("Invoices-Inspections")
                os.mkdir("Photos")

                shutil.copy('C:\\Users\\rcrgr\\Desktop\\E-mail Programs\\Transaction Info Sheet(f).pdf', f'{path}\\{property_address}\\Contract-Addenda')
                
            #Execute Button
            submit_button = Button(self, text = 'Submit', command = seller_folder)
            submit_button.grid(column = 3, row = 3)

            close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
            close_button.grid(column=3, row=4)


if __name__ == '__main__':
    app = MainFrame()
    app.mainloop()

