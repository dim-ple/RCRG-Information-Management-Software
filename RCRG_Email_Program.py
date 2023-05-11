import tkinter as tk
from tkinter import font as tkfont
from tkinter import StringVar, BooleanVar, Label, Entry, OptionMenu, Radiobutton, Button

import shutil

import os
import win32com.client as win32

from Contact_Dictionaries import rcrg, lender, attorney

from DateAndTime import Time
from DateAndTime import datetime

from fillpdf import fillpdfs

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
        self.id.set("Harrison Goehring")

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

        clicked_attorneys = StringVar()
        clicked_attorneys.set("Attorneys")


        #1st Q & A - Property Address
        lbl1 = Label(self, text = "What is the Property Address?")
        lbl1.grid(column = 2, row = 0)
        txt1 = Entry(self, width=38)
        txt1.grid(column = 3, row = 0)

        #2nd Q & A - Agent
        lbl2 = Label(self, text = "Who is the Selling Agent?")
        lbl2.grid(column = 2, row = 1)
        drop1 = OptionMenu(self, clicked_agents, *agents)
        drop1.grid(column = 3, row = 1)

        #3rd Q & A - Commission
        lbl3 = Label(self, text = "What is the Selling Agent's Commission")
        lbl3.grid(column = 2, row = 2)
        txt2 = Entry(self, width=8)
        txt2.grid(column = 3, row = 2)

        #5th Q & A - Client Name
        lbl4 = Label(self, text = "What is the Client's Full Name? For Multiple Names, separate with a ';'")
        lbl4.grid(column = 2, row = 3)
        txt3 = Entry(self, width=38)
        txt3.grid(column = 3, row = 3)

        #6th Q & A - Lender
        lbl4 = Label(self, text = "Who is the Lender?")
        lbl4.grid(column = 2, row = 4)
        drop2 = OptionMenu(self, clicked_lenders, *lenders)
        drop2.grid(column = 3, row = 4)

        #7th Q & A - EMD
        lbl5 = Label(self, text="Do we have the EMD?")
        lbl5.grid(column = 2, row = 5)
        radio1 = Radiobutton(self, text = "Yes", variable = clicked_boolean,
                            value=True)
        radio1.grid(column = 3, row = 5)
        radio2 = Radiobutton(self, text = "No", variable = clicked_boolean,
                            value=False)
        radio2.grid(column = 4, row = 5)

        #8th Q & A - Attorney Contact
        lbl6 = Label(self, text = "Who is the Attorney?")
        lbl6.grid(column = 2, row = 6)
        drop3 = OptionMenu(self, clicked_attorneys, *attorneys)
        drop3.grid(column = 3, row = 6)

        #9th Q & A - Client E-mail
        lbl7 = Label(self, text = "What is the Client's E-mail?")
        lbl7.grid(column = 2, row = 7)
        txt4 = Entry(self, width=38)
        txt4.grid(column = 3, row = 7)

        #10th Q & A - Listing Agent Name
        lbl7 = Label(self, text = "What is the Listing Agent's Name")
        lbl7.grid(column = 2, row = 8)
        txt5 = Entry(self, width=38)
        txt5.grid(column = 3, row = 8)

        #11th Q & A - Listing Agent Name
        lbl7 = Label(self, text = "What is the Listing Agent's E-mail")
        lbl7.grid(column = 2, row = 9)
        txt6 = Entry(self, width=38)
        txt6.grid(column = 3, row = 9)


        #Initialize Variables for E-mail program
        Property_Address = txt1.get()
        Selling_Agent = clicked_agents.get()
        Commission = txt2.get()
        Client_Name = txt3.get()
        client_email = txt4.get()
        Lender_Contact = clicked_lenders.get()
        EMD_Status = clicked_boolean.get()
        Attorney_Contact = clicked_attorneys.get()
        Attorney_Status = False
        Listing_Agent = txt5.get()
        Listing_Email = txt6.get()

        if clicked_attorneys.get() == "Other":
            Attorney_Status = False
        else:
            Attorney_Status = True


        def buyer_folder():
            if os.getcwd() != 'C:\\Users\\rcrgr\\Desktop\\E-mail Programs':
                os.chdir('C:\\Users\\rcrgr\\Desktop\\E-mail Programs')
        
            Property_Address = txt1.get()
            Selling_Agent = clicked_agents.get()
            Listing_Agent = txt5.get()
            Listing_Email = txt6.get()
            Commission = txt2.get()
            Client1 = txt3.get()
            Client2 = ' '
            client_email1 = txt4.get()
            client_email2 = ' '
            Attorney_Contact = clicked_attorneys.get()
            Lender_Contact = clicked_lenders.get()
            

            if ";" in Client1:
                i = Client1.find(";")
                Client2 = Client1[(i+2):]
                Client1 = Client1[0:i]

            if ";" in client_email1:
                i = client_email1.find(";")
                client_email2 = client_email1[(i+2):]
                client_email1 = client_email1[0:i]

            fillpdfs.get_form_fields("Transaction Info Sheet(Fillable).pdf")


            data_dict = {'Property Address': Property_Address, 'City': '', 'State': '', 'Zip': '', 'County': '',
                        'CVRMLS': '', 'Sales Price': '', 'Offer Date_af_date': '', 'Date2_af_date': '',
                        'Rat-Date_af_date': '', 'Closing Date_af_date': '', 'List Price': '', 'Closing Costs Paid by Seller': '',
                        'Seller': '', 'Purchaser': '', 'Seller 1': '', 'Seller 2': '', 'Seller Email 1': '', 'Seller Email 2': '',
                        'Seller Cell': '', 'Seller Work': '', 'Seller Home': '', 'Seller Fax': '', 'Seller Forwarding Address': '',
                        'Seller City': '', 'Seller State': '', 'Seller Zip': '', 'Buyer 1': Client1, 'Buyer 2': Client2,
                        'Buyer Email': client_email1, 'Buyer Email 2': client_email2, 'Buyer Cell': '', 'Buyer Work': '', 'Buyer Home': '',
                        'Buyer Fax': '', 'Home Warranty': '', 'Home Inspec\x98on Co': '', 'Termite Co': '', 'FuelOil Co': '',
                        'Well  Sep\x98c Co': '', 'Lender': Lender_Contact, 'Loan Officer Name': lender[Lender_Contact][2], 'Loan Officer Phone': lender[Lender_Contact][3], 'Loan Officer Email': lender[Lender_Contact][0],
                        'Seller Attorney Firm': '', 'Seller Attorney Contact': '', 'Seller Office Phone': '', 'Seller Attorney Fax': '',
                        'Seller Attorney Email': '', 'Buyer Attorney Firm': attorney[Attorney_Contact][2], 'Buyer Attorney Contact': attorney[Attorney_Contact][3], 'Buyer Attorney Office Phone': '',
                        'Buyer Attorney Fax': '', 'Buyer Attorney Email': attorney[Attorney_Contact][0], 'HOA Name': '', 'HOA Mgmt Co': '', 'HOA Phone': '', 'HOA Email': '',
                        'Listing Company Name': '', 'Listing Agent Name': Listing_Agent, 'Transaction Coordinator': '', 'Listing Agent Phone': '',
                        'Listing Agent E-mail': Listing_Email, 'Selling Company Name': 'The Rick Cox Realty Group', 'Selling Agent Name': rcrg[Selling_Agent][1], 'Selling Agent TC': 'Harrison Goehring - harrison@rickcoxrealty.com',
                        'Selling Agent Phone': rcrg[Selling_Agent][4], 'Selling Agent Email': rcrg[Selling_Agent][0], 'Escrow Deposit': '', 'Held by': '', 'Commission': Commission + ' to Selling Agent',
                        'Transac\x98on Fee': '395.00', 'Referral Fee': '', 'Paid to': '', 'Referral Address': '', 'Reset': ''}
            
            fillpdfs.write_fillable_pdf('Transaction Info Sheet(Fillable).pdf', 'Transaction Info Sheet(f).pdf', data_dict)
            
            if Selling_Agent == "Other":
                path = " "
            else:
                path = rcrg[Selling_Agent][3]
                os.chdir(path)

            os.mkdir(Property_Address)

            os.chdir(f"{path}\\{Property_Address}")

            os.mkdir("Contract-Addenda")
            os.mkdir("Invoices-Inspections")

            shutil.copy('C:\\Users\\rcrgr\\Desktop\\E-mail Programs\\Transaction Info Sheet(f).pdf', f'{path}\\{Property_Address}\\Contract-Addenda')

        def buyer_email():
            Property_Address = txt1.get()
            Selling_Agent = clicked_agents.get()
            Commission = txt2.get()
            Client_Name = txt3.get()
            client_email = txt4.get()

            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Your New Purchase of ' + Property_Address
            mailItem.BodyFormat = 1

            if Selling_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Selling_Agent][1]
                mailItem.CC = rcrg[Selling_Agent][0] + " amy@rickcoxrealty.com;"

            html_body = f"""
                <p class=MsoNormal>Good {Time}, {Client_Name}!<br><br></p>
                <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I work with {Agent_Name} and will be assisting with your purchase of {Property_Address}. Attached, you will find copies of the fully-executed contract and any addenda or disclosures in conjunction with your closing.<br><br></p>
                <p class=MsoNormal>Should you have any questions regarding closing or any aspect of the transaction leading up to that point, please feel free to reach out me. My congratulations to you on your upcoming home purchase!<br><br></p>
                <p class=MsoNormal>CC: Your agent, {Agent_Name}; Team Administrator, Amy Foldes; <br><br></p>
                <p class=MsoNormal>Kind Regards, <br><br></p>
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
            Property_Address = txt1.get()
            Selling_Agent = clicked_agents.get()
            Commission = txt2.get()
            Attorney_Contact = clicked_attorneys.get()

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'New Purchase-Side Transaction - ' + Property_Address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Attorney E-mail'

            #To: Operating Logic - Dictionary Call
            if Attorney_Contact == "Other":
                Attorney_Name = " "
                mailItem.To = " "
            else:
                Attorney_Name = attorney[Attorney_Contact][1]
                mailItem.To = attorney[Attorney_Contact][0]

            #CC: Operating Logic - Dictionary Call
            if Selling_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Selling_Agent][1]
                mailItem.CC = rcrg[Selling_Agent][0] + " amy@rickcoxrealty.com;"
                

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Attorney_Name}!<br><br></p>
                <p class=MsoNormal>{Agent_Name}'s client would like to use your office for the title and settlement work needed for their purchase of {Property_Address}. Please find the ratified contract, transaction information sheet and tax record attached!<br><br></p>
                <p class=MsoNormal> Please note that the selling agent's commission for this transaction will be {Commission}. Additionally, our brokerage will chargea $395.00 Administrative Fee to the purchaser at closing. Please overnight both checks to our office at <b> 2913 Fox Chase Lane, Midlothian, VA 23112. </b> Thank you! <br><br></p>
                <p class=MsoNormal>CC: {Agent_Name}, Selling Agent; Team Administrator, Amy Foldes;<br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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
            Lender_Contact = clicked_lenders.get()
            EMD_Status = clicked_boolean.get()
            Property_Address = txt1.get()
            Selling_Agent = clicked_agents.get()
            Client_Name1 = txt3.get()
            Client_Name2 = ' '
            
            if ";" in Client_Name1:
                i = Client_Name1.find(";")
                Client_Name2 = Client_Name1[(i+2):]
                Client_Name1 = Client_Name1[0:i]
            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = f'New Purchase Contract - {Property_Address} for ({Client_Name1})'
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Lender E-mail'

            #Addressee Operating Logic - Database
            if Lender_Contact == "Other":
                Lender_Name = " "
                mailItem.To = " "
            else:
                Lender_Name = lender[Lender_Contact][1]
                mailItem.To = lender[Lender_Contact][0]
                
            if Selling_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Selling_Agent][1]
                mailItem.CC = rcrg[Selling_Agent][0] + " amy@rickcoxrealty.com;"

            #EMD Logic
            if EMD_Status == True:
                EMD = "We have received the earnest money deposit, please find a copy of the check attached."
            elif EMD_Status == False:
                EMD = "We have not yet received the earnest money deposit. Once received, we will forward a copy of the check to you!"
            else:
                EMD = ""
                
            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Lender_Name}!<br><br></p>
                <p class=MsoNormal>Please find a ratified contract attached for {Agent_Name}'s client, {Client_Name}! {EMD}<br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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
            Property_Address = txt1.get()
            Selling_Agent = clicked_agents.get()
            Commission = txt2.get()
            Attorney_Contact = clicked_attorneys.get()
            Listing_Agent = txt5.get()
            Listing_Email = txt6.get()


            if Attorney_Status == True:
                attorney_msg = f"Our purchaser will be using {attorney[Attorney_Contact][2]} for their title and settlement needs. The primary contact will be {attorney[Attorney_Contact][3]}, their e-mail is {attorney[Attorney_Contact][0]}."
            elif Attorney_Status == False:
                attorney_msg = "Our purchaser has not yet decided on who they will be using for their title and settlement needs. Once they have decided, I will let you know!"
            else:
                attorney_msg = ""

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Coordinator Introduction - ' + Property_Address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Coordinator Introduction'

            #To: Operating Logic - Dictionary Call
            if Listing_Email == "":
                mailItem.To = " "
            else:
                mailItem.To = Listing_Email

            #CC: Operating Logic - Dictionary Call
            if Selling_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Selling_Agent][1]
                mailItem.CC = rcrg[Selling_Agent][0] + " amy@rickcoxrealty.com;"
                

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Listing_Agent}!<br><br></p>
                <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I will be assisting {Agent_Name} and their client on the purchase of {Property_Address}. I look forward to working with you!<br><br></p>
                <p class=MsoNormal>{attorney_msg} Would you mind providing me with the contact for the Seller's Attorney or Title Company who will be handling the deed preparation for the Seller once that information becomes available?<br><br></p>
                <p class=MsoNormal>Additionally, would your seller be willing to share who their current utility providers for Electricity, Water/Sewer, Internet, Trash and Gas are?<br><br></p>
                <p class=MsoNormal>CC: {Agent_Name}, Selling Agent; Team Administrator, Amy Foldes;<br><br></p>
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
        
        
        #Execute Button
        submit_button = Button(self, text = 'Submit',
                               command = lambda:[buyer_email(), attorney_email(), listing_agent_email(), lender_email()])
        submit_button.grid(column = 3, row = 10)

        new_folder_button = Button(self, text = "Create New Folder",
                                   command = lambda:[buyer_folder()])
        new_folder_button.grid(column=3, row=11)

        close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
        close_button.grid(column=3, row=12)


class  SellerTran(tk.Frame):
    
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


        #Initialize Variables for E-mail program
        Property_Address = txt1.get()
        Listing_Agent = clicked_agents.get()
        Commission = clicked_commissions.get()
        Client_Name = txt3.get()
        Client_Email = txt4.get()
        Attorney_Contact = clicked_attorneys.get()


        def seller_email():
            Property_Address = txt1.get()
            Listing_Agent = clicked_agents.get()
            Client_Name = txt3.get()
            Client_Email = txt4.get()

            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Your Sale of ' + Property_Address
            mailItem.BodyFormat = 1

            if Listing_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Listing_Agent][1]
                mailItem.CC = rcrg[Listing_Agent][0] + " amy@rickcoxrealty.com;"

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Client_Name}!<br><br></p>
                <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I work with {Agent_Name} and will be assisting with your sale of {Property_Address}. Attached, you will find copies of the fully-executed contract and any addenda or disclosures in conjunction with your sale.<br><br></p>
                <p class=MsoNormal>It should be noted that as a part of your real estate transaction, we will need to have a Termite inspection done at your property within 30 days of closing. Either myself or our Team Administrator, Amy Foldes (CC’d on this e-mail), will reach out to schedule a convenient time and date to complete this inspection!<br><br></p>
                <p class=MsoNormal>Should you have any questions regarding closing or any aspect of the sale leading up to that point, please feel free to reach out me. My congratulations to you on your upcoming home sale!<br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
            
            mailItem.HTMLBody = html_body
            mailItem.To = Client_Email
            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

            mailItem.Display()

        def attorney_email():
            Property_Address = txt1.get()
            Listing_Agent = clicked_agents.get()
            Commission = clicked_commissions.get()
            Attorney_Contact = clicked_attorneys.get()

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'New Seller-Side Transaction - ' + Property_Address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Attorney E-mail'

            #To: Operating Logic - Dictionary Call
            if Attorney_Contact == "Other":
                Attorney_Name = " "
                mailItem.To = " "
            else:
                Attorney_Name = attorney[Attorney_Contact][1]
                mailItem.To = attorney[Attorney_Contact][0]

            #CC: Operating Logic - Dictionary Call
            if Listing_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Listing_Agent][1]
                mailItem.CC = rcrg[Listing_Agent][0] + " amy@rickcoxrealty.com;"
            
            #Operation Logic - Commission String based on Option Menu choice
            if Commission == "Other":
                Commission_Split = "*ENTER COMMISSION HERE*"
            elif Commission == "6% Total, 3/3":
                Commission_Split = "6% total, split 3% to the Listing Agent and 3% to the Selling Agent"
            elif Commission == "5.5% Total, 2.75/2.75":
                Commission_Split = "5.5% total, split 2.75% to the Listing Agent and 2.75% to the Selling Agent"
            else:
                Commission_Split = "5% total, split 2.5% to the Listing Agent and 2.5% to the Selling Agent"    

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Attorney_Name}!<br><br></p>
                <p class=MsoNormal>{Agent_Name}'s client would like to use your office for the deed preparation necessary for their sale of {Property_Address}. Please find the ratified contract, transaction information sheet and tax record attached!<br><br></p>
                <p class=MsoNormal> Please note that the commission for this transaction will be {Commission_Split}. Additionally, our brokerage will charge a $395.00 Administrative Fee to the seller at closing. Should the purchaser's attorney ask, we would like both checks mailed to our office at <b> 2913 Fox Chase Lane, Midlothian, VA 23112. </b> Thank you! <br><br></p>
                <p class=MsoNormal>CC: {Agent_Name}, Listing Agent; Amy Foldes, Team Administrator<br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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
                <p class=MsoNormal> Kind Regards, <br><br></p>
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
            mailItem.BCC = "eleni@findhomerva.com; melanies1274@yahoo.com; rick@rickcoxrealty.com; brettmlynes@gmail.com; gregsellsva@gmail.com; mbarlowrvahomes@gmail.com;  kathyhole1@gmail.com; tundehasthekey@gmail.com; christinemottleyrva@gmail.com; amy@rickcoxrealty.com"
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

        Team_On = clicked_team.get()


        def team_meeting_email():
            Team_On = clicked_team.get()
            if Team_On == "Alpha":
                Team_Off = "Bravo"
            elif Team_On == "Bravo":
                Team_Off = "Alpha"

            e = datetime.datetime.now()
            if e.strftime("%p") == "AM": 
                AM_PM = "Morning"
            elif e.strftime("%p") == "PM":
                AM_PM = "Afternoon"

            #E-mail Dispatch Code
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = f'{Team_On} Team Active for Zillow Leads'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {AM_PM}, {Team_On} Team!<br><br></p>
                <p class=MsoNormal>This is a friendly reminder that you are now on to receive Zillow leads for the week.<br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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
            if Team_On == "Alpha":
                mailItem.To = ""
                mailItem.CC = ""
                mailItem.BCC = "eleni@findhomerva.com; melanies1274@yahoo.com; brettmlynes@gmail.com; kathyhole1@gmail.com; rick@rickcoxrealty.com"
            elif Team_On == "Bravo":
                mailItem.To = ""
                mailItem.CC = ""
                mailItem.BCC = " GregSellsVA@Gmail.com; MBarlowRVAHomes@gmail.com; ChristineMottleyRVA@gmail.com; Rick@RickCoxRealty.com"
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
            mailItem.Subject = f'{Team_Off} Team Paused for Zillow Leads'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {AM_PM}, {Team_Off} Team! <br><br></p>
                <p class=MsoNormal>This is a friendly reminder that you are now paused for Zillow leads for the week. <br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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
            if Team_Off == "Alpha":
                mailItem.To = ""
                mailItem.CC = ""
                mailItem.BCC = "eleni@findhomerva.com; melanies1274@yahoo.com; brettmlynes@gmail.com; kathyhole1@gmail.com; rick@rickcoxrealty.com"
            elif Team_Off == "Bravo":
                mailItem.To = ""
                mailItem.CC = ""
                mailItem.BCC = "GregSellsVA@Gmail.com; MBarlowRVAHomes@gmail.com; ChristineMottleyRVA@gmail.com; Rick@RickCoxRealty.com"
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


        Property_Address = txt1.get()
        Selling_Agent = clicked_agents.get()
        Client_Name = txt2.get()
        Client_Email = txt3.get()

        def buyer_zillow_email():
            Property_Address = txt1.get()
            Selling_Agent = clicked_agents.get()
            Client_Name = txt2.get()
            Client_Email = txt3.get()
            
            if Selling_Agent == "Melanie":    
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzhc49pnkdwcp_31kcj'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Eleni":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzptgpd1ekdu1_5886z'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Rick":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzagn27cmg8i1_1y8fb'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Greg":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUy0ye6f74j2tl_9hc6g'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Christine":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5sfhm4mnvgp_9y704'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Jessica":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZU15heqyb7atkw9_6y17r'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Tunde":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUxqyuo617aoeh_1zmzs'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Matt":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUt5yuy1qjw3k9_8jmj4'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Brett":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5xj0syh2ivd_wb2k'>Link to Zillow Review for {Selling_Agent}</a>"
            elif Selling_Agent == "Kathy":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZU10zkkrzdo9csp_4ok8d'>Link to Zillow Review for {Selling_Agent}</a>"
            else:
                Zillow_Link = "**PUT ZILLOW LINK HERE**"



            olApp = win32.Dispatch('Outlook.Application')
            
            olNS = olApp.GetNameSpace('MAPI')
            
            mailItem = olApp.CreateItem(0)
            
            if Selling_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Selling_Agent][1]
                mailItem.CC = rcrg[Selling_Agent][0] + " amy@rickcoxrealty.com;"
            
            mailItem.To = Client_Email
            mailItem.Subject = f'Zillow Review - Your Sale of {Property_Address}'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Client_Name}! <br><br></p>
                <p class=MsoNormal>Congratulations on your purchase of {Property_Address}! Thank you for choosing to work with {Agent_Name} and our real estate group. <br><br></p>
                <p class=MsoNormal>The purpose of this e-mail and the reason I am reaching out to you today is to ask for your assistance in determining how our team and team members preformed for you. If you have some time and felt like our team provided exceptional service to you, would you be able to give {Selling_Agent} a 5-star review on Zillow via the link below? We would love to hear from you! <br><br></p>
                <p class=MsoNormal>{Zillow_Link} <br><br></p>
                <p class=MsoNormal>We sincerely appreciate your business! <br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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


        Property_Address = txt1.get()
        Listing_Agent = clicked_agents.get()
        Client_Name = txt2.get()
        Client_Email = txt3.get()

        def seller_zillow_email():
            Property_Address = txt1.get()
            Listing_Agent = clicked_agents.get()
            Client_Name = txt2.get()
            Client_Email = txt3.get()
            
            if Listing_Agent == "Melanie":    
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzhc49pnkdwcp_31kcj'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Eleni":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzptgpd1ekdu1_5886z'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Rick":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUzagn27cmg8i1_1y8fb'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Greg":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUy0ye6f74j2tl_9hc6g'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Christine":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5sfhm4mnvgp_9y704'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Jessica":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZU15heqyb7atkw9_6y17r'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Tunde":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUxqyuo617aoeh_1zmzs'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Matt":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUt5yuy1qjw3k9_8jmj4'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Brett":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZUw5xj0syh2ivd_wb2k'>Link to Zillow Review for {Listing_Agent}</a>"
            elif Listing_Agent == "Kathy":
                Zillow_Link = f"<a href='https://www.zillow.com/reviews/write/?s=X1-ZU10zkkrzdo9csp_4ok8d'>Link to Zillow Review for {Listing_Agent}</a>"
            else:
                Zillow_Link = "**PUT ZILLOW LINK HERE**"



            olApp = win32.Dispatch('Outlook.Application')
            
            olNS = olApp.GetNameSpace('MAPI')
            
            mailItem = olApp.CreateItem(0)
            
            if Listing_Agent == "Other":
                Agent_Name = " "
                mailItem.CC = " "
            else:
                Agent_Name = rcrg[Listing_Agent][1]
                mailItem.CC = rcrg[Listing_Agent][0] + " amy@rickcoxrealty.com;"
            
            mailItem.To = Client_Email
            mailItem.Subject = f'Zillow Review - Your Sale of {Property_Address}'
            mailItem.BodyFormat = 1

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {Client_Name}! <br><br></p>
                <p class=MsoNormal>Congratulations on your sale of {Property_Address}! Thank you for choosing to work with {Agent_Name} and our real estate group. <br><br></p>
                <p class=MsoNormal>The purpose of this e-mail and the reason I am reaching out to you today is to ask for your assistance in determining how our team and team members preformed for you. If you have some time and felt like our team provided exceptional service to you, would you be able to give {Listing_Agent} a 5-star review on Zillow via the link below? We would love to hear from you!<br><br></p>
                <p class=MsoNormal>{Zillow_Link} <br><br></p>
                <p class=MsoNormal>We sincerely appreciate your business! <br><br></p>
                <p class=MsoNormal> Kind Regards, <br><br></p>
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

            Property_Address = txt1.get()
            Listing_Agent = clicked_agents.get()


            def seller_folder():
                Property_Address = txt1.get()
                Listing_Agent = clicked_agents.get()

                if Listing_Agent == "Other":
                    path = " "
                else:
                    path = rcrg[Listing_Agent][3]
                    os.chdir(path)

                os.mkdir(Property_Address)

                os.chdir(f"{path}\\{Property_Address}")

                os.mkdir("Contract-Addenda")
                os.mkdir("Invoices-Inspections")
                os.mkdir("Photos")

            #Execute Button
            submit_button = Button(self, text = 'Submit', command = seller_folder)
            submit_button.grid(column = 3, row = 3)

            close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
            close_button.grid(column=3, row=4)

if __name__ == '__main__':
    app = MainFrame()
    app.mainloop()

