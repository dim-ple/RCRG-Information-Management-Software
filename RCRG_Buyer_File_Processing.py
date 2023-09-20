import tkinter as tk
from tkinter import font as tkfont
from tkinter import StringVar, BooleanVar, Label, Entry, OptionMenu, Radiobutton, Button, Toplevel

# Library for Calendar & Data entry widgets used in the UI
from tkcalendar import Calendar, DateEntry

# Library for copying and flattening PDFs
import shutil

#Imports os library to help interact with MS Outlook and the Windows OS
import os
import win32com.client as win32

# Imports our system time reference module which will help determine the proper greeting in our e-mail templates
from DateAndTime import Time
from DateAndTime import datetime

# Imports a library to help fill our PDFs
from fillpdf import fillpdfs 

# Imports the sqlite3 library so we can access the DB, run and commit querys
import sqlite3

# Empty lists to be utilized by the tkinter UI. Populated by the brute force for-loops below
# **Eventually will be replaced with SQL query population**
from SQLPopList import SQLPopList


class MainFrame(tk.Tk):

    # Constructor Method setting our window size, font, font size, container, frame ID
    def __init__(self, *args, **kwargs):
        
        # Calls our tkinter constructor
        tk.Tk.__init__(self, *args, **kwargs)
       
        # Sets the font for our MainFrame and all child frames defined later in the program
        self.titlefont = tkfont.Font(family = 'Verdana', size = 12,
                                     weight = "bold", slant = 'roman')
        
        # To make things simple, we're setting our parent and any child frames to the grid set-up. As much as I'd just liek to pack everything, 
        # labels and entry boxes may need to pair up on the same row.
        container = tk.Frame()
        container.grid(row=0, column=0, sticky='nesw')

        # Sets the base dimensions of our MainFrame and any children of the MainFrame (this will be inherited by most frames defined later in the program)
        self.geometry('1000x800')
        
        # Sets our class ID to a string variable to be used late when setting our welcome message on each frame
        self.id = tk.StringVar()
        self.id.set("RCRG Admin")

        # Initilize an empty dictionary that will serve as our frame stack
        self.listing = {}
        
        # Iterates through all of our created child frames, appends them to our listing dictionary stack so user can transition from frame to frame when
        # a frame is selected and the up_frame method is called
        for p in (WelcomePage, BuyerTran):
            page_name = p.__name__
            frame = p(parent = container, controller = self)
            frame.grid(row=0, column=0, sticky='nsew')
            self.listing[page_name] = frame
        
        # When the program starts, up_frame method is called and the user's landing page will always be the WelcomePage frame
        self.up_frame('WelcomePage')

    # When a child frame is called, the program pages "up" to the selected frame
    def up_frame(self, page_name):
        # When the method is called, the page variable is set to which ever page from the listing dictionary we would like to navigate to
        page = self.listing[page_name]
        
        # Then we use the tkraise method to push up to the frame we need in our stack
        page.tkraise()



class WelcomePage(tk.Frame):    

    def __init__(self, parent, controller):
        
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        # Add labels and buttons to access other frames
        label = tk.Label(self, text = 'Welcome Page \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        bou1 = tk.Button(self, text = "New Buyer Transaction", 
                        command = lambda: controller.up_frame("BuyerTran"))
        bou1.grid(column=2, row=1)

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

        # Populates our agent names list and database for use with the UI and fillpdf
        rcrg_agent_options, rcrg_agent_db = SQLPopList('rcrg')
        lender_options, lender_db = SQLPopList('lenders')
        attorney_options, attorney_db = SQLPopList('attorneys')

        #1st Q & A - Property Address
        prop_add_lbl = Label(self, text = "What is the Property Address?")
        prop_add_lbl.grid(column = 2, row = 0)
        prop_add_ent = Entry(self, width=38)
        prop_add_ent.grid(column = 3, row = 0)

        #City
        prop_city_lbl = Label(self, text = "What is the Property City?")
        prop_city_lbl.grid(column = 2, row = 1)
        prop_city_ent = Entry(self, width=20)
        prop_city_ent.grid(column = 3, row = 1)

        #Zip
        prop_zip_lbl = Label(self, text = "What is the Property Zip?")
        prop_zip_lbl.grid(column = 2, row = 2)
        prop_zip_ent = Entry(self, width=8)
        prop_zip_ent.grid(column = 3, row = 2)

        #County
        prop_county_lbl = Label(self, text = "What is the Property County?")
        prop_county_lbl.grid(column = 2, row = 3)
        prop_county_ent = Entry(self, width=20)
        prop_county_ent.grid(column = 3, row = 3)

        #MLS Number
        mls_lbl = Label(self, text = "What is the MLS Number?")
        mls_lbl.grid(column = 2, row = 4)
        mls_ent = Entry(self, width=10)
        mls_ent.grid(column = 3, row = 4)

        #Sales Price
        sp_lbl = Label(self, text = "What is the Sales Price? (Do not enter special characters)")
        sp_lbl.grid(column = 2, row = 5)
        sp_ent = Entry(self, width=20)
        sp_ent.grid(column = 3, row = 5)

        #List Price
        lp_lbl = Label(self, text = "What was the List Price? (Do not enter special characters)")
        lp_lbl.grid(column = 2, row = 6)
        lp_ent = Entry(self, width=20)
        lp_ent.grid(column = 3, row = 6)

        #List Date
        list_date_lbl = Label(self, text = "What was the List Date?")
        list_date_lbl.grid(column = 2, row = 7)
        list_date_picker = DateEntry(self, width=16, background="magenta3", foreground="white", bd=2)
        list_date_picker.grid(column = 3, row = 7)
        
        #Offer Date
        offer_date_lbl = Label(self, text = "What was the Offer Date?")
        offer_date_lbl.grid(column = 2, row = 8)
        offer_date_picker = DateEntry(self, width=16, background="magenta3", foreground="white", bd=2)
        offer_date_picker.grid(column = 3, row = 8)

        #Ratification Date
        ratif_date_lbl = Label(self, text = "What was the Date of Ratification?")
        ratif_date_lbl.grid(column = 2, row = 9)
        ratif_date_picker = DateEntry(self, width=16, background="magenta3", foreground="white", bd=2)
        ratif_date_picker.grid(column = 3, row = 9)

        #Close Date
        close_date_lbl = Label(self, text = "What is the Closing Date?")
        close_date_lbl.grid(column = 2, row = 10)
        close_date_picker = DateEntry(self, width=16, background="magenta3", foreground="white", bd=2)
        close_date_picker.grid(column = 3, row = 10)

        #Seller Paid Closing Costs
        spcc_lbl = Label(self, text = "Seller Paid Closing Costs? (Do not enter special characters)")
        spcc_lbl.grid(column = 2, row = 11)
        spcc_ent = Entry(self, width=38)
        spcc_ent.grid(column = 3, row = 11)

        #Seller Name
        seller_name_lbl = Label(self, text = "What is the Seller(s) Full Name? For multiple names, separate with a ';'")
        seller_name_lbl.grid(column = 2, row = 12)
        seller_name_ent = Entry(self, width=38)
        seller_name_ent.grid(column = 3, row = 12)

        #2nd Q & A - Selling Agent
        sell_agent_lbl = Label(self, text = "Who is the Selling Agent?")
        sell_agent_lbl.grid(column = 2, row = 13)
        sell_agent_drop = OptionMenu(self, clicked_agents, *rcrg_agent_options)
        sell_agent_drop.grid(column = 3, row = 13)

        #3rd Q & A - Commission
        comm_lbl = Label(self, text = "What is the Selling Agent's Commission")
        comm_lbl.grid(column = 2, row = 14)
        comm_ent = Entry(self, width=8)
        comm_ent.grid(column = 3, row = 14)

        #Transaction Fee (Radio 3 option - 395, 495, 0)
        admin_fee_lbl = Label(self, text="What is the Admin Fee?")
        admin_fee_lbl.grid(column=2, row=15)
        radio0 = Radiobutton(self, text="N/A", variable = clicked_admin_fee,
                            value="0")
        radio0.grid(column=3, row=15)
        radio495 = Radiobutton(self, text="$495", variable = clicked_admin_fee,
                            value="495")
        radio495.grid(column=4, row=15)
        radio395 = Radiobutton(self, text="$395", variable = clicked_admin_fee,
                             value="395")
        radio395.grid(column=5, row=15)
        clicked_admin_fee.set("395")

        #5th Q & A - Client Name
        client_name_lbl = Label(self, text = "What is the Client's Full Name? For multiple names, separate with a ';'")
        client_name_lbl.grid(column=2, row=16)
        client_name_ent = Entry(self, width=38)
        client_name_ent.grid(column=3, row=16)

        #Client Phone Number(s)
        client_phone_lbl = Label(self, text = "What is the Client's Phone Number? For multiple numbers, separate with a ';'")
        client_phone_lbl.grid(column=2, row=17)
        client_phone_ent = Entry(self, width=38)
        client_phone_ent.grid(column=3, row=17)

        #6th Q & A - Lender
        lender_lbl = Label(self, text = "Who is the Lender?")
        lender_lbl.grid(column = 2, row = 18)
        lender_drop = OptionMenu(self, clicked_lenders, *lender_options)
        lender_drop.grid(column = 3, row = 18)

        #7th Q & A - EMD
        emd_lbl = Label(self, text="Do we have the EMD?")
        emd_lbl.grid(column = 2, row = 19)
        radio_emd_yes = Radiobutton(self, text = "Yes", variable = clicked_boolean,
                            value=True)
        radio_emd_yes.grid(column = 3, row = 19)
        radio_emd_no = Radiobutton(self, text = "No", variable = clicked_boolean,
                            value=False)
        radio_emd_no.grid(column = 4, row = 19)

        #8th Q & A - Attorney Contact
        attorney_lbl = Label(self, text = "Who is the Attorney?")
        attorney_lbl.grid(column = 2, row = 20)
        attorney_drop = OptionMenu(self, clicked_attorneys, *attorney_options)
        attorney_drop.grid(column = 3, row = 20)

        #9th Q & A - Client E-mail
        client_email_lbl = Label(self, text = "What is the Client's E-mail? For multiple emails, separate with a ';'")
        client_email_lbl.grid(column = 2, row = 21)
        client_email_ent = Entry(self, width=38)
        client_email_ent.grid(column = 3, row = 21)

        #10th Q & A - Listing Agent Name
        la_name_lbl = Label(self, text = "Who is the Listing Agent?")
        la_name_lbl.grid(column = 2, row = 22)
        la_name_ent = Entry(self, width=38)
        la_name_ent.grid(column = 3, row = 22)

        la_phone_lbl = Label(self, text = "What is the Listing Agent's Phone Number?")
        la_phone_lbl.grid(column=2, row=23)
        la_phone_ent = Entry(self, width=38)
        la_phone_ent.grid(column = 3, row = 23)

        la_email_lbl = Label(self, text = "What is the Listing Agent's E-mail?")
        la_email_lbl.grid(column=2, row=24)
        la_email_ent = Entry(self, width=38)
        la_email_ent.grid(column = 3, row = 24)

        listing_broker_lbl = Label(self, text = "What is the name of the Listing Agent's Brokerage?")
        listing_broker_lbl.grid(column=2, row=25)
        listing_broker_ent = Entry(self, width=38)
        listing_broker_ent.grid(column = 3, row = 25)

        #Execute Button
        submit_button = Button(self, text = "Submit",
                               command = lambda:[buyer_email(), attorney_email(), listing_agent_email(), lender_email()])
        submit_button.grid(column = 3, row = 26)

        new_folder_button = Button(self, text = "Create New Folder",
                                   command = lambda:[buyer_folder()])
        new_folder_button.grid(column = 3, row = 27)

        clear_fields_button = Button(self, text = "Reset Fields",
                                     command = lambda:[clear_fields()])
        clear_fields_button.grid(column = 3, row = 28)

        fill_fields_test_button = Button(self, text = "Test Fill the Fields",
                                         command = lambda:[fill_fields()])
        fill_fields_test_button.grid(column=3, row=29)

        close_button = Button(self, text = "Close the Window",
                              command = controller.destroy)
        close_button.grid(column = 3, row = 30)


        def buyer_folder():
            if os.getcwd() != 'C:\\Users\\rcrgr\\Desktop\\E-mail Programs':
                os.chdir('C:\\Users\\rcrgr\\Desktop\\E-mail Programs')
        
            property_address = prop_add_ent.get()
            city = prop_city_ent.get()
            zip = prop_zip_ent.get()
            county = prop_county_ent.get()
            mls = mls_ent.get()
            sp = sp_ent.get()
            spcc = spcc_ent.get()
            lp = lp_ent.get()
            list_date = list_date_picker.get()
            offer_date = offer_date_picker.get()
            ratif_date = ratif_date_picker.get()
            close_date = close_date_picker.get()
            seller1 = seller_name_ent.get()
            seller2 = ' '
            selling_agent = clicked_agents.get()
            listing_agent = la_name_ent.get()
            listing_email = la_email_ent.get()
            listing_phone = la_phone_ent.get()
            listing_broker = listing_broker_ent.get()
            commission = comm_ent.get()
            client1 = client_name_ent.get()
            client2 = ' '
            client_phone1 = client_phone_ent.get()
            client_phone2 = ' '
            client_email1 = client_email_ent.get()
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

            if ";" in client_phone1:
                i = client_phone1.find(";")
                client_phone2 = client_phone1[(i+2):]
                client_phone1 = client_phone1[0:i]

            if ";" in seller1:
                i = seller1.find(";")
                seller2 = seller1[(i+2):]
                seller1 = seller1[0:i]


            fillpdfs.get_form_fields("Transaction Info Sheet(Fillable).pdf")

            data_dict = {'Property Address': property_address, 'City': city, 'State': 'VA', 'Zip': zip, 'County': county,
                        'CVRMLS': mls, 'Sales Price': sp, 'Offer Date_af_date': offer_date, 'Date2_af_date': list_date,
                        'Rat-Date_af_date': ratif_date, 'Closing Date_af_date': close_date, 'List Price': lp, 'Closing Costs Paid by Seller': spcc,
                        'Seller': '', 'Purchaser': 'Yes', 'Seller 1': seller1, 'Seller 2': seller2, 'Seller Email 1': '', 'Seller Email 2': '',
                        'Seller Cell': '', 'Seller Work': '', 'Seller Home': '', 'Seller Fax': '', 'Seller Forwarding Address': '',
                        'Seller City': '', 'Seller State': '', 'Seller Zip': '', 'Buyer 1': client1, 'Buyer 2': client2,
                        'Buyer Email': client_email1, 'Buyer Email 2': client_email2, 'Buyer Cell': client_phone1, 'Buyer Work': client_phone2, 'Buyer Home': '',
                        'Buyer Fax': '', 'Home Warranty': '', 'Home Inspec\x98on Co': '', 'Termite Co': '', 'FuelOil Co': '',
                        'Well  Sep\x98c Co': '', 'Lender': lender_db[lender_contact][1], 'Loan Officer Name': lender_db[lender_contact][2] + " " + lender_db[lender_contact][3], 'Loan Officer Phone': lender_db[lender_contact][4], 'Loan Officer Email': lender_db[lender_contact][5],
                        'Seller Attorney Firm': '', 'Seller Attorney Contact': '', 'Seller Office Phone': '', 'Seller Attorney Fax': '',
                        'Seller Attorney Email': '', 'Buyer Attorney Firm': attorney_db[attorney_contact][1], 'Buyer Attorney Contact': attorney_db[attorney_contact][2] + " " + attorney_db[attorney_contact][3], 'Buyer Attorney Office Phone': attorney_db[attorney_contact][4],
                        'Buyer Attorney Fax': '', 'Buyer Attorney Email': attorney_db[attorney_contact][5], 'HOA Name': '', 'HOA Mgmt Co': '', 'HOA Phone': '', 'HOA Email': '',
                        'Listing Company Name': listing_broker, 'Listing Agent Name': listing_agent, 'Transaction Coordinator': '', 'Listing Agent Phone': listing_phone,
                        'Listing Agent E-mail': listing_email, 'Selling Company Name': rcrg_agent_db[selling_agent][7], 'Selling Agent Name': selling_agent, 'Selling Agent TC': 'Harrison Goehring - harrison@rickcoxrealty.com',
                        'Selling Agent Phone': rcrg_agent_db[selling_agent][3], 'Selling Agent Email': rcrg_agent_db[selling_agent][4], 'Escrow Deposit': '', 'Held by': '', 'Commission': commission + ' to Selling Agent',
                        'Transac\x98on Fee': admin_fee, 'Referral Fee': '', 'Paid to': '', 'Referral Address': '', 'Reset': ''}
            
            fillpdfs.write_fillable_pdf('Transaction Info Sheet(Fillable).pdf', 'Transaction Info Sheet(f).pdf', data_dict)
            
            path = " " if selling_agent == "Other" else rcrg_agent_db[selling_agent][8]
            
            

            
            os.chdir(path)

            os.mkdir(property_address)

            os.chdir(f"{path}\\{property_address}")

            os.mkdir("Contract-Addenda")
            os.mkdir("Invoices-Inspections")

            shutil.copy('C:\\Users\\rcrgr\\Desktop\\E-mail Programs\\Transaction Info Sheet(f).pdf', f'{path}\\{property_address}\\Contract-Addenda')

        def buyer_email():
            property_address = prop_add_ent.get()
            selling_agent = clicked_agents.get()
            client_name1 = client_name_ent.get()
            client_email = client_email_ent.get()
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

            agent_name = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][1] + " " + rcrg_agent_db[selling_agent][2]
            mailItem.CC = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][4]

            html_body = f"""
                <p class=MsoNormal>Good {Time}, {Address_To_Client}!<br><br></p>
                <p class=MsoNormal>My name is Amy Foldes and I am the Team Administrator for the Rick Cox Realty Group. I work with {agent_name} and will be assisting with your purchase of {property_address}. Attached, you will find copies of the fully-executed contract and any addenda or disclosures in conjunction with your closing.<br><br></p>
                <p class=MsoNormal>Should you have any questions regarding closing or any aspect of the transaction leading up to that point, please feel free to reach out me. My congratulations to you on your upcoming home purchase!<br><br></p>
                <p class=MsoNormal>CC: Your agent, {agent_name}; Team Administrator, Amy Foldes; <br><br></p>
                <p class=MsoNormal>Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Amy Foldes</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Team Administrator @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Amy@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Amy@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
            
            mailItem.HTMLBody = html_body
            mailItem.To = client_email
            

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('amy@rickcoxrealty.com')))

            mailItem.Display()

        def attorney_email():
            property_address = prop_add_ent.get()
            selling_agent = clicked_agents.get()
            commission = comm_ent.get()
            attorney_contact = clicked_attorneys.get()

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'New Purchase-Side Transaction - ' + property_address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Attorney E-mail'

            #To: Operating Logic - Dictionary Call
            attorney_name = " " if (attorney_contact == "Other") else attorney_db[attorney_contact][2]
            mailItem.To = " " if (attorney_contact == "Other") else attorney_db[attorney_contact][5]

            #CC: Operating Logic - Dictionary Call
            agent_name = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][1] + " " + rcrg_agent_db[selling_agent][2]
            mailItem.CC = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][4] + "; amy@rickcoxrealty.com;"
                

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {attorney_name}!<br><br></p>
                <p class=MsoNormal>{agent_name}'s client would like to use your office for the title and settlement work needed for their purchase of {property_address}. Please find the ratified contract, transaction information sheet and tax record attached!<br><br></p>
                <p class=MsoNormal> Please note that the selling agent's commission for this transaction will be {commission}. Additionally, our brokerage will charge a $395.00 Administrative Fee to the purchaser at closing. Please overnight both checks to our office at <b> 2913 Fox Chase Lane, Midlothian, VA 23112. </b> Thank you! <br><br></p>
                <p class=MsoNormal>CC: {agent_name}, Selling Agent; Team Administrator, Amy Foldes;<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Amy Folders</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Team Administrator @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Amy@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Amy@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('amy@rickcoxrealty.com')))

            mailItem.Display()

        def lender_email():
            lender_contact = clicked_lenders.get()
            EMD_Status = clicked_boolean.get()
            property_address = prop_add_ent.get()
            selling_agent = clicked_agents.get()
            client_name1 = client_name_ent.get()
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
            lender_name = " " if (lender_contact == "Other") else lender_db[lender_contact][2]
            mailItem.To = " " if (lender_contact == "Other") else lender_db[lender_contact][5]
                
            agent_name = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][1] + " " + rcrg_agent_db[selling_agent][2]
            mailItem.CC = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][4] + "; amy@rickcoxrealty.com;"

            #EMD Logic
            if EMD_Status == True:
                EMD = "We have received the earnest money deposit, please find a copy of the check attached."
            elif EMD_Status == False:
                EMD = "We have not yet received the earnest money deposit. Once received, we will forward a copy of the check to you!"
            else:
                EMD = ""
                
            html_body =f"""
                <p class=MsoNormal>Good {Time}, {lender_name}!<br><br></p>
                <p class=MsoNormal>Please find a ratified contract attached for {agent_name}'s {client_email_Message}! {EMD}<br><br></p>
                <p class=MsoNormal> Kind regards, <br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Amy Foldes</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Team Administrator @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Amy@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Amy@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('amy@rickcoxrealty.com')))

            mailItem.Display()

        def listing_agent_email():
            property_address = prop_add_ent.get()
            selling_agent = clicked_agents.get()
            attorney_contact = clicked_attorneys.get()
            listing_agent = la_name_ent.get()
            #listing_email = la_email_ent.get()
            
            # Conditional statment to determine string included in our intro e-mail to the listing agent. It is far more likely that our clients have already decided on
            # an Attorney or Title Company at the get-go, so the first check is whether the option "Other" was not selected.
            if clicked_attorneys.get() != "Other":
                attorney_msg = f"Our purchaser will be using {attorney_db[attorney_contact][1]} for their title and settlement needs. The primary contact will be {attorney_db[attorney_contact][2]} {attorney_db[attorney_contact][3]}, their e-mail is {attorney_db[attorney_contact][5]}."
            else:
                attorney_msg = "Our purchaser has not yet decided on who they will be using for their title and settlement needs. Once they have decided, I will let you know!"

            
            olApp = win32.Dispatch('Outlook.Application')
            olNS = olApp.GetNameSpace('MAPI')

            mailItem = olApp.CreateItem(0)
            mailItem.Subject = 'Coordinator Introduction - ' + property_address
            mailItem.BodyFormat = 1
            mailItem.HTMLBody = 'Coordinator Introduction'

            #To: Operating Logic - Dictionary Call
            #if listing_email == "":
                #mailItem.To = " "
            #else:
                #mailItem.To = listing_email

            #CC: Operating Logic - Dictionary Call
            agent_name = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][1] + " " + rcrg_agent_db[selling_agent][2]
            mailItem.CC = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][4]
                

            html_body =f"""
                <p class=MsoNormal>Good {Time}, {listing_agent}!<br><br></p>
                <p class=MsoNormal>My name is Amy Foldes and I am the Team Administrator for the Rick Cox Realty Group. I will be assisting {agent_name} and their client on the purchase of {property_address}. I look forward to working with you!<br><br></p>
                <p class=MsoNormal>{attorney_msg} Would you mind providing me with the contact for the Seller's Attorney or Title Company who will be handling the deed preparation for the Seller once that information becomes available?<br><br></p>
                <p class=MsoNormal>Additionally, would your seller be willing to share who their current utility providers for Electricity, Water/Sewer, Internet, Trash and Gas are?<br><br></p>
                <p class=MsoNormal>CC: {agent_name}, Selling Agent; Team Administrator, Amy Foldes;<br><br></p>
                <p class=MsoNormal>Kind regards & thanks,<br><br></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Amy Foldes</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Team Administrator @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Amy@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Amy@RickCoxRealty.com</span> </a><o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
            """
                
            mailItem.HTMLBody = html_body

            mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('amy@rickcoxrealty.com')))

            mailItem.Display()
        
        def clear_fields():
            clicked_agents.set("Agents")
            clicked_lenders.set("Lenders")
            clicked_attorneys.set("Attorneys")
            prop_zip_ent.delete("0", "end")
            prop_city_ent.delete("0", "end")
            prop_county_ent.delete("0", "end")
            mls_ent.delete("0", "end")
            sp_ent.delete("0", "end")
            lp_ent.delete("0", "end")
            spcc_ent.delete("0", "end")
            seller_name_ent.delete("0", "end")
            prop_add_ent.delete("0", "end")
            comm_ent.delete("0", "end")
            client_name_ent.delete("0", "end")
            client_email_ent.delete("0", "end")
            client_phone_ent.delete("0", "end")
            la_name_ent.delete("0", "end")
            listing_broker_ent.delete("0", "end")
            la_email_ent.delete("0", "end")
            clicked_boolean.set(False)
            clicked_admin_fee.set("395")

        def fill_fields():
            clicked_agents.set("Rick Cox")
            clicked_lenders.set("Alcova Mortgage (Eric)")
            clicked_attorneys.set("Atlantic Coast Settlement Services (Susan)")
            prop_zip_ent.insert("0", "23112")
            prop_city_ent.insert("0", "Chesterfield")
            prop_county_ent.insert("0", "Chesterfield")
            mls_ent.insert("0", "2311441")
            sp_ent.insert("0", "300000")
            lp_ent.insert("0", "300000")
            spcc_ent.insert("0", "3000")
            seller_name_ent.insert("0", "Billy Bob; Millie Bo")
            prop_add_ent.insert("0", "123 Busy Street")
            comm_ent.insert("0", "3%")
            client_name_ent.insert("0", "Robbie Bob; Bobbie Rob")
            client_email_ent.insert("0", "www@gmail.com; eee@gmail.com")
            client_phone_ent.insert("0", "804-555-5555; 912-104-5556")
            la_name_ent.insert("0", "Test Agent")
            listing_broker_ent.insert("0", "Test Broker")
            la_phone_ent.insert("0", "804-111-3332")
            la_email_ent.insert("0", "Test@email.com")



if __name__ == '__main__':
    app = MainFrame()
    app.mainloop()