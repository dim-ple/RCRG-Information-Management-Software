# Imports tkinter for UI design
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

# Import dictionaries with Agent, lender and attorney info - will remove from repository when SQL query is full fleshed out and 
# lists can be populated with data from the database
from Contact_Dictionaries import rcrg, lender, attorney

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



lenders = []
attorneys =[]


# Loops to populate our agents, lenders and attorneys lists
for lend in lender:
    lenders.append(lend)

for office_name in attorney:
    attorneys.append(office_name)


# Creates our MainFrame class which will be our Parent class for the UI
class MainFrame(tk.Tk):

    # Constructor Method setting our window size, font, font size, container, frame ID
    def __init__(self, *args, **kwargs):
        
        # Calls our tkinter constructor
        tk.Tk.__init__(self, *args, **kwargs)
       
        # Sets the font for our MainFrame and all child frames defined later in the program
        self.titlefont = tkfont.Font(family = 'Verdana', size = 12,
                                     weight = "bold", slant = 'roman')
        
        # To make things simple, we're setting our parent and any child frames to the grid set-up. As much as I'd just like to pack everything, 
        # labels and entry boxes may need to pair up on the same row.
        container = tk.Frame()
        container.grid(row=0, column=0, sticky='nesw')

        # Sets the base dimensions of our MainFrame and any children of the MainFrame (this will be inherited by most frames defined later in the program)
        self.geometry('1000x800')
        
        # Sets our class ID to a string variable to be used later when setting our welcome message on each frame
        self.id = tk.StringVar()
        self.id.set("RCRG Admin")

        # Initilize an empty dictionary that will serve as our frame stack
        self.listing = {}
        
        # Iterates through all of our created child frames, appends them to our listing dictionary stack so user can transition from frame to frame when
        # a frame is selected and the up_frame method is called
        for p in (WelcomePage, BuyerTran, SellerTran, TeamMeeting, ZillowTeam, BuyerZillow, SellerZillow, NewListing):
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
    
    # Constructor Method for our WelcomePage (again, our landing page whenever a user utilizes the program), our constructor uses a Model-View-Controller model
    # self is passed to represent our current frame, parent is passed to represent our MainFrame application window or root window
    # and controller which helps us interact with other defined class frames
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
        
        # Initializes our construtor method, our controller and frame ID for our frame with the welcome page as our parent frame, allowing us to keep the same
        # attributes of the parent frame(size)
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        # Header label for our frame to tell the user where they've landed
        label = tk.Label(self, text = 'New Buyer Transaction \n' + controller.id.get(), font = controller.titlefont)
        label.grid(column=1, row=0)

        # Creates a button which will bring the user back up a frame to the Welcome Page
        bou1 = tk.Button(self, text = "Back to Main", 
                        command = lambda: controller.up_frame("WelcomePage"))
        bou1.grid(column=1, row=1)

        # Initializes our String Variables which will convert our selection from the agents, lender admin fee and attorneys dropdown menu into a string for use
        # in referencing dict entries
        clicked_agents = StringVar()
        clicked_agents.set("Agents")

        clicked_lenders = StringVar()
        clicked_lenders.set("Lenders")

        clicked_admin_fee = StringVar()

        clicked_attorneys = StringVar()
        clicked_attorneys.set("Attorneys")

        # Initializes our boolean variable for the EMD radial button selection (do we have the EMD, y/n?)
        clicked_boolean = BooleanVar()

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
        seller_name_lbl = Label(self, text = "What is the Seller(s) Full Name? For Multiple Names, separate with a ';'")
        seller_name_lbl.grid(column = 2, row = 12)
        seller_name_ent = Entry(self, width=10)
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
        client_name_lbl = Label(self, text = "What is the Client's Full Name? For Multiple Names, separate with a ';'")
        client_name_lbl.grid(column=2, row=16)
        client_name_ent = Entry(self, width=38)
        client_name_ent.grid(column=3, row=16)

        #Client Phone Number(s)
        client_phone_lbl = Label(self, text = "What is the Client's Phone Number? For Multiple Numbers, separate with a ';'")
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
        client_email_lbl = Label(self, text = "What is the Client's E-mail?")
        client_email_lbl.grid(column = 2, row = 21)
        client_email_ent = Entry(self, width=38)
        client_email_ent.grid(column = 3, row = 21)

        #10th Q & A - Listing Agent Name
        la_name_lbl = Label(self, text = "Who is the Listing Agent?")
        la_name_lbl.grid(column = 2, row = 22)
        la_name_ent = Entry(self, width=38)
        la_name_ent.grid(column = 3, row = 22)

        # The Buyer Folder function aims to create a new directory, folder structure and creates & fills out a new transaction information sheet in our shared dropbox
        # when we input all the required data for a newly received contract
        def buyer_folder():
           
           # First, the function checks to ensure that our current working directory is within the root folder for our program
           # this is important as we will be writing to a blank transaction information sheet which will be copied to the new file directory later in our program.
            if os.getcwd() != 'C:\\Users\\rcrgr\\Desktop\\E-mail Programs':
                os.chdir('C:\\Users\\rcrgr\\Desktop\\E-mail Programs')
        
            # Retreives all the data entered or selected by the user in the tkinter interface and initializes variables for each        
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
            #listing_email = la_email_ent.get()
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

            # The user is instructed in the tkinter interface to add a semi-colon between client names, e-mails, phone numbers and seller names to indicate to the program
            # that there is a 2nd client, seller, phone # or email. The program checks for the semi-colon using the find method in each of these string fields. If it detects a semi-colon,
            # it initializes an additional variable
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

            # We call the fillpdfs libraries get_form_fields method to pull all the form field names from our blank transaction information sheet
            fillpdfs.get_form_fields("Transaction Info Sheet(Fillable).pdf")

             # We initialize a dictionary that we will later use to fill out the transaction information sheet with the user's entered data.
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
                        'Listing Company Name': '', 'Listing Agent Name': listing_agent, 'Transaction Coordinator': '', 'Listing Agent Phone': '',
                        'Listing Agent E-mail': '', 'Selling Company Name': rcrg_agent_db[selling_agent][7], 'Selling Agent Name': selling_agent, 'Selling Agent TC': 'Harrison Goehring - harrison@rickcoxrealty.com',
                        'Selling Agent Phone': rcrg_agent_db[selling_agent][3], 'Selling Agent Email': rcrg_agent_db[selling_agent][4], 'Escrow Deposit': '', 'Held by': '', 'Commission': commission + ' to Selling Agent',
                        'Transac\x98on Fee': admin_fee, 'Referral Fee': '', 'Paid to': '', 'Referral Address': '', 'Reset': ''}
            
            
            # We then write the user's information into the pdf using the write_fillable_pdf method and our data_dict
            fillpdfs.write_fillable_pdf('Transaction Info Sheet(Fillable).pdf', 'Transaction Info Sheet(f).pdf', data_dict)
            
            # We initialize our path which is referenced from the SQL database for the selected agent
            path = " " if selling_agent == "Other" else rcrg_agent_db[selling_agent][8]
            
            # We change our working directory to our path, we create a folder with the property address as its name
            os.chdir(path)
            os.mkdir(property_address)
            
            # We then move up a level to the folder we just created for the new property transaction. 
            # Once the directory change has been made, we create two sub-folders 
            os.chdir(f"{path}\\{property_address}")
            os.mkdir("Contract-Addenda")
            os.mkdir("Invoices-Inspections")

            # Finally, we use the shutil copy method to copy our transaction information sheet to the new directory and contract-addenda subfolder we created
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
            mailItem.CC = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][4] + "; amy@rickcoxrealty.com;"

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
            mailItem.CC = " " if (selling_agent == "Other") else rcrg_agent_db[selling_agent][4] + "; amy@rickcoxrealty.com;"
                

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
            prop_add_ent.delete("0", "end")
            comm_ent.delete("0", "end")
            client_name_ent.delete("0", "end")
            client_email_ent.delete("0", "end")
            la_name_ent.delete("0", "end")
            #la_email_ent.delete("0", "end")
            clicked_boolean.set(False)
            clicked_admin_fee.set("395")

        def query_creator(table_name, *cols):
    
            rcrg_cols = "(agentfirst, agentlast, agentphone, agentemail, agenttype, agentlicensenum, agentbroker, path)"
            lender_cols ="(lendercompany, lofirst, lolast, lophone, loemail, lpemail)"
            client_cols = "()"
            attorney_cols = "()"
            hoa_cols = "()"
            property_cols = "()"
            
            # Declare string for the beginning of our insert statement
            ins_statement = 'INSERT INTO '
    
            # Declare string for the beginning of our Values statement
            val_statement = 'VALUES ("'
    
            # Declare string for the end of our statement 
            end_char = ')'

            # Declare empty list which will help us adjust our expanded query strings later 
            list_string = []
    
            # Check table name to import correct column names for sql query
            if table_name == 'rcrg':
                ins_statement += rcrg_cols
            elif table_name == 'attorneys':
                ins_statement += attorney_cols
            elif table_name == 'lenders':
                ins_statement += lender_cols
            elif table_name == 'clients':
                ins_statement += client_cols
            elif table_name == 'properties':
                ins_statement += property_cols
            elif table_name == 'hoas':
                ins_statement += hoa_cols
            

            # Once the correct table is found, iterate through all argument columns passed 
            # into the function and append them to the Values statement
            for col in cols:
                val_statement += col + '", "'

            # Convert our string into a list, each character is assigned to an index value
            # within the list
            list_string = list(val_statement)
            
            # We know that we will have to remove 3 characters from the end of the string
            # so we use the pop function at the last index of the string to remove these
            list_string.pop(len(list_string)-1)
            list_string.pop(len(list_string)-1)
            list_string.pop(len(list_string)-1)
            
            # We then join the characters back together to form a new string
            new_string = "".join(list_string)

            # Then we append the string to our end character, assigning it to the values
            # statement
            val_statement = new_string + end_char

            # We then concatenate the Insert & Values statements, assigning them to the query string
            query = ins_statement + val_statement

            # The function returns our query as a string
            return query

        def data_submit(query):
            
            try:
                conn = sqlite3.connect('rcrgbroker.db')
                c = conn.cursor()
                print("Successfully Connected to Database!")

                c.execute(query)

                conn.commit() 

            except:
                print("There was an error connecting to the Database!")

            finally:
                c.close()
                conn.close()
                print("Connection to Database Closed!")
  
        def new_lender_info():

            top = Toplevel(parent)
            top.geometry("450x175")
            top.title("New Lender Info - Input Form")

            lender_name_lbl = Label(top, text = "Mortgage Company Name:")
            lender_name_lbl.grid(column = 2, row = 0)
            lender_name_ent = Entry(top, width=20)
            lender_name_ent.grid(column = 3, row = 0)

            lo_first_lbl = Label(top, text = "Loan Officer First Name:")
            lo_first_lbl.grid(column = 2, row = 1)
            lo_first_ent = Entry(top, width=20)
            lo_first_ent.grid(column = 3, row = 1)

            lo_last_lbl = Label(top, text = "Loan Officer Last Name:")
            lo_last_lbl.grid(column = 2, row = 2)
            lo_last_ent = Entry(top, width=20)
            lo_last_ent.grid(column = 3, row = 2)
           
            lo_phone_lbl = Label(top, text = "Loan Officer Phone #:")
            lo_phone_lbl.grid(column = 2, row = 3)
            lo_phone_ent = Entry(top, width=38)
            lo_phone_ent.grid(column = 3, row = 3)

            # Add Agent Type field (Dropdown selection, default to 'Salesperson')
            lo_email_lbl = Label(top, text = "Loan Officer E-mail:")
            lo_email_lbl.grid(column = 2, row = 4)
            lo_email_ent = Entry(top, width=38)
            lo_email_ent.grid(column = 3, row = 4)

            lp_email_lbl = Label(top, text = "Loan Processor E-mail:")
            lp_email_lbl.grid(column = 2, row = 5)
            lp_email_ent = Entry(top, width=17)
            lp_email_ent.grid(column = 3, row = 5)

            pass_data_button = Button(top, text = "Submit Data",
                                      command = lambda:[data_submit(query_creator("lenders", lender_name_ent.get(), lo_first_ent.get(), lo_last_ent.get(), 
                                                        lo_phone_ent.get(), lo_email_ent.get(), lp_email_ent.get()))])
            pass_data_button.grid(column=3, row=6)

            close_button = Button(top, text = "Close the Window",
                              command= top.destroy)
            close_button.grid(column=3, row=7)
               
        def new_agent_info():
            
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
   
            pass_data_button = Button(top, text = "Submit Data",
                                      command = lambda:[data_submit(query_creator("agents", agent_first_ent.get(), agent_last_ent.get(), agent_cell_ent.get(), 
                                                        agent_email_ent.get(), clicked_agent_type.get(), agent_dpor_ent.get(), agent_broker_ent.get()))])
            pass_data_button.grid(column=3, row=7)

            close_button = Button(top, text = "Close the Window",
                              command= top.destroy)
            close_button.grid(column=3, row=8)

        #Execute Button
        submit_button = Button(self, text = "Submit",
                               command = lambda:[buyer_email(), attorney_email(), listing_agent_email(), lender_email()])
        submit_button.grid(column = 3, row = 23)

        new_folder_button = Button(self, text = "Create New Folder",
                                   command = lambda:[buyer_folder()])
        new_folder_button.grid(column = 3, row = 24)

        clear_fields_button = Button(self, text = "Reset Fields",
                                     command = lambda:[clear_fields()])
        clear_fields_button.grid(column = 3, row = 25)

        close_button = Button(self, text = "Close the Window",
                              command = controller.destroy)
        close_button.grid(column = 3, row = 26)

        new_agent_button = Button(self, text="New Agent",
                                  command = lambda: new_agent_info())
        new_agent_button.grid(column = 4, row = 22)

        new_lender_button = Button(self, text="New Lender",
                                  command = lambda: new_lender_info())
        new_lender_button.grid(column = 4, row = 18)


        '''search_list_agent_btn = Button(self, text="Agent Search",
                                  command = lambda: new_agent_info())
        search_list_agent_btn.grid(column = 4, row = 12)
        
        search_sell_agent_btn = Button(self, text="Agent Search",
                                  command = lambda: new_agent_info())
        search_sell_agent_btn.grid(column = 5, row = 20)'''

class SellerTran(tk.Frame):
    
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self.controller = controller
        self.id = controller.id

        # Default Commission options for SellerTran Function
        commissions = ["6% Total, 3/3",
                       "5.5% Total, 2.75/2.75",
                       "5% Total, 2.5/2.5",
                       "5% Total, 2/3",
                       "Other"]
        
        rcrg_agent_options, rcrg_agent_db  = SQLPopList('rcrg')
        attorney_options, attorney_db = SQLPopList('attorneys')
        
        
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
        drop1 = OptionMenu(self, clicked_agents, *rcrg_agent_options)
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
        drop3 = OptionMenu(self, clicked_attorneys, *attorney_options)
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

            agent_name = " " if (listing_agent == "Other") else rcrg_agent_db[listing_agent][1] + " " + rcrg_agent_db[listing_agent][2]
            mailItem.CC = " " if (listing_agent == "Other") else rcrg_agent_db[listing_agent][4] + "; amy@rickcoxrealty.com;"
            

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
            attorney_name = " " if (attorney_contact == "Other") else attorney_db[attorney_contact][2]
            mailItem.To = " " if (attorney_contact == "Other") else attorney_db[attorney_contact][5]

            #CC: Operating Logic - Dictionary Call
            agent_name = " " if (listing_agent == "Other") else rcrg_agent_db[listing_agent][1] + " " + rcrg_agent_db[listing_agent][2]
            mailItem.CC = " " if (listing_agent == "Other") else rcrg_agent_db[listing_agent][4] + "; amy@rickcoxrealty.com;"
            
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
                <p class=MsoNormal>Good {Time}, {attorney_name}!<br><br></p>
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

        teams = ["Alpha", "Bravo"]

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

        rcrg_agents = SQLPopList('rcrg')

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
        drop1 = OptionMenu(self, clicked_agents, *rcrg_agents)
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

        rcrg_agents = SQLPopList('rcrg')

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
        drop1 = OptionMenu(self, clicked_agents, *rcrg_agents)
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

            rcrg_agents, rcrg_agent_db = SQLPopList('rcrg')

            lbl1 = Label(self, text = "What is the Property Address?")
            lbl1.grid(column = 2, row = 0)
            txt1 = Entry(self, width=38)
            txt1.grid(column = 3, row = 0)

            #2nd Q & A - Agent
            lbl2 = Label(self, text = "Who is the Listing Agent?")
            lbl2.grid(column = 2, row = 1)
            drop1 = OptionMenu(self, clicked_agents, *rcrg_agents)
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
                        'Listing Company Name': 'The Rick Cox Realty Group', 'Listing Agent Name': rcrg_agent_db[listing_agent][1], 'Transaction Coordinator': 'Harrison Goehring - harrison@rickcoxrealty.com', 'Listing Agent Phone': rcrg_agent_db[listing_agent][4],
                        'Listing Agent E-mail': rcrg_agent_db[listing_agent][4], 'Selling Company Name': '', 'Selling Agent Name': '', 'Selling Agent TC': '',
                        'Selling Agent Phone': '', 'Selling Agent Email': '', 'Escrow Deposit': '', 'Held by': '', 'Commission': '',
                        'Transac\x98on Fee': '395.00', 'Referral Fee': '', 'Paid to': '', 'Referral Address': '', 'Reset': ''}
            
                fillpdfs.write_fillable_pdf('Transaction Info Sheet(Fillable).pdf', 'Transaction Info Sheet(f).pdf', data_dict)

                path = " " if listing_agent == "Other" else rcrg_agent_db[listing_agent][8]

                os.chdir(path)
                
                os.mkdir(property_address)

                os.chdir(f"{path}\\{property_address}")

                os.mkdir("Contract-Addenda")
                os.mkdir("Invoices-Inspections")
                os.mkdir("Photos")

                shutil.copy(f'C:\\Users\\rcrgr\\Desktop\\E-mail Programs\\Transaction Info Sheet(f).pdf', f'{path}\\{property_address}\\Contract-Addenda')

            def admin_email():
                property_address = txt1.get()
                listing_agent = clicked_agents.get()
                
                olApp = win32.Dispatch('Outlook.Application')
                olNS = olApp.GetNameSpace('MAPI')
                mailItem = olApp.CreateItem(0)

                if listing_agent == "Other":
                    agent_name = " "
                    mailItem.CC = " "
                else:
                    agent_name = rcrg_agent_db[listing_agent][1]
                    mailItem.CC = rcrg_agent_db[listing_agent][4]
            
                mailItem.Subject = 'New Listing Request: ' + property_address + f' ({agent_name})'
                mailItem.BodyFormat = 1

                html_body =f"""
                    <p class=MsoNormal>Good {Time}, Amy!<br><br></p>
                    <p class=MsoNormal> {listing_agent} has a listing up and coming for the property located at {property_address}. Would you mind starting the below new listing processes?:<br><br></p>
                    <ol>
                        <li class=MsoListParagraph style='margin-left:0in;mso-list:l0 level1 lfo1'>Create the Incomplete Listing<o:p></o:p></li>
                        <li class=MsoListParagraph style='margin-left:0in;mso-list:l0 level1 lfo1'>Create a Transaction in the Agent&#8217;s TransactionDesk<o:p></o:p></li>
                        <li class=MsoListParagraph style='margin-left:0in;mso-list:l0 level1 lfo1'>Add the Listing to the &#8220;Coming Soon&#8221; whiteboard<o:p></o:p></li>
                    </ol>
                    <p class=MsoNormal> Thank you, <br><br></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-size:14.0pt;font-family:"Arial",sans-serif;color:#1F3864'>Harrison Goehring</span> </b><o:p></o:p></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif'>Office Manager @ The Rick Cox Realty Group</span> </b><o:p></o:p></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>Phone:</span> </b><span style='font-family:"Arial",sans-serif'>(804)447-2834</span> <o:p></o:p></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>E-mail:</span> </b><a href="mailto:Harrison@RickCoxRealty.com"><span style='font-family:"Arial",sans-serif'>Harrison@RickCoxRealty.com</span> </a><o:p></o:p></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>2913 Fox Chase Lane</span> <o:p></o:p></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><span style='font-family:"Arial",sans-serif;color:#1F3864'>Midlothian, VA 23112</span> <o:p></o:p></p>
                    <p class=MsoNormal style='mso-margin-top-alt:auto;mso-margin-bottom-alt:auto;mso-add-space:auto'><a href="http://www.rickcoxrealty.com/"><b><span style='font-family:"Arial",sans-serif;color:#1F3864'>www.RickCoxRealty.com</span> </b></a><o:p></o:p></p>
                """
                
                mailItem.HTMLBody = html_body
                mailItem.To = 'amy@rickcoxrealty.com'
                mailItem._oleobj_.Invoke(*(64209, 0, 8, 0, olNS.Accounts.Item('harrison@rickcoxrealty.com')))

                mailItem.Display()

            #Execute Button
            submit_button = Button(self, text = 'Submit', command = lambda:[seller_folder(), admin_email()])
            submit_button.grid(column = 3, row = 3)

            close_button = Button(self, text = "Close the Window",
                              command= controller.destroy)
            close_button.grid(column=3, row=4)

# Starts our application
if __name__ == '__main__':
    app = MainFrame()
    app.mainloop()