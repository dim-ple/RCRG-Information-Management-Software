def selling_agent_email():
        
        property_address = txt1.get()
        listing_agent = clicked_agents.get()
        attorney_contact = clicked_attorneys.get()
        selling_agent = txt.get()

        
        olApp = win32.Dispatch('Outlook.Application')
        olNS = olApp.GetNameSpace('MAPI')

        mailItem = olApp.CreateItem(0)
        mailItem.Subject = 'New Seller-Side Transaction - ' + property_address
        mailItem.BodyFormat = 1
        mailItem.HTMLBody = 'Attorney E-mail'

        #To: Operating Logic - Dictionary Call
        selling_agent_name = " " if (selling_agent == "Other") else agent_db[selling_agent][2]
        

        #CC: Operating Logic - Dictionary Call
        listing_agent_name = " " if (listing_agent == "Other") else rcrg_agent_db[listing_agent][1] + " " + rcrg_agent_db[listing_agent][2]
        mailItem.CC = " " if (listing_agent == "Other") else rcrg_agent_db[listing_agent][4] + "; amy@rickcoxrealty.com;"



        html_body =f"""
            <p class=MsoNormal>Good {Time}, {selling_agent_name}!<br><br></p>
            <p class=MsoNormal>My name is Harrison Goehring and I am the Office Manager for the Rick Cox Realty Group. I will be assisting {listing_agent_name} and their client with the sale of {property_address}.<br><br></p>
            <p class=MsoNormal>{attorney_msg}<br><br></p>
            <p class=MsoNormal>CC: {listing_agent_name}, Listing Agent; Amy Foldes, Team Administrator<br><br></p>
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