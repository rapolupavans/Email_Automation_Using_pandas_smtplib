import pandas as pd
import smtplib

#my details
My_name = "Test Name"   # Replace Test name with the name you like
My_email = "Testmail@email.com"  # replace Testmail@email.com with your mail
My_pw = "TestPassword" # Replace Test Password with your mail 

#setting SMTP server
s = smtplib.SMTP('Outgoing Mail (SMTP) Server', 587) # Replace Outgoing Mail (SMTP) Server with smtp.gmail.com(For gmail) or
                                                                                           #smtp.office365.com (for outlook)
s.starttls()
s.login(My_email,My_pw)

#read the file
email_list = pd.read_excel("Test_File_Name.xls") # Replace File name with you file that to be send

#get all the names and scores
all_names = email_list['Name']
all_mails = email_list['Email']
all_scores = email_list['score']
all_tabs = email_list['Tab Switches']
all_links = email_list['Report access URL']
#loop for sending mails

for mail in range(len(all_mails)):
    #get individual records
    name = all_names[mail]
    email = all_mails[mail]
    score = all_scores[mail]
    tab = all_tabs[mail]
    link = all_links[mail]

    full_email = (
"From: {0} <{1}>\n"
"To: {2} <{3}>\n"
"Subject: Test Scores \n\n"
"""
Hi Mr/Ms {2},

Thank you for actively participating in Test CONTEST and making it successful.

We hope you had fun and learnt something while participating in the contest.

You have genuinely participated in the contest with {4} tab switches and scored {5}.

Check your performance here: {6}

If you have any queries(related to contest,questions or any-other related to the event)mail back to us and we will respond to it as soon as possible.

    Thank you
    Test TEAM
    STAY HOME STAY SAFE
""".format(My_name,My_email,name,email,tab,score,link))

# in full_email you can change as per your requirement.
    
    #attempting mails
    
    try:
        s.sendmail(My_email,email,full_email)
        print("Email to {0} successfully sent!\n\n".format(email))
    except Exception as e:
        print("Email to {0} could not be sent :( because {1}\n\n".format(email,str(e)))
    
#close the server
s.close()

        
