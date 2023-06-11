import gspread
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

def send_email(sender_email, sender_password, receiver_email, subject, message):
    # Create a multipart message
    email_message = MIMEMultipart()
    email_message["From"] = sender_email
    email_message["To"] = receiver_email
    email_message["Subject"] = subject

    # Add the message body
    email_message.attach(MIMEText(message, "plain"))

    # Connect to the SMTP server
    smtp_server = smtplib.SMTP("smtp.gmail.com", 587)  # Replace with your SMTP server and port
    smtp_server.starttls()
    smtp_server.login(sender_email, sender_password)

    # Send the email
    smtp_server.sendmail(sender_email, receiver_email, email_message.as_string())

    # Close the SMTP server connection
    smtp_server.quit()

def genmessage(worksheet):
    message=[]
    rows=worksheet.get_all_records()
    for x in range(len(rows)):
        if rows[x]['Token Number']==0:
            message.append([rows[x]['Email Address'],"You have already filled out this form. Kindly check your inbox for a previous mail. Try filling out with a different email address if you think this is a mistake."])
        elif rows[x]['Token Number']==-1:
            message.append([rows[x]['Email Address'],"So sorry, we were unable to get an appointment for you today. Kindly try again tomorrow."])
        else:
            if rows[x]['Age']<18:
                message.append([rows[x]['Email Address'],"Hello "+rows[x]['Name']+", you have been assigned the token number "+str(rows[x]['Token Number'])+" at our "+rows[x]['Location']+" clinic. We have detected that you are under 18 years of age, please be accompanied by a legal guardian. Kindly be on time for your appointment."])
            elif rows[x]['Age']<60:
                message.append([rows[x]['Email Address'],"Hello "+rows[x]['Name']+", you have been assigned the token number "+str(rows[x]['Token Number'])+" at our "+rows[x]['Location']+" clinic. Congratulations! You are eligible for a senior citizen discount. Please check with the reception for more details. Kindly be on time for your appointment."])     
            else:
                message.append([rows[x]['Email Address'],"Hello "+rows[x]['Name']+", you have been assigned the token number "+str(rows[x]['Token Number'])+" at our "+rows[x]['Location']+" clinic. Please carry a valid government ID card as proof of age. Kindly be on time for your appointment."])
    return message        

def sendmessage(message):
    sender_email = "tokenizerbotcs@gmail.com"
    sender_password = "awvhbhlmdndyparf "
    #receiver_email = "arnavg1808@gmail.com"
    subject = "Hello from clinic!"
    for x in range(len(message)):
        send_email(sender_email, sender_password, message[x][0],subject,message[x][1])

def initialize():
    SHEET_URL = 'https://docs.google.com/spreadsheets/d/16OCCYVb8J6uGF6t1GXo0IHV1P57_r_8tpJccRHVZVUA/edit#gid=938282913'
    gc = gspread.service_account(filename='cs104-project-389321-cf3321a3f981.json')
    spreadsheet = gc.open_by_url(SHEET_URL)
    worksheet = spreadsheet.get_worksheet(0)
    return worksheet

def tokenizer(worksheet):
    rows = worksheet.get_all_records()
    worksheet.update('G1','Token Number')
    l=['Delhi','Mumbai','Chennai','Kolkata','Bhopal']
    female_token=[1,1,1,1,1]
    male_token=[1,1,1,1,1]
    emaillist=[]
    for x in range(len(rows)):
        pointer=l.find(rows[x]['Location'])
        if rows[x]['Sex']=='Male':
            if rows[x]['Email Address'] not in emaillist:
                if male_token[pointer]<=6:
                    worksheet.update_cell(x+2,7,male_token[pointer])
                else:
                    worksheet.update_cell(x+2,7,-1)
                male_token[pointer]+=1
                emaillist.append(rows[x]['Email Address'])
            else:
                worksheet.update_cell(x+2,7,0)
        else:
            if rows[x]['Email Address'] not in emaillist:
                if female_token[pointer]<=6:
                    worksheet.update_cell(x+2,7,female_token[pointer])
                else:
                    worksheet.update_cell(x+2,7,-1)
                female_token[pointer]+=1
                emaillist.append(rows[x]['Email Address'])
            else:
                worksheet.update_cell(x+2,7,0)

if __name__=="__main__":
    worksheet = initialize()
    tokenizer(worksheet)
    message = genmessage(worksheet)
    sendmessage(message)
    #Send the email
    #send_email(sender_email, sender_password, receiver_email, subject, message)


    print('==============================')