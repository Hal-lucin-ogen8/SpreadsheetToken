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
    email_message.attach(MIMEText(message, "html"))

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
            message.append([rows[x]['Email Address'],"You have already filled out this form. Kindly check your inbox for a previous mail. Try filling out the <a href='https://forms.gle/w6ZDp3KAKvdmpdnMA'>form</a> with a different email address if you think this is a mistake.<br/>Making lives easier,<br/>Hope Clinic"])
        elif rows[x]['Appointment']=='Not assigned':
            message.append([rows[x]['Email Address'],"So sorry " +rows[x]['Name']+", we were unable to get an appointment for you today. Kindly try again tomorrow.<br/>Hope Clinic"])
        else:
            if rows[x]['Age']<18:
                message.append([rows[x]['Email Address'],"Hello "+rows[x]['Name']+", you have been assigned the token number <b>"+str(rows[x]['Token Number'])+"</b> at our <i>"+rows[x]['Location']+"</i> clinic. Your appointment is scheduled for <b>"+str(rows[x]['Appointment'])+"</b>. We have detected that you are under 18 years of age, please be accompanied by a legal guardian. Kindly be on time for your appointment.<br/>Making lives easier,<br/>Hope Clinic"])
            elif rows[x]['Age']<60:
                message.append([rows[x]['Email Address'],"Hello "+rows[x]['Name']+", you have been assigned the token number <b>"+str(rows[x]['Token Number'])+"</b> at our <i>"+rows[x]['Location']+"</i> clinic. Your appointment is scheduled for <b>"+str(rows[x]['Appointment'])+"</b>. Congratulations! You are eligible for a senior citizen discount. Please carry a valid ID card as age proof. Please check with the reception for more details. Kindly be on time for your appointment.<br/>Making lives easier,<br/>Hope Clinic"])
            else:
                message.append([rows[x]['Email Address'],"Hello "+rows[x]['Name']+", you have been assigned the token number <b>"+str(rows[x]['Token Number'])+"</b> at our <i>"+rows[x]['Location']+"</i> clinic. Your appointment is scheduled for <b>"+str(rows[x]['Appointment'])+"</b>. Please carry a valid government ID card as proof of age. Kindly be on time for your appointment.<br/>Making lives easier,<br/>Hope Clinic"])
    return message        

def sendmessage(message):
    sender_email = "tokenizerbotcs@gmail.com"
    sender_password = "awvhbhlmdndyparf "
    subject = "Regarding your upcoming appointment at Hope Clinic"
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
    worksheet.update('H1','Token Number')
    l=['Delhi','Mumbai','Chennai','Kolkata','Bhopal']
    female_token=[1,1,1,1,1]
    male_token=[1,1,1,1,1]
    emaillist=[]
    for x in range(len(rows)):
        pointer=l.index(rows[x]['Location'])
        if rows[x]['Sex']=='Male':
            if rows[x]['Email Address'] not in emaillist:
                worksheet.update_cell(x+2,8,male_token[pointer])
                male_token[pointer]+=1
                emaillist.append(rows[x]['Email Address'])
            else:
                worksheet.update_cell(x+2,8,0)
        else:
            if rows[x]['Email Address'] not in emaillist:
                worksheet.update_cell(x+2,8,female_token[pointer])
                female_token[pointer]+=1
                emaillist.append(rows[x]['Email Address'])
            else:
                worksheet.update_cell(x+2,8,0)
    allocation(worksheet)

def allocation(worksheet):
    l=['Delhi','Mumbai','Chennai','Kolkata','Bhopal']
    rows=worksheet.get_all_records()
    worksheet.update('I1','Appointment')
    testlist = [{'10AM':[[],[]],'10:30AM':[[],[]],'11AM':[[],[]],'11:30AM':[[],[]],'12PM':[[],[]],'12:30PM':[[],[]],'1PM':[[],[]]},
                {'10AM':[[],[]],'10:30AM':[[],[]],'11AM':[[],[]],'11:30AM':[[],[]],'12PM':[[],[]],'12:30PM':[[],[]],'1PM':[[],[]]},
                {'10AM':[[],[]],'10:30AM':[[],[]],'11AM':[[],[]],'11:30AM':[[],[]],'12PM':[[],[]],'12:30PM':[[],[]],'1PM':[[],[]]},
                {'10AM':[[],[]],'10:30AM':[[],[]],'11AM':[[],[]],'11:30AM':[[],[]],'12PM':[[],[]],'12:30PM':[[],[]],'1PM':[[],[]]},
                {'10AM':[[],[]],'10:30AM':[[],[]],'11AM':[[],[]],'11:30AM':[[],[]],'12PM':[[],[]],'12:30PM':[[],[]],'1PM':[[],[]]},]
    for x in range(len(rows)):
        pointer=l.index(rows[x]['Location'])
        string1 = rows[x]['Time preference']
        if rows[x]['Token Number']!=0:
            if rows[x]['Sex']=='Male':
                list2 = string1.split(", ")
                for k in list2:
                    if len(testlist[pointer][k][0])==0:
                        testlist[pointer][k][0].append(rows[x]['Token Number'])
                        worksheet.update_cell(x+2,9,k)
                        break
                else:
                    worksheet.update_cell(x+2,9,'Not assigned')
            else:
                list2 = string1.split(", ")
                for k in list2:
                    if len(testlist[pointer][k][1])==0:
                        testlist[pointer][k][1].append(rows[x]['Token Number'])
                        worksheet.update_cell(x+2,9,k)
                        break
                else:
                    worksheet.update_cell(x+2,9,'Not assigned')

if __name__=="__main__":
    worksheet = initialize()
    tokenizer(worksheet)
    print("Appointments and tokens generated.")
    message = genmessage(worksheet)
    sendmessage(message)
    print('Intimation mails successfully sent.')