import smtplib
import mail_config
import time

import gspread

##access sheets
gc = gspread.service_account(filename= "/home/suryathanush/inventech/api_credentials.json")
sheet1 = gc.open_by_key('1nqFm5ROoHK-_DcoV6IBRiC8SgpBWGQCo4S7jWkUNOUw').worksheet("response form")
sheet2 = gc.open_by_key('1nqFm5ROoHK-_DcoV6IBRiC8SgpBWGQCo4S7jWkUNOUw').worksheet("reward form")

##acess gmail-smtp
def send_email(to,subject, msg):
    try:
        server = smtplib.SMTP('smtp.gmail.com:587')
        server.ehlo()
        server.starttls()
        server.login(mail_config.EMAIL_ADDRESS, mail_config.PASSWORD)
        message = 'Subject: {}\n\n{}'.format(subject, msg)
        server.sendmail(mail_config.EMAIL_ADDRESS, to , message)
        server.quit()
        print("Success: Email sent to:", to)
    except:
        print("Email failed to send to:", to)

#------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                    # code for sending acknowledge mail #

subject_ack = "Thank you for your response"
with open('/home/suryathanush/inventech/response_ack_content.txt', 'r') as f:
    msg_ack = f.read()
  
  ##take row value from until wwhere responses are acknowledged (saves time)##
with open('/home/suryathanush/inventech/serial_no.txt', 'r') as f:
    row = int(f.read())
while(True):
         ##--------if in a row email is present but ackowledge isnt present, then send acknowledge mail-------##
    if(not(len(sheet1.cell(row,2).value)==0))and(len(sheet1.cell(row,6).value) == 0):

        to = sheet1.cell(row,2).value
        send_email(to,subject_ack,msg_ack)
        sheet1.update_cell(row,6, "res_acknowledged")     # fill mail acknowledgement cell of that row with "res-acknowledged"
    
        ##-------if in a row email is not present ,stop while loop----------------##
    if len(sheet1.cell(row,2).value)==0:
        print("responses acknowledged upto date")
        break
    else:
        pass        
    row+=1

## -----------save the rows until where ackowledged------------##
with open('/home/suryathanush/inventech/serial_no.txt', 'w') as f:
    f.write(str(row))


time.sleep(10)
#------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                     # code for sending reward mail for selected #

subject_reward="Congratulations!! your response at inventech was selected for the reward"
with open('/home/suryathanush/inventech/reward_ack_content.txt', 'r') as f:
    msg_reward = f.read()
loop =2          # since first row is heading
while(True):
                ##--------if reward was present but reward_ack was empty, send reward mail----------------##
    if (not(len(sheet1.cell(loop,7).value)==0))and(len(sheet1.cell(loop,8).value) == 0):

        to = sheet1.cell(loop,2).value
        send_email(to,subject_reward,msg_reward)
        sheet1.update_cell(loop,8, "rev_acknowledged")    #after sending mail, fill reward-ack cell of that row with "rev_acknowwledged"
            
            ##--------if in a rrow email is not present ,stop while loop
    if len(sheet1.cell(loop,2).value)==0:
        print("rewards acknowledged upto date")
        break
    else:
        pass
    
    loop+=1         


time.sleep(10)
#--------------------------------------------------------------------------------------------------------------------------------------------------------------#
                                           # code for updating sheet1 with phone munbers from sheet2 #

i=3
while(True):
    j=2
    if len(sheet2.cell(i,2).value) ==0:
        print("phone numbers updated upto date")   # if email was empty in a row of sheet2, stop there #
        break
    while(True):

                    # if emails in sheet1 and sheet2 of particularr rows are equal, and reward was present but payment not yet done #
        if (sheet2.cell(i,2).value == sheet1.cell(j,2).value)and(not(len(sheet1.cell(j,7).value) == 0))and(len(sheet1.cell(j,10).value) == 0):
            if not sheet2.cell(i,3).value == sheet1.cell(j,9).value :
                sheet1.update_cell(j,9,sheet2.cell(i,3).value)
                print("phone number of", sheet1.cell(j,2).value ,"is updated")
            else:
                print("@@@@@@@@@@")    
        
        if len(sheet1.cell(j,6).value) == 0:
            break

        else:
            pass

        j+=1
    i+=1

    
#--------------------------------------------------------------------------------------------------------------------------------------------------------------#