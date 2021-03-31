# Birthday Wisher using Python
# Send birthday wishes to the person automatically on whatsapp or you can send email to them.
# install all imports.
# For sending mail, allow from your gamil account less secure app to login (Turn it on), later you can turn it off.
# Fill your data in the excel sheet wishes.xlsx.
# You can schedule this by using task schedular in windows 10 and for linux cron job.
# All imports
import pandas as pd
import datetime
import smtplib
import pywhatkit as pw
import os

# provide the directory, so after scheduling there must be no error occur in reading the data.
os.chdir("directory-of -project")

# install pywhatkit module for it
#pw.sendwhatmsg("mobile number",message,time_starting,time_end).
#pywhatkit use 24 hour time format, therefore ex: 18:20 == 6:20.
#Number must be save in your contact and your  whatsup must be open on web.whatsapp.com.
# It will open web.whatsapp.com and search for that number and send your message automatically. 
#below code will send message to the whatsup you provide, with message, at 12:30.
def sendWhatsappMessages(number,message,time_hour,time_minute):
    pw.sendwhatmsg(number,message,time_hour,time_minute)

#You can also send message using gmail.
# install smtplib module for it.
#provide your gmail and password in below variables.
Gmail_id="your email"
Gmail_pass="your email password"

#this function will send email to all email you provided.
def sendEmail(to,subject,msg):
    #it will open gmail server
    s=smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    #will login using your id and password.
    s.login(Gmail_id,Gmail_pass)
    #will send email to the person you selected
    s.sendmail(Gmail_id,to,f" {subject} \n\n {msg}")
    #after sending mail , it will end the session.
    s.quit()

#starting main 
if __name__=="__main__":
    #Reading the excel file you provided.
    df=pd.read_excel("wishes.xlsx")
    #getting today date and month using datetime module.
    #for it install datetime module.
    today= datetime.datetime.now().strftime("%d-%m")
    #getting year using datetime module.
    yearNow=datetime.datetime.now().strftime("%Y")
    
    # list created to update the excel sheet, we doesnot wish again the  same person in a .
    writeInd=[]

    #for loop for reading content from excel sheet which is saved in df variable.
    #iterrows() builten function after reading the excel sheet, it will iterate the rows from data frame.
    for index, item in df.iterrows():
        #selecting birthday date from sheet.
        birthday=item['Birthday'].strftime("%d-%m")
        
        #Now after getting the date checking that today date and birthday of that person is same or not.
        # Second condition => year(current) we select up there must not be in the column of year, if it is mena we already wished him/her.
        if today==birthday and yearNow not in str(item['Year']):
            #Inside, we are calling the function, created above,
            # send whatsupp message
            sendWhatsappMessages(item['Whatsapp'],"message",00,00)
            #                           OR
            # Send email
            sendEmail(item['Email'],"Subject",item['Dialogue'])
            #After sending message we are adding the row-index of that in our list.
            writeInd.append(index)
    
    
    #Here we are saving our time, by using if/else
    # for updating the excel sheet, first we are checking that list is empty or not.
    # if it is, mean no birthday message send, means no update.
    if len(writeInd)==0:
        print("empty")
    else:
    #Here we are updating the excel sheet.
         for i in writeInd:
            #getting the year written on excel sheet of that birthday folk.
            yr=df.loc[i,"Year"]
            #updating it by adding current year because we already wish them.
            df.loc[i,"Year"]=str(yr)+"," +str(yearNow)
        #saving the excel sheet with the same name and no indexing agin is required.
         df.to_excel("wishes.xlsx",index=False)