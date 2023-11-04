# Expiration Reminder 

A collection of VBA macro in excel, to generate expiration email list from outlook and send expiration reminder email from outlook.

Assume you want to record expiration email in excel and send reminder email, here is quick solution for it using VBA macro in excel.

## How it works

- create outlook rule to move all incoming emails from specific sender with keyword like expire, expired or expiration etc to your own folder, for example Dev\reminder in this case
- run macro GenExpirationEmailList to generate expiration email list in new sheet <active-sheetname>-expiration, will include sender, date and email body in descending order
- build up expiration table in new sheet to inlcude column 'Customer' and 'Expiration Date'(yyyy-mm-dd) manually or macro
- run maco SendReminderEmail to send reminder email from outlook, to remind action needed if expiration date less than 7 days, otherwise send closet expiration date enties for reference
- you can run macro from Developer tab or create customzied ribbon in excel 

## VBA macro list 

**Sub GenExpirationList()**

generate expiration list in excel 
- it scans email subject with keyword Expire/Expired/Expiration on all emails under outlook folder Dev\reminder, you can change keyword and subfolder location 
- generate new sheet <active-sheetname>-expiration to include scan result - sender, date and email body in descending order

**Sub SendReminderEmail()**

send expiration reminder email from outlook 
- scan record from active sheet with table column 'Customer' and 'Expiration Date'
- reminder email with detail if expiration date less than 7 days or closet expiration date entries 

Private Function GetInfoFromEmailBody(emailBody As String, infoLabel As String) As String

get info from outlook email body with keyword

**Sub FormatActiveSheet()**

format whole range of visiable cells on active sheet 
- font: Calibri size 11
- color: from color picker 
- cell border: solid black line 
- auto fit cell content
  
Private Function ChooseColor() As Long

color picker
- pop up color picker 
- return selected color code

