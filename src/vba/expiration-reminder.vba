Sub SendReminderEmail()
' send expiration reminder email from outlook 
' - scan record from active sheet with table column 'Customer'	'Expiration Date'
' - reminder email with detail if expiration date less than 7 days or closet expiration date entries 

    Dim ws As Worksheet
    Dim OutApp As Object
    Dim OutMail As Object
    Dim lastRow As Long
    Dim i As Long
    Dim ToEmail As String
    Dim subject As String
    Dim body As String
    Dim expirationDate As Date
    Dim TodayDate As Date
    Dim ClosestExpiration As Date
    Dim ClosestCustomers As String
    Dim ExpirationLessThan7Days As String
    
    ' Set the active sheet as the worksheet to extract data from
    Set ws = ActiveSheet
    
    ' Get the last row in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Initialize Outlook
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
    
    ' Set your email address
    ToEmail = "robert@company.com" ' Change to your email address
    
    ' Get today's date
    TodayDate = Date
    
    ' Initialize ClosestExpiration to a distant future date
    ClosestExpiration = Date + 365 ' 1 year from today
    ClosestCustomers = ""
    ExpirationLessThan7Days = ""
    
    ' Loop through the rows in the active sheet starting from the 2nd row
    For i = 2 To lastRow
        ' Extract customer name from column A
        customerName = ws.Cells(i, 1).Value
        
        ' Check if the cell in column B contains a valid date
        If IsDate(ws.Cells(i, 2).Value) Then
            ' Extract expiration date from column B
            expirationDate = CDate(ws.Cells(i, 2).Value)
        Else
            ' Handle cases where the date is not valid
            expirationDate = Date + 365 ' Set to a distant future date or handle it as appropriate
        End If
        
        ' Calculate the time difference between the expiration date and today
        DateDifference = expirationDate - TodayDate
        
        ' Calculate the remaining days until expiration
        remainingDays = DateDifference
        
        ' Check if expiration date is less than 7 days from today
        If DateDifference < 7 Then
            ' Add this entry to ExpirationLessThan7Days
            ExpirationLessThan7Days = ExpirationLessThan7Days & customerName & " " & Format(expirationDate, "yyyy-mm-dd") & " " & remainingDays & " days" & vbCrLf
        ElseIf DateDifference < ClosestExpiration - TodayDate Then
            ' Update ClosestExpiration and ClosestCustomers with the closest expiration date
            ClosestExpiration = expirationDate
            ClosestCustomers = customerName & " " & Format(expirationDate, "yyyy-mm-dd") & " " & remainingDays & " days"
        ElseIf DateDifference = ClosestExpiration - TodayDate Then
            ' Handle multiple customers on the same closest expiration date
            ClosestCustomers = ClosestCustomers & vbCrLf & customerName & " " & Format(expirationDate, "yyyy-mm-dd") & " " & remainingDays & " days"
        End If
    Next i
    
    ' Get the active sheet name
    ActiveSheetName = ws.Name
    
    ' Compose email subject
    subject = ActiveSheetName & " Reminder: Upcoming Expirations"
    
    ' Compose email body with a table
    body = ""
    
    ' Check if there are entries with expiration dates less than 7 days
    If ExpirationLessThan7Days <> "" Then
        body = "Entries with Expiration Dates Less Than 7 Days:" & vbCrLf & vbCrLf & "Customer Expiration-Date Remaining-Days" & vbCrLf
        body = body & ExpirationLessThan7Days & vbCrLf
        ' Add the "Please take action ASAP!" message at the end of the email
        body = body & vbCrLf & "Please take action ASAP!"
    Else
        ' If there are no entries with expiration dates less than 7 days, bring attention to the closest expiration date
        body = "No entries with expiration dates less than 7 days." & vbCrLf & vbCrLf
        If ClosestCustomers <> "" Then
            body = body & "Closest Expiration Date:" & vbCrLf & vbCrLf & "Customer Expiration-Date Remaining-Days" & vbCrLf
            body = body & ClosestCustomers & vbCrLf
        End If
    End If
    
    ' Display a message box to verify the email content
    If MsgBox("Verify Email Content" & vbCrLf & vbCrLf & "Subject: " & subject & vbCrLf & vbCrLf & "Body: " & body & vbCrLf & vbCrLf & "Send this email?", vbYesNo) = vbYes Then
        ' Create and send the reminder email
        With OutMail
            .To = ToEmail
            .subject = subject
            .HTMLBody = "<html><body>" & Replace(body, vbCrLf, "<br>") & "</body></html>"
            .Send
        End With
    End If
    
    ' Release Outlook objects
    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub

Sub GenExpirationList()
' generate expiration list in excel 
' - it scans email subject with keyword Expire/Expired/Expiration on all emails under outlook folder Dev\reminder, you can change keyword and subfolder location 
' - generate new sheet <active-sheetname>-expiration to include scan result
    Dim olApp As Object
    Dim olNs As Object
    Dim olFolder As Object
    Dim olItems As Object
    Dim olItem As Object
    Dim iRow As Integer
    Dim subjectText As String
    Dim emailBody As String
    Dim reminderText As String
    Dim senderName As String
    Dim sendDate As Date
    Dim ws As Object
    Dim expSheet As Object

    ' Set Outlook Application and Namespace
    Set olApp = CreateObject("Outlook.Application")
    Set olNs = olApp.GetNamespace("MAPI")

    ' Set the Outlook folder where the emails are located (Dev\reminder)
    Set olFolder = olNs.Folders("robert@company.com").Folders("Dev").Folders("reminder")

    ' Get items in the folder
    Set olItems = olFolder.Items

    ' Set the active sheet as the worksheet to extract data from
    Set ws = ActiveSheet
    
    ' Create a new worksheet named after the active sheet with '-expiration' suffix
    On Error Resume Next
    Set expSheet = ThisWorkbook.Sheets("expirationEmail")
    On Error GoTo 0
    
    ' Delete the existing sheet if it exists
    If Not expSheet Is Nothing Then
        Application.DisplayAlerts = False
        expSheet.Delete
        Application.DisplayAlerts = True
    End If
    
    Set ws = ActiveSheet
    
    ' Create a new worksheet with the specified name
    Set expSheet = ThisWorkbook.Sheets.Add(After:=ws)
    expSheet.Name = "expirationEmail"

    ' Find the last row with data in Excel
    iRow = expSheet.Cells(expSheet.Rows.Count, "A").End(xlUp).Row
    
    ' Add headers to columns A, B, and C
    expSheet.Cells(1, 1).Value = "Sender"
    expSheet.Cells(1, 2).Value = "Date"
    expSheet.Cells(1, 3).Value = "Email"

    ' Set the current date
    today = Date
    
    iRow = 2
    
    ' Loop through Outlook items in the "Dev\reminder" folder
    For Each olItem In olItems
        ' Check the subject for keywords (modify as needed)
        subjectText = olItem.subject
        If InStr(1, subjectText, "Expire", vbTextCompare) > 0 Or InStr(1, subjectText, "Expired", vbTextCompare) > 0 Or InStr(1, subjectText, "Expiration", vbTextCompare) > 0 Then
            
            ' Extract customer and expiration information from the email body
            emailBody = olItem.body

            senderName = Trim(olItem.senderName)
            
            sendDate = Format(Left(Trim(olItem.ReceivedTime), 10), "yyyy-MM-dd")
            
          
            ' Insert the data into Excel
            expSheet.Cells(iRow, 1).Value = senderName
            expSheet.Cells(iRow, 2).Value = sendDate
            expSheet.Cells(iRow, 3).Value = emailBody

            expSheet.Columns("C").ColumnWidth = 80
            expSheet.Columns("A:C").AutoFit

            ' Increase the row counter
            iRow = iRow + 1
    
            ' Construct the reminder text
            'reminderText = "Email Sender: " & senderName & vbCrLf & _
                           "Date: " & sendDate & vbCrLf & _
                           "Original email: " & vbCrLf & emailBody & vbCrLf
    
            ' Send a reminder email to yourself
            'MsgBox "Subject: Expiration email Reminder" & vbCrLf & reminderText & vbCrLf
            'SendReminderEmail "robert@company.com", "Expiration email Reminder", reminderText
        End If
    Next olItem
    
    ' Sort the data on column B ('Date') in descending order
    expSheet.Range("A1:C" & iRow).Sort Key1:=expSheet.Range("B2:B" & iRow), Order1:=xlDescending, Header:=xlYes

    ' Release the objects
    Set expSheet = Nothing
    Set olItem = Nothing
    Set olItems = Nothing
    Set olFolder = Nothing
    Set olNs = Nothing
    Set olApp = Nothing
End Sub

Private Function GetInfoFromEmailBody(emailBody As String, infoLabel As String) As String
' get info from outlook email body with keyword
                              
    Dim lines() As String
    Dim line As String
    Dim info As String
    Dim i As Integer

    ' Split the email body into lines
    lines = Split(emailBody, vbCrLf)

    ' Initialize the counter
    i = 0

    ' Search for the information label in each line
    Do While i < UBound(lines)
        line = lines(i)
        If InStr(1, line, infoLabel, vbTextCompare) > 0 Then
            ' Extract the information following the label
            info = Trim(Mid(line, Len(infoLabel) + 1))
            Exit Do
        End If
        i = i + 1
    Loop

    GetInfoFromEmailBody = info
End Function

Sub FormatActiveSheet()
' format whole range of visiable cells on active sheet 
' - font: Calibri size 11 
' - color: from color picker 
' - cell border: solid black line 
' - auto fit cell content
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    Dim myRange As Range
    Dim colorRGB As Long

    ' Set the active sheet
    Set ws = ActiveSheet

    ' Call the color picker to get a new color
    colorRGB = ChooseColor()

    ' Find the last row and last column with data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastColumn = ws.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

    ' Define the range based on the last row and last column with data
    Set myRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn))

    ' Apply formatting to the range
    With myRange
        ' Set font properties
        .Font.Name = "Calibri"
        .Font.Size = 11

        ' Set cell color to the selected color
        .Interior.color = colorRGB

        ' Set cell borders
        .Borders.LineStyle = xlContinuous ' Solid line
        .Borders.color = RGB(0, 0, 0) ' Border color (adjust RGB values)

        ' Autofit the selected range to wrap text
        .WrapText = True
    End With
End Sub

Private Function ChooseColor() As Long
' color picker 
' - pop up color picker 
' - return selected color code
                            
    ' Create variables for the color codes
    Dim FullColorCode As Long

    ' Open the ColorPicker dialog box and get the selected color
    If Application.Dialogs(xlDialogEditColor).Show(1) = True Then
        FullColorCode = ActiveWorkbook.Colors(1)
    Else
        ' Default to Grey if the user cancels
        FullColorCode = RGB(192, 192, 192) ' Grey color
    End If

    ' Return the selected color code
    ChooseColor = FullColorCode
End Function
