Sub ScrtExport()
' secureCRT export session xml generator
' - input data is selection of column Hostname,HostIP,RemotePort,Type,Username
' - will generate session file .\Export\Session\scrt-<active-sheet>-<timestamp>.xml
' - all sessions will be under folder which name from active sheet
' - open generated secureCRT session file in notepad for review

    Dim SecureCRTFilePath As String
    Dim SecureCRTContent As String
    Dim SelectedRows As Range
    Dim Row As Range
    Dim CustomerName As String
    Dim YearStr As String
    Dim MonthStr As String
    Dim DayStr As String
    Dim MinuteStr As String

    ' Check if rows are selected
    On Error Resume Next
    Set SelectedRows = Selection
    On Error GoTo 0

    If SelectedRows Is Nothing Then
        MsgBox "No rows are selected.", vbExclamation
        Exit Sub
    End If

    ' Get the customer name from the active sheet's tab name
    CustomerName = ActiveSheet.Name

    ' Initialize SecureCRT content
    SecureCRTContent = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbNewLine
    SecureCRTContent = SecureCRTContent & "<VanDyke version=""3.0"">" & vbNewLine
    SecureCRTContent = SecureCRTContent & "    <key name=""Sessions"">" & vbNewLine
    SecureCRTContent = SecureCRTContent & "        <key name=""" & CustomerName & """>" & vbNewLine

    ' Build the SecureCRT content from selected rows with type validation
    For Each Row In SelectedRows.Rows
        Dim HostName As String
        Dim HostIP As String
        Dim RemotePort As String
        Dim Username As String

        HostName = Row.Cells(1).value
        HostIP = Row.Cells(2).value
        RemotePort = Row.Cells(3).value
        Username = Row.Cells(4).value

        SecureCRTContent = SecureCRTContent & "            <key name=""" & CustomerName & "_" & HostName & """>" & vbNewLine
        SecureCRTContent = SecureCRTContent & "                <dword name=""[SSH2] Port"">" & RemotePort & "</dword>" & vbNewLine
        SecureCRTContent = SecureCRTContent & "                <string name=""Hostname"">" & HostIP & "</string>" & vbNewLine
        SecureCRTContent = SecureCRTContent & "                <string name=""Username"">" & Username & "</string>" & vbNewLine
        SecureCRTContent = SecureCRTContent & "                <dword name=""Scrollback"">50000</dword>" & vbNewLine
        SecureCRTContent = SecureCRTContent & "            </key>" & vbNewLine
    Next Row

    SecureCRTContent = SecureCRTContent & "        </key>" & vbNewLine
    SecureCRTContent = SecureCRTContent & "    </key>" & vbNewLine
    SecureCRTContent = SecureCRTContent & "</VanDyke>"

    ' Generate a unique SecureCRT filename based on the current date and time
    YearStr = Format(Now, "yyyy")
    MonthStr = Format(Now, "mm")
    DayStr = Format(Now, "dd")
    MinuteStr = Format(Now, "nn")

    SecureCRTFilePath = ThisWorkbook.Path & "\Export\Session\scrt-" & CustomerName & "-" & YearStr & "-" & MonthStr & "-" & DayStr & "-" & MinuteStr & ".xml"

    ' Write the SecureCRT content to the file
    Open SecureCRTFilePath For Output As #1
    Print #1, SecureCRTContent
    Close #1

    ' Open the SecureCRT file in Notepad
    Shell "notepad.exe """ & SecureCRTFilePath & """", vbNormalFocus
End Sub

Sub mobaExport()
' mobaXterm export session mxtsessions generator
' - input data is selection of column Hostname,HostIP,RemotePort,Type,Username
' - will generate session file .\Export\Session\\mobaxterm-<active-sheet>-<timestamp>.mxtsessions
' - all sessions will be under folder which name from active sheet
' - open generated mobaXterm session file in notepad for review

    Dim MobaXtermFilePath As String
    Dim MobaXtermContent As String
    Dim SelectedRows As Range
    Dim Row As Range
    Dim CustomerName As String
    Dim YearStr As String
    Dim MonthStr As String
    Dim DayStr As String
    Dim MinuteStr As String

    ' Check if rows are selected
    On Error Resume Next
    Set SelectedRows = Selection
    On Error GoTo 0

    If SelectedRows Is Nothing Then
        MsgBox "No rows are selected.", vbExclamation
        Exit Sub
    End If

    ' Get the customer name from the active sheet's tab name
    CustomerName = ActiveSheet.Name

    ' Initialize MobaXterm content
    MobaXtermContent = "[Bookmarks]" & vbNewLine
    MobaXtermContent = MobaXtermContent & "SubRep=" & CustomerName & vbNewLine
    MobaXtermContent = MobaXtermContent & "ImgNum=41" & vbNewLine

    ' Build the MobaXterm content from selected rows with type validation
    For Each Row In SelectedRows.Rows
        Dim HostName As String
        Dim HostIP As String
        Dim RemotePort As String
        Dim Username As String

        HostName = Row.Cells(1).value
        HostIP = Row.Cells(2).value
        RemotePort = Row.Cells(3).value
        Username = Row.Cells(4).value

        MobaXtermContent = MobaXtermContent & HostName & "= #109#0%" & HostIP & "%" & RemotePort & "%" & Username & "%%-1%-1%%%%%0%0%0%%%-1%0%0%0%%1080%%0%0%1%%0%%%%0%-1%-1%0#MobaFont%10%0%0%-1%15%236,236,236%30,30,30%180,180,192%0%-1%0%%xterm%-1%0%_Std_Colors_0_%80%24%0%1%-1%<none>%%0%0%-1%0#0# #-1" & vbNewLine
    Next Row

    ' Generate a unique MobaXterm filename based on the current date and time
    YearStr = Format(Now, "yyyy")
    MonthStr = Format(Now, "mm")
    DayStr = Format(Now, "dd")
    MinuteStr = Format(Now, "nn")

    MobaXtermFilePath = ThisWorkbook.Path & "\Export\Session\mobaxterm-" & CustomerName & "-" & YearStr & "-" & MonthStr & "-" & DayStr & "-" & MinuteStr & ".mxtsessions"
    
    ' Write the MobaXterm content to the file
    Open MobaXtermFilePath For Output As #1
    Print #1, MobaXtermContent
    Close #1

    ' Open the MobaXterm file in Notepad
    Shell "notepad.exe """ & MobaXtermFilePath & """", vbNormalFocus
End Sub
