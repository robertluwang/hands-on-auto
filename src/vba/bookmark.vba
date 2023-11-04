Sub Bookmark()
' bookmark generator
' - input data from selection of column name and url
' - generate bookmark file as .\Export\Bookmark\bookmark-<active-sheet>-<timestamp>.html
' - open bookmark file in notebook for review

    Dim BookmarkFilePath As String
    Dim BookmarkContent As String
    Dim SelectedRows As Range
    Dim Row As Range
    Dim YearStr As String
    Dim MonthStr As String
    Dim DayStr As String
    Dim MinuteStr As String
    Dim FolderName As String
    
    ' Check if rows are selected
    On Error Resume Next
    Set SelectedRows = Selection
    On Error GoTo 0
    
    If SelectedRows Is Nothing Then
        MsgBox "No rows are selected.", vbExclamation
        Exit Sub
    End If
    
    ' Get the folder name from the active sheet's tab name
    FolderName = ActiveSheet.Name
    
    ' Initialize bookmark content
    BookmarkContent = "<!DOCTYPE NETSCAPE-Bookmark-file-1>" & vbNewLine
    BookmarkContent = BookmarkContent & "<META HTTP-EQUIV=""Content-Type"" CONTENT=""text/html; charset=UTF-8"">" & vbNewLine
    BookmarkContent = BookmarkContent & "<TITLE>Bookmarks</TITLE>" & vbNewLine
    BookmarkContent = BookmarkContent & "<H1>Bookmarks</H1>" & vbNewLine
    BookmarkContent = BookmarkContent & "<DL><p>" & vbNewLine
    BookmarkContent = BookmarkContent & "    <DT><H3 ADD_DATE=""1635453051"" LAST_MODIFIED=""1635453051"">" & FolderName & "</H3>" & vbNewLine
    BookmarkContent = BookmarkContent & "    <DL><p>" & vbNewLine
    
    ' Build the bookmark content from selected rows with URL validation
    For Each Row In SelectedRows.Rows
        Dim Name As String
        Dim URL As String
        
        Name = Row.Cells(1).value
        URL = Row.Cells(2).value
        
        ' Validate the URL
        If Not (Left(URL, 4) = "http" Or Left(URL, 5) = "https") Then
            MsgBox "URL should start with http or https!", vbExclamation
            Exit Sub
        End If
        
        BookmarkContent = BookmarkContent & "        <DT><A HREF=""" & URL & """>" & Name & "</A>" & vbNewLine
    Next Row
    
    BookmarkContent = BookmarkContent & "    </DL><p>" & vbNewLine
    BookmarkContent = BookmarkContent & "</DL><p>"
    
    ' Generate a unique bookmark filename based on the current date and time
    YearStr = Format(Now, "yyyy")
    MonthStr = Format(Now, "mm")
    DayStr = Format(Now, "dd")
    MinuteStr = Format(Now, "nn")
    
    BookmarkFilePath = ThisWorkbook.Path & "\Export\Bookmark\bookmark-" & FolderName & "-" & YearStr & "-" & MonthStr & "-" & DayStr & "-" & MinuteStr & ".html"
    
    ' Write the bookmark content to the file
    Open BookmarkFilePath For Output As #1
    Print #1, BookmarkContent
    Close #1
    
    ' Open the bookmark file in Notepad
    Shell "notepad.exe """ & BookmarkFilePath & """", vbNormalFocus
End Sub
