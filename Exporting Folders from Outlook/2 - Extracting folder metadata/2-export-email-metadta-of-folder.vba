Option Explicit

Dim StrSavePath As String

Sub ExtractEmailInfoToExcel_savedtoDesktop()

    Dim i As Long
    Dim j As Long
    Dim StrFolder As String
    Dim iNameSpace As NameSpace
    Dim myOlApp As Outlook.Application
    Dim ChosenFolder As MAPIFolder
    Dim mItem As Object ' Generic Object to handle multiple item types
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim DesktopPath As String
    
    ' Set up Outlook and Excel objects
    On Error GoTo ErrorHandler
    Set myOlApp = Outlook.Application
    Set iNameSpace = myOlApp.GetNamespace("MAPI")
    Set ChosenFolder = iNameSpace.PickFolder
    If ChosenFolder Is Nothing Then GoTo ExitSub

    ' Set Desktop Path for saving the Excel file
    DesktopPath = Environ("USERPROFILE") & "\Desktop\"

    ' Initialize Excel Application
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False ' Keep Excel hidden during the process
    
    ' Create a new workbook and set the sheet
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.Name = "Email Log"

    ' Set headers for the Excel file
    xlSheet.Cells(1, 1).Value = "Subject"
    xlSheet.Cells(1, 2).Value = "Received Date"
    xlSheet.Cells(1, 3).Value = "Sender Name"
    
    ' Set folder name as the file name
    StrFolder = StripIllegalChar(ChosenFolder.Name)

    ' Loop through all emails in the selected folder
    For j = 1 To ChosenFolder.Items.Count
        Set mItem = ChosenFolder.Items(j)
        
        ' Only process Mail Items
        If TypeOf mItem Is MailItem Then
            ' Add email details to the Excel sheet
            xlSheet.Cells(j + 1, 1).Value = mItem.Subject
            xlSheet.Cells(j + 1, 2).Value = mItem.ReceivedTime
            xlSheet.Cells(j + 1, 3).Value = mItem.SenderName
        End If
    Next j

    ' Save the Excel workbook with the folder's name on the Desktop
    xlBook.SaveAs DesktopPath & StrFolder & ".xlsx"
    xlBook.Close False
    xlApp.Quit

    GoTo ExitSub

ExitSub:
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

' Function to strip illegal characters from filenames
Function StripIllegalChar(StrInput As String) As String
    Dim RegX As Object
    Set RegX = CreateObject("vbscript.regexp")
    
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
    
    StripIllegalChar = RegX.Replace(StrInput, "")
End Function
