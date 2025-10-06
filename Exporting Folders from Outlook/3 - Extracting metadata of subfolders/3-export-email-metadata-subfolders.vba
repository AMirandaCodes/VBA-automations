Option Explicit

Dim StrSavePath As String

Sub ExtractEmailInfoFromSubfolders()
    Dim iNameSpace As NameSpace
    Dim myOlApp As Outlook.Application
    Dim ChosenFolder As MAPIFolder
    Dim SubFolder As MAPIFolder
    Dim DesktopPath As String
    
    ' Set up Outlook objects
    On Error GoTo ErrorHandler
    Set myOlApp = Outlook.Application
    Set iNameSpace = myOlApp.GetNamespace("MAPI")
    Set ChosenFolder = iNameSpace.PickFolder
    If ChosenFolder Is Nothing Then GoTo ExitSub

    ' Set Desktop path and create "Email Info" folder
    DesktopPath = Environ("USERPROFILE") & "\Desktop\Email Info\"
    If Dir(DesktopPath, vbDirectory) = "" Then
        MkDir DesktopPath
    End If

    ' Loop through each subfolder and export email details
    For Each SubFolder In ChosenFolder.Folders
        ExportEmailsToExcel SubFolder, DesktopPath
    Next SubFolder

    MsgBox "Export completed successfully.", vbInformation

    GoTo ExitSub

ExitSub:
    Set myOlApp = Nothing
    Set iNameSpace = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Sub ExportEmailsToExcel(ByVal Folder As MAPIFolder, ByVal SavePath As String)
    Dim mItem As Object
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim j As Long
    Dim FolderName As String

    ' Set up Excel objects
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Add
    Set xlSheet = xlBook.Sheets(1)
    xlSheet.Name = "Email Log"

    ' Set headers for the Excel sheet
    xlSheet.Cells(1, 1).Value = "Subject"
    xlSheet.Cells(1, 2).Value = "Received Date"
    xlSheet.Cells(1, 3).Value = "Sender Name"

    ' Clean up folder name for file use
    FolderName = StripIllegalChar(Folder.Name)

    ' Loop through emails in the subfolder
    j = 1
    For Each mItem In Folder.Items
        If TypeOf mItem Is MailItem Then
            xlSheet.Cells(j + 1, 1).Value = mItem.Subject
            xlSheet.Cells(j + 1, 2).Value = mItem.ReceivedTime
            xlSheet.Cells(j + 1, 3).Value = mItem.SenderName
            j = j + 1
        End If
    Next mItem

    ' Save Excel workbook with subfolder's name
    xlBook.SaveAs SavePath & FolderName & ".xlsx"
    xlBook.Close False
    xlApp.Quit

    ' Release Excel objects
    Set xlApp = Nothing
    Set xlBook = Nothing
    Set xlSheet = Nothing
End Sub

Function StripIllegalChar(StrInput As String) As String
    Dim RegX As Object
    Set RegX = CreateObject("vbscript.regexp")
    
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
    
    StripIllegalChar = RegX.Replace(StrInput, "")
End Function
