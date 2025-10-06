Option Explicit
Dim StrSavePath As String

Sub Export_AllEmails_from_Outlook()

    Dim i As Long
    Dim j As Long
    Dim n As Long
    Dim StrSubject As String
    Dim StrName As String
    Dim StrFile As String
    Dim StrReceived As String
    Dim StrFolder As String
    Dim StrSaveFolder As String
    Dim StrFolderPath As String
    Dim iNameSpace As NameSpace
    Dim myOlApp As Outlook.Application
    Dim SubFolder As MAPIFolder
    Dim mItem As Object ' Changed to generic Object to handle various item types
    Dim FSO As Object
    Dim ChosenFolder As Object
    Dim Folders As New Collection
    Dim EntryID As New Collection
    Dim StoreID As New Collection
    Dim FileLog As Object
    Dim maxFileNameLength As Integer
     
    ' Setup FSO and Outlook objects
    On Error GoTo ErrorHandler
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set myOlApp = Outlook.Application
    Set iNameSpace = myOlApp.GetNamespace("MAPI")
    Set ChosenFolder = iNameSpace.PickFolder
    If ChosenFolder Is Nothing Then GoTo ExitSub
    
    ' Ask user for save folder location
    CustomBrowseForFolder StrSavePath
    If StrSavePath = "" Then
        MsgBox "No save location chosen. Exiting.", vbExclamation
        GoTo ExitSub
    End If
    
    ' Create log file to track errors
    Set FileLog = FSO.CreateTextFile(StrSavePath & "\EmailExport_Log.txt", True)

    ' Get all folders
    Call GetFolder(Folders, EntryID, StoreID, ChosenFolder)

    ' Set max filename length (255 total path limit minus date, file extension, and buffer for the directory)
    maxFileNameLength = 80 ' Adjust this if needed based on folder path length
    
    ' Loop through all folders and emails
    For i = 1 To Folders.Count
        StrFolder = StripIllegalChar(Folders(i))
        n = InStr(3, StrFolder, "\") + 1
        StrFolder = Mid(StrFolder, n, 256)
        StrFolderPath = StrSavePath & "\" & StrFolder & "\"
        StrSaveFolder = Left(StrFolderPath, Len(StrFolderPath) - 1) & "\"
        
        ' Ensure the full path exists
        EnsureFolderExists StrFolderPath
        
        ' Get folder and process emails
        Set SubFolder = myOlApp.Session.GetFolderFromID(EntryID(i), StoreID(i))
        For j = 1 To SubFolder.Items.Count
            Set mItem = SubFolder.Items(j)
            
            ' Only process Mail Items
            If TypeOf mItem Is MailItem Then
                ' Prepare email file
                StrReceived = Format(mItem.ReceivedTime, "YYYYMMDD-hhmm")
                StrSubject = mItem.Subject
                StrName = StripIllegalChar(StrSubject)
                
                ' Check for length limit to prevent truncation
                If Len(StrName) > maxFileNameLength Then
                    StrName = Left(StrName, maxFileNameLength) ' Trim subject line
                End If
                
                StrFile = StrSaveFolder & StrReceived & "_" & StrName & ".msg"
                
                ' Save email as .msg (using Unicode format)
                On Error GoTo SaveError
                mItem.SaveAs StrFile, olMSGUnicode
            End If
        Next j
    Next i
    
    GoTo ExitSub

SaveError:
    ' Log problematic emails
    FileLog.WriteLine "Error saving email: " & mItem.Subject & " in folder: " & SubFolder.Name & " Error: " & Err.Description
    Resume Next
    
ExitSub:
    FileLog.Close
    Set FileLog = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbCritical
    Resume ExitSub
End Sub

Function StripIllegalChar(StrInput)
    Dim RegX As Object
    
    On Error GoTo ErrorHandler
    Set RegX = CreateObject("vbscript.regexp")
    
    RegX.Pattern = "[\" & Chr(34) & "\!\@\#\$\%\^\&\*\(\)\=\+\|\[\]\{\}\`\'\;\:\<\>\?\/\,]"
    RegX.IgnoreCase = True
    RegX.Global = True
    
    StripIllegalChar = RegX.Replace(StrInput, "")
    
    Exit Function

ErrorHandler:
    MsgBox "Error stripping illegal characters: " & Err.Description, vbCritical
    Resume Next
End Function

Sub GetFolder(Folders As Collection, EntryID As Collection, StoreID As Collection, Fld As MAPIFolder)
    Dim SubFolder As MAPIFolder
    
    On Error GoTo ErrorHandler
    Folders.Add Fld.FolderPath
    EntryID.Add Fld.EntryID
    StoreID.Add Fld.StoreID
    For Each SubFolder In Fld.Folders
        GetFolder Folders, EntryID, StoreID, SubFolder
    Next SubFolder
    
    Exit Sub

ErrorHandler:
    MsgBox "Error getting folder: " & Err.Description, vbCritical
    Resume Next
End Sub

Sub CustomBrowseForFolder(ByRef StrFolderPath As String)
    Dim objShell As Object
    Dim objFolder As Object
    
    ' Create Shell application object
    Set objShell = CreateObject("Shell.Application")
    
    ' Open folder browser dialog, with a root folder at desktop level (0), which lets you browse any folder
    Set objFolder = objShell.BrowseForFolder(0, "Please select a folder:", &H1, 0)
    
    If Not objFolder Is Nothing Then
        ' Return the selected folder's path
        StrFolderPath = objFolder.self.Path
    Else
        ' If nothing is selected, return an empty string
        StrFolderPath = ""
    End If
End Sub

Sub EnsureFolderExists(StrFolderPath As String)
    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not FSO.FolderExists(StrFolderPath) Then
        ' If the full path does not exist, create it recursively
        CreateFullPath StrFolderPath
    End If
End Sub

Sub CreateFullPath(StrFolderPath As String)
    Dim FSO As Object
    Dim ParentFolder As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Get the parent directory path
    ParentFolder = FSO.GetParentFolderName(StrFolderPath)
    
    ' If the parent folder does not exist, create it first
    If Not FSO.FolderExists(ParentFolder) Then
        CreateFullPath ParentFolder
    End If
    
    ' Now create the folder
    If Not FSO.FolderExists(StrFolderPath) Then
        FSO.CreateFolder (StrFolderPath)
    End If
End Sub
