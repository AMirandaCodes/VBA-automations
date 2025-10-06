Sub ConvertToLatestWordFormat()
    Dim strFolderPath As String
    Dim strFileName As String
    Dim doc As Document
    Dim convertedFileName As String
    
    ' Prompt user to select a folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Word Documents"
        If .Show = -1 Then
            strFolderPath = .SelectedItems(1) & "\"
        Else
            MsgBox "No folder selected. Process cancelled."
            Exit Sub
        End If
    End With

    ' Get the first .doc file in the folder
    strFileName = Dir(strFolderPath & "*.doc")
    
    ' Loop through all .doc files in the folder
    Do While strFileName <> ""
        ' Open the document
        Set doc = Documents.Open(strFolderPath & strFileName, ReadOnly:=False)
        
        ' Set the new file name (change extension to .docx)
        convertedFileName = Left(strFileName, InStrRev(strFileName, ".")) & "docx"
        
        ' Save the document in the latest Word format (.docx)
        doc.SaveAs2 FileName:=strFolderPath & convertedFileName, FileFormat:=wdFormatXMLDocument
        
        ' Close the document
        doc.Close SaveChanges:=False
        
        ' Delete the original .doc file
        Kill strFolderPath & strFileName
        
        ' Get the next .doc file
        strFileName = Dir
    Loop
    
    MsgBox "Conversion completed successfully!"
End Sub
