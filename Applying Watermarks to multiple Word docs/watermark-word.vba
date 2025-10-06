Sub ApplyCustomWatermarkToDocuments()
    Dim folderPath As String
    Dim fileName As String
    Dim doc As Document
    Dim watermarkName As String
    Dim templatePath As String
    
    ' Customize these variables
    folderPath = "C:\Path\To\Your\Documents\" ' Folder containing the Word documents
    watermarkName = "Watermark Name" ' Name of your custom watermark
    templatePath = "C:\Path\To\Watermark Template" ' Full path to the template containing the watermark
    
    ' Ensure folder path ends with a backslash
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Get the first file in the folder
    fileName = Dir(folderPath & "*.doc*") ' Looks for .doc and .docx files
    
    ' Load the template containing the custom watermark
    Dim templateDoc As Document
    Set templateDoc = Documents.Open(templatePath, ReadOnly:=True)
    
    ' Process each file in the folder
    While fileName <> ""
        ' Open the document
        Set doc = Documents.Open(folderPath & fileName)
        
        ' Apply the custom watermark
        On Error Resume Next ' Avoid errors if the watermark is already applied
        templateDoc.AttachedTemplate.AutoTextEntries(watermarkName).Insert _
            Where:=doc.Sections(1).Headers(wdHeaderFooterPrimary).Range, _
            RichText:=True
        On Error GoTo 0
        
        ' Save and close the document
        doc.Save
        doc.Close
        fileName = Dir ' Get the next file
    Wend
    
    ' Close the template
    templateDoc.Close SaveChanges:=False
    
    MsgBox "Watermark applied to all documents in the folder.", vbInformation
End Sub
