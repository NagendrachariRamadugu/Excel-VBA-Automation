Attribute VB_Name = "Module1"
Sub renameFiles()

    Dim folderPath As String
    Dim newFileName As String
    Dim oldFileName As String
    Dim oldFilePath As String
    Dim newFilePath As String
    Dim nameOnly As String
    
    folderPath = "C:\Users\[username]\Excel\ExcelFiles\"
    'Replace [username] with your windows _
    user profile name. Make sure you create the _
    necessary folders present in the folderPath _
    Or else, customize folderpath at your own convenience.
    
    oldFileName = Dir(folderPath & "*.xlsx")
    
    Do While oldFileName <> ""
        oldFilePath = folderPath & oldFileName
        nameOnly = Left(oldFileName, WorksheetFunction.Find(".", oldFileName, 1))
        newFileName = Replace(oldFileName, "ExcelFiles", "")
        newFilePath = folderPath & newFileName
        Name oldFilePath As newFilePath
        oldFileName = Dir()
    Loop
        
End Sub
