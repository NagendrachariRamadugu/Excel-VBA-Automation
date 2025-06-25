Attribute VB_Name = "Module2"
Sub deleteFiles()

    Application.DisplayAlerts = False
    
    Dim i As Integer
    Dim folderPath As String
    Dim filePath As String
    Dim fileName As String
    
    folderPath = "C:\Users\[username]\Excel\ExcelFiles\"
    'Replace [username] with your windows _
    'user profile name. Make sure you create the _
    'necessary folders present in the folderPath _
    'Or else, customize folderpath at your own convenience.
    
    fileName = Dir(folderPath & "*.xlsx")
    
    Application.DisplayAlerts = False
    
    Do While fileName <> ""
        filePath = folderPath & fileName
        Kill (filePath)
        fileName = Dir()
     Loop
     
     MsgBox "All files have been deleted"

End Sub
