Attribute VB_Name = "Module3"
Sub CreateFiles()

    Dim i As Integer
    Dim folderPath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim filePath As String
    
    folderPath = "C:\Users\[username]\Excel\ExcelFiles\"
    'Replace [username] with your windows _
    user profile name. Make sure you create the _
    necessary folders present in the folderPath _
    Or else, customize folderpath at your own convenience.
    
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 1 To 100
        fileName = "ExcelFilesBook" & i
        filePath = folderPath & fileName
        Set wb = Workbooks.Add
        wb.SaveAs _
            fileName:=filePath, _
            FileFormat:=xlOpenXMLWorkbook
        wb.Close True
    Next i
    
    MsgBox "The files have been created and saved successfully"
    
End Sub
