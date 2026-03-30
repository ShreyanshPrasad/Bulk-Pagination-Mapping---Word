Sub GeneratePaginationMapping()

    Dim wdApp As Object
    Dim wdDoc As Object
    Dim folderPath As String
    Dim fileName As String
    
    Dim rowNum As Long
    Dim startPage As Long
    Dim pageCount As Long
    Dim endPage As Long
    
    ' ?? CHANGE THIS PATH
    folderPath = "your\input\path\where\files\are\kept\"
    
    ' Start Word
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    
    ' Excel headers
    Cells(1, 1).Value = "File Name"
    Cells(1, 2).Value = "Start Page"
    Cells(1, 3).Value = "End Page"
    
    rowNum = 2
    startPage = 1
    
    fileName = Dir(folderPath & "*.docx")
    
    Do While fileName <> ""
        
        On Error Resume Next
        
        Set wdDoc = wdApp.Documents.Open(folderPath & fileName, False, True)
        
        If Err.Number <> 0 Then
            Cells(rowNum, 1).Value = fileName
            Cells(rowNum, 2).Value = "ERROR"
            Cells(rowNum, 3).Value = "ERROR"
            Err.Clear
            GoTo NextFile
        End If
        
        ' Get actual page count
        pageCount = wdDoc.ComputeStatistics(2) ' wdStatisticPages
        
        ' Calculate end page
        endPage = startPage + pageCount - 1
        
        ' Write to Excel
        Cells(rowNum, 1).Value = fileName
        Cells(rowNum, 2).Value = startPage
        Cells(rowNum, 3).Value = endPage
        
        ' Update start page for next file
        startPage = endPage + 1
        
        wdDoc.Close False
        
        rowNum = rowNum + 1
        
NextFile:
        fileName = Dir
        
    Loop
    
    wdApp.Quit
    
    MsgBox "Mapping file created successfully!"

End Sub
