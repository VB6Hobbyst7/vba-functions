Attribute VB_Name = "functions_Excel"
Option Explicit
Option Base 1


Public Function excelOpen(filePath As String) As Workbook
    
    Dim wb As Workbook ', wbThis As Workbook
    On Error Resume Next
    
    'Set wbThis = ActiveWorkbook
    
    For Each wb In Workbooks
        If wb.FullName = filePath Then: Set excelOpen = wb
    Next wb
    If excelOpen Is Nothing Then: Set excelOpen = Workbooks.Open(filePath) 'Workbooks.Add 'Set excelOpen = Workbooks.Open(filePath)
    'ActiveWindow.Visible = False
    'excelOpen.Visible = False

    'wbThis.Activate
    
    'excelOpen.Windows(1).Visible = True
   ' excelOpen.Visible = True
    

End Function


Public Function excelClose(wb As Workbook, Optional saveChanges As Boolean)

    On Error Resume Next
    
    If Not (wb Is Nothing) Then
        If saveChanges = True Then: wb.save
        wb.Close
        Set wb = Nothing
    End If
    
End Function



Public Function excelClear(startCell As String, endCell As String, Optional wb As Workbook, Optional ws As Worksheet, Optional clearAll As Boolean)

    On Error Resume Next
    
    If wb Is Nothing Then: Set wb = getActiveBook
    If ws Is Nothing Then: Set ws = getActiveSheet
    
    If clearAll = True Then
        ws.Range(startCell & ":" & endCell).Clear
    Else
        ws.Range(startCell & ":" & endCell).ClearContents
    End If
    
    DoEvents

End Function


Public Function excelCopy(startCell As String, endCell As String, Optional wb As Workbook, Optional ws As Worksheet)

    On Error Resume Next
    
    If wb Is Nothing Then: Set wb = getActiveBook
    If ws Is Nothing Then: Set ws = getActiveSheet
    
    ws.Range(startCell & ":" & endCell).copy
    DoEvents

End Function


Public Function excelPaste(destinationRange As String, Optional wb As Workbook, Optional ws As Worksheet, Optional pasteValues As Boolean, Optional skipBlanksYesOrNo As Boolean, Optional transposeYesOrNo As Boolean)

    If wb Is Nothing Then: Set wb = getActiveBook
    If ws Is Nothing Then: Set ws = getActiveSheet
    
    'wb.Activate
    
    If pasteValues Then
        ws.Range(destinationRange).PasteSpecial paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:= _
            skipBlanksYesOrNo, Transpose:=transposeYesOrNo
    Else
        ws.Range(destinationRange).PasteSpecial paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            skipBlanksYesOrNo, Transpose:=transposeYesOrNo
    End If
    
    'Clears the clipboard
    Application.CutCopyMode = False
        
End Function



'Need to figure out how to set the array to the number of columns - count commas?
Public Function readTxt(filePath As String, hasHeader As Boolean, delimiter As String) As Workbook

    '===========================================================
    Dim ws As Worksheet
    Dim fileName As String
    Dim qt As QueryTable
    '==========================================================
    
    Set readTxt = Workbooks.Add
    Set ws = readTxt.Sheets(1)
    
    fileName = reMatch(filePath, ".*[\\][A-z0-9]+", ".*[\\]")
    
    
'    Set qt = ws.QueryTables.Add( _
'        Connection:="TEXT;" & filePath, _
'        Destination:=ws.Range("$A$1") _
'        )
    
'    With ws.QueryTables.Add(Connection:="TEXT;" & filePath, Destination:=ws.Range("$A$1"))
'        .Name = fileName
'        .FieldNames = hasHeader
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .TextFileConsecutiveDelimiter = False
'        .TextFileTabDelimiter = True
'        .TextFileOtherDelimiter = delimiter
'        .TextFileColumnDataTypes = Array(1)
'        .Refresh BackgroundQuery:=False
'    End With
    
End Function



Sub readTxt2() '(filePath As String, delimiter As String)

    '============================================
    Dim filePath As String
    Dim delimiter As String
    
    Dim currentFile As String
    Dim i As Long
    Dim currentLine As String
    Dim lineItems() As String
    '=============================================
    

    filePath = "C:\Users\1398909107A\Desktop\excel macros\vc.txt"
    delimiter = "|"
    
    Open filePath For Input As #1
    
    Do Until EOF(1)
        i = i + 1
        Line Input #1, currentLine
        lineItems = Split(currentLine, delimiter)
        MsgBox (i & vbNewLine & lineItems(1))
    Loop
    
    Close #1

End Sub





