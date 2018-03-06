Attribute VB_Name = "functions_General"
Option Explicit
Option Base 1


Public Function appAlerts(onOrOff As Boolean)

    Application.DisplayAlerts = onOrOff
    Application.ScreenUpdating = onOrOff

End Function


Public Function getMyDocuments() As String

    getMyDocuments = CreateObject("WScript.Shell").SpecialFolders("MyDocuments")

End Function


Public Function getActiveBook()

    On Error Resume Next
    
    Set getActiveBook = ActiveWorkbook
    
End Function


Public Function getActiveSheet()

    Dim wb As Workbook
    On Error Resume Next
    
    Set wb = getActiveBook()
    
    If Not (wb Is Nothing) Then: Set getActiveSheet = wb.ActiveSheet
    
End Function


Public Function workbookByName(workbookName As String) As Workbook

    Dim wb As Workbook
    On Error Resume Next
        
    For Each wb In Application.Workbooks
        If wb.name = workbookName Then
            Set workbookByName = wb
            wb.Activate
        End If
    Next wb
    
    If workbookByName Is Nothing Then: Set workbookByName = getActiveBook

End Function


Public Function worksheetByName(Optional wb As Workbook, Optional worksheetName As String) As Worksheet
    
    Dim ws As Worksheet
    On Error Resume Next
    
    If wb Is Nothing Then: Set wb = getActiveBook
    
    If wb.Sheets.Count = 1 Then
        Set worksheetByName = wb.Sheets(1)
    Else
        For Each ws In wb.Sheets
            If ws.name = worksheetName Then
                Set worksheetByName = ws
                ws.Activate
            End If
        Next ws
    End If
    
    If worksheetByName Is Nothing Then: Set worksheetByName = getActiveSheet

End Function




Public Function findLastCell(cellWithinRange As String, Optional wb As Workbook, Optional ws As Worksheet)

    Dim sht As Worksheet
    
    If Not (ws Is Nothing) Then
        Set sht = ws
    Else
        If Not (wb Is Nothing) Then
            Set sht = wb.ActiveSheet
        Else
            Set sht = ActiveSheet
        End If
    End If
        
    findLastCell = sht.Range(cellWithinRange).SpecialCells(xlLastCell).Address

End Function



Public Function findLastRow(cellWithinRange As String, Optional wb As Workbook, Optional ws As Worksheet)

    Dim sht As Worksheet
    
    If Not (ws Is Nothing) Then
        Set sht = ws
    Else
        If Not (wb Is Nothing) Then
            Set sht = wb.ActiveSheet
        Else
            Set sht = ActiveSheet
        End If
    End If
    
    findLastRow = sht.Range(cellWithinRange).SpecialCells(xlLastCell).Row

End Function



Public Function findLastColumn(cellWithinRange As String, Optional wb As Workbook, Optional ws As Worksheet)

    Dim sht As Worksheet
    
    If Not (ws Is Nothing) Then
        Set sht = ws
    Else
        If Not (wb Is Nothing) Then
            Set sht = wb.ActiveSheet
        Else
            Set sht = ActiveSheet
        End If
    End If
    
    findLastColumn = sht.Range(cellWithinRange).SpecialCells(xlLastCell).Column

End Function


Public Function rangeRC(rng As String)

    '======================================================
    On Error Resume Next
    '=====================================================
    
    'If the range supplied is in the format "RXCY", just return it
    If Len(reMatch(rng, "R[0-9]+C[0-9]+")) > 0 Then
        rangeRC = rng
    ' If the range supplied is in the format "A1", convert to RXCY format
    ElseIf Len(reMatch(rng, "^[$]?[A-Z]+[$]?[0-9]+$")) > 0 Then
        rangeRC = Range(rng).Address(ReferenceStyle:=xlR1C1)
    End If
    
End Function



Public Function rangeAbs(rng As String)

    '======================================================
    Dim rngRow As Long, rngCol As Long
    
    On Error Resume Next
    '=====================================================
    
    'If the range supplied is in the format "RXCY", convert to "A1" format
    If Len(reMatch(rng, "R[0-9]+C[0-9]+")) > 0 Then
        rngRow = Val(reMatch(rng, "R[0-9]+", "R"))
        rngCol = Val(reMatch(rng, "C[0-9]+", "C"))
        rangeAbs = Cells(rngRow, rngCol).Address(ReferenceStyle:=xlA1)
        rangeAbs = Replace(rangeAbs, "$", "")
    ' If the range supplied is in the format "A1", just return it
    ElseIf Len(reMatch(rng, "^[$]?[A-Z]+[$]?[0-9]+$")) > 0 Then
        rangeAbs = rng
    End If
    
End Function



Public Function countRows(Optional nameWS As String, Optional hasHeader As Boolean)
'Counts the number of rows in the selected sheet
'If no sheet name is provided, it uses the active sheet in the active workbook
'If header is not specified, it assumes there is none (counts all rows)

    '============================================================
    Dim wb As Workbook, ws As Worksheet
    '==========================================================
    
    Set wb = ActiveWorkbook
    
    'Check if a worksheet name is provided. If not, refer to the active sheet
    If nameWS = "" Then
        Set ws = wb.ActiveSheet
    Else
        Set ws = wb.Worksheets(nameWS)
    End If
    
    countRows = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    'If the sheet has a header, subtract 1 from the count
    If hasHeader Then
        countRows = countRows - 1
    End If


End Function



Public Function countColumns(Optional nameWS As String)
'Counts the number of columns in the selected sheet
'If no sheet name is provided, it uses the active sheet in the active workbook

    '===================================================
    Dim wb As Workbook, ws As Worksheet
    '==================================================
    
    Set wb = ActiveWorkbook
    
    'Check if a worksheet name is provided. If not, refer to the first sheet in the workbook
    If nameWS = "" Then
        Set ws = wb.ActiveSheet
    Else
        Set ws = wb.Worksheets(nameWS)
    End If
    
    countColumns = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

End Function



Public Function whichColumn(varName As String, Optional rowHeader As Long) As Long
'Finds the column containing the column name of interest

    '====================================================
    Dim wb As Workbook, ws As Worksheet
    Dim nCol As String
    Dim j As Long
    '====================================================
    
    Set wb = ActiveWorkbook
    Set ws = wb.ActiveSheet
    
    nCol = countColumns()
    
    'If the optional parameter rowHeader is not provided, set it to 1
    If rowHeader = 0 Then
        rowHeader = 1
    End If
    
    While (whichColumn = 0) And (j < nCol)
        j = j + 1
        If ws.Cells(rowHeader, j) = varName Then
            whichColumn = j
        End If
    Wend

End Function



Public Function calcMinDate(dateColumnName As String)
'Calculates the minimum date in a column of dates - needed for filtering dates in pivot tables

    '====================================================
    Dim colDate As String
    Dim arrDates() As Double
    Dim minDate As Double
    Dim lastRow As Double
    '====================================================
    
    'Find the column index (ex: A, AA, etc.) using regex to remove the $ that show up when .Address is called
    colDate = columnLetter(dateColumnName)
    
    'Create array of all dates
    lastRow = findLastRow(colDate & "1")
    arrDates = rangeToArray(colDate & 2, colDate & lastRow)
    'Calculate the minimum date
    calcMinDate = min(arrDates, False)

End Function



Public Function columnLetter(colName As String)

    '====================================================
    Dim colAddress As String
    '====================================================
    
    colAddress = Cells(1, whichColumn(colName)).Address
    
    columnLetter = reMatch(colAddress, "[A-Z]+")

End Function


Public Function reMatch(stringToSearch As String, stringToFind As String, Optional substringToRemove As String)

    '====================================================
    Dim regex As Object, regexMatches As Object
    'Dim reRemoveSubstring As String
    
    On Error Resume Next
    '====================================================
    
    'Find the column index (ex: A, AA, etc.) using regex to remove the $ that show up when .Address is called
    Set regex = CreateObject("VBScript.RegExp")
    'Find matches containing one or more capital letters - this will only find the column index (A, B, C,...)
    regex.Pattern = stringToFind
    'Find a string matching the column's address (the address will be in the form "$A$1", etc.
    Set regexMatches = regex.Execute(stringToSearch)
    reMatch = regexMatches(0)
    
    regex.Pattern = substringToRemove
    Set regexMatches = regex.Execute(stringToSearch)
    substringToRemove = regexMatches(0)
    
    'Removes a part of the matched string, if this argument is provided
    reMatch = Replace(reMatch, substringToRemove, "")

End Function


Public Function rangeToArray(startCell As String, endCell As String)
'Convert a range in Excel into an array in VBA
'Need to add error catching

    '============================= Assign variables ==============================
    Dim nRow As Long, nCol As Long
    Dim rngLen As Long
    Dim outArray() As String
    Dim i As Long, j As Long
    '==============================================================================

    'Count the number of rows and columns in the selected range - either nRow or nCol will equal 1
    nRow = Range(endCell).Row - Range(startCell).Row + 1
    nCol = Range(endCell).Column - Range(startCell).Column + 1
    'Set the range length to the larger dimension
    'rngLen = Application.WorksheetFunction.max(nRow, nCol)
    
    'Resize the output array to the appropriate size
    ReDim outArray(1 To nRow, 1 To nCol)
    
    'Assign each value in the selected range to the array
    For i = 1 To nRow
        For j = 1 To nCol
            outArray(i, j) = Range(startCell).Offset(i - 1, j - 1)
            'outArray(i, j) = Range(startCell).Offset((i - 1) * (nRow - 1) / (rngLen - 1), (i - 1) * (nCol - 1) / (rngLen - 1))
        Next j
    Next i
    
    'It seems like you need to pass around the array variable outArray, then assign it to rangeToArray at the end
    'This is because rangeToArray is already a function that takes inputs, so you can't reference its indices (such as rangeToArray(i) in the loop above)
    rangeToArray = outArray

End Function


Public Function min(arr As Variant, Optional includeZero As Boolean) As Double

    '=======================================
    Dim i As Long
    Dim inf As Double
    '=======================================
    
    inf = infinity()
    min = inf
    'Loop through each element in the array to find the minimum
    For i = 1 To UBound(arr)
        'If the current array value is 0 and ignore zero is true, skip it
        If arr(i) <> 0 Or includeZero = True Then
            If arr(i) < min Then
                min = arr(i)
            End If
        End If
    Next i
    
    If min = inf Then: min = 0
    
End Function



Public Function max(arr As Variant, Optional includeZero As Boolean) As Double

    '=======================================
    Dim i As Long
    Dim inf As Double
    '=======================================
    
    inf = infinity()
    max = -inf
    'Loop through each element in the array to find the max
    For i = 1 To UBound(arr)
        'If the current array value is 0 and ignore zero is true, skip it
        If arr(i) <> 0 Or includeZero = True Then
            If arr(i) > max Then
                max = arr(i)
            End If
        End If
    Next i
    
    If max = -inf Then: max = 0

End Function


Public Function iMatch(findVal As Variant, searchIn As Variant)

    Dim i As Long
    Dim matchFound As Boolean
    
    'If the value to search is an array, find the array index containing a match
    If IsArray(searchIn) Then
        'Search until a match is found or the array ends
        While matchFound = False And i < UBound(searchIn)
            i = i + 1
            If findVal = searchIn(i) Then
                iMatch = i
                matchFound = True
            End If
        Wend
    'If the value to search is not an array, check whether they match
    Else
        If findVal = searchIn Then: iMatch = 1
    End If

End Function



Public Function infinity() As Double
'Allows infinity to be assigned to a variable
'Need to turn error checking back on

    On Error Resume Next
    infinity = 1 / 0

End Function



Public Function createArray(arrayAsString, Optional sep As String, Optional numeric As Boolean) As Variant

    '=============================
    Dim i As Integer
    '==============================
    
    'Checks if separator is entered. If not, use comma
    If sep = "" Then
        sep = ","
    End If
    
    createArray = Split(arrayAsString, sep)
    
    If numeric Then
        For i = 0 To UBound(createArray)
            'MsgBox (i & ": " & createArray(i))
            'createArray(i) = Val(createArray(i))
        Next i
    End If

End Function



Public Function deleteSheets(nameOfSheetToKeep As String)

    '===================================================
    Dim wb As Workbook
    Dim ws As Worksheet
    '====================================================
    
    Call appAlerts(False)
    Set wb = getActiveBook
    'Set ws = getActiveSheet
    
    If wb.Worksheets.Count > 1 Then
        For Each ws In wb.Worksheets
            If ws.name <> nameOfSheetToKeep Then
                ws.Delete
            End If
        Next ws
    End If
        
    Call appAlerts(True)
    
End Function



Public Function deleteSheetsPrompt()

    Dim sheetPrompt As String
    
    sheetPrompt = InputBox("Enter sheet name to keep:")
    
    If sheetPrompt <> "" Then: Call deleteSheets(sheetPrompt)

End Function



Public Function errorMessage(requiredFile As String)

    MsgBox ("In order to run, the data sheet of a(n) " & requiredFile & " file must be active." & vbNewLine & _
            "The current active sheet is " & wsThis.name & " in the workbook " & wbThis.name & "." & vbNewLine & vbNewLine & _
            "Also verify that filters are turned off and headers have not been renamed.")

End Function



Public Function mostRecentVC(Optional fy As Long)
'Returns the FULL PATH to the most recent ViewCreation file

    '=========================================================================
    Const yearlyDataFolder = "\\TYMX-FS-001v\afrs\AFRS_Org\RSO (FOUO)\RSOA (FOUO)\RSOAP\_Data\Yearly Data"
    Dim searchFolder As String
    Dim maxMonthFolder As String, maxDayFolder As String
    '========================================================================
    
    'If FY was not supplied, take the FY of today's date
    If fy = 0 Then
        fy = Year(Date) - 2000
        If Month(Date) >= 10 Then: fy = fy + 1
    End If
    
    'Search within the Yearly data folder for the max month
    searchFolder = yearlyDataFolder & "\" & "FY" & fy
    maxMonthFolder = findMaxFolder(searchFolder)
    'Search within the max month folder for the max date
    searchFolder = searchFolder & "\" & maxMonthFolder
    maxDayFolder = findMaxFolder(searchFolder)
    
    searchFolder = searchFolder & "\" & maxDayFolder
    mostRecentVC = searchFolder & "\" & findFile(searchFolder, "ViewCreation", True, "xlsx") '"txt"
   
End Function


Public Function findMaxFolder(searchDirectory As String) As String
'Finds the folder with the largest numeric value

    '======================================================
    Dim findDirs As Variant
    Dim folderVal As Long, maxVal As Long
    '======================================================
    
    'Search for all directories within the specified search location
    findDirs = Dir(searchDirectory & "\", vbDirectory)
    'Loop through each directory found
    While findDirs <> ""
        'Extract the numeric value from the folder
        folderVal = Val(reMatch(CStr(findDirs), "[0-9]+"))
        'Determine whether the current folder value is the max value found
        If folderVal > maxVal Then
            maxVal = folderVal
            findMaxFolder = CStr(findDirs)
        End If
        'This line is neeeded to go to the next item in the Dir search...I don't understand exactly how it works
        findDirs = Dir
    Wend

End Function


Public Function findFile(searchDirectory As String, fileName As String, Optional partialMatch As Boolean, Optional fileExt As String)
'Find a file within a search directory - either a full match or partial, and with or without a certain file extension

    '======================================================
    Dim findDirs As Variant
    Dim stringToMatch As String
    '=====================================================
    
    'Remove any periods that were included in the file extension
    fileExt = Replace(fileExt, ".", "")

    'If a partial match is specified, search for the matching string anywhere within the file name
    If partialMatch = True Then
        stringToMatch = "^.*" & fileName & ".*[\.]"
    Else
    'If a partial match is not specified, find an exact match
        stringToMatch = "^" & fileName & "[\.]"
    End If
    
    'If a file extension is specified, add it to the end of the string to match
    If fileExt <> "" Then
        stringToMatch = stringToMatch & fileExt & "$"
    Else
    'If a file extension is not specified, only search for the file name
        stringToMatch = stringToMatch & ".*$"
    End If
    
    findDirs = Dir(searchDirectory & "\", vbNormal)
    While findDirs <> ""
        If Len(reMatch(CStr(findDirs), stringToMatch)) > 0 Then
            findFile = CStr(findDirs)
        End If
        findDirs = Dir
    Wend

End Function


Public Function doesWorkbookExist(workbookPath As String, Optional wbToSet As Workbook) As Boolean

    Dim wb As Workbook
    Dim msgText As String
    
    For Each wb In Workbooks
        If wb.FullName = workbookPath Then
            doesWorkbookExist = True
            If wbToSet Is Nothing Then: Set wbToSet = wb
        End If
            'msgText = "File found:" & vbNewLine & wb.FullName
    Next wb
    
    'If wbToSet Is Nothing And doesWorkbookExist = True Then: Set wbToSet = wb
    'MsgBox msgText
    
End Function



Public Function selectFiles()

    '================================================
    Dim intChoice As Integer
    Dim nFiles As Integer
    Dim i As Integer
    Dim filePaths() As String
    '==================================================
    
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        intChoice = .Show
        
        If intChoice <> 0 Then
            nFiles = .SelectedItems.Count
            ReDim filePaths(nFiles)
            
            selectFiles = .SelectedItems(1)
            For i = 1 To nFiles
                filePaths(i) = .SelectedItems(i)
            Next i
        End If
    End With
    
    'Need to assign the array to the function at the end
    selectFiles = filePaths


End Function



Public Function selectFolder() As String

    '================================================
    Dim intChoice As Integer
    '==================================================
    
    With Application.FileDialog(msoFileDialogFolderPicker)
        intChoice = .Show
        If intChoice <> 0 Then: selectFolder = .SelectedItems(1)
    End With


End Function


Public Function uniqueFileName(fileSaveLocation As String, desiredFileName As String, fileExt As String) As String

    Dim i As Integer
    
    'If a file with the desired name already exists, add a unique suffix
    If findFile(fileSaveLocation, desiredFileName, False, fileExt) <> "" Then
        i = 1
        While findFile(fileSaveLocation, desiredFileName & "\(" & i & "\)", False, fileExt) <> ""
            i = i + 1
        Wend
        uniqueFileName = desiredFileName & "(" & i & ")"
    Else
        uniqueFileName = desiredFileName
    End If

End Function
