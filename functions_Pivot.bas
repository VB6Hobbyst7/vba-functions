Attribute VB_Name = "functions_Pivot"
Option Base 1
Option Explicit


Public pt As PivotTable
Public wsPivot As Worksheet



Public Function pivotCreate(wb As Workbook, ws As Worksheet, dataStartingCell As String, Optional ptName As String)
'Create a new pivot table
    
    '=========================================================
    Dim ptCache As PivotCache
    Dim ptStartCell As String, ptEndCell As String
    Dim ptDataSource As String
    
    'Dim wb As Workbook, ws As Worksheet
    '=========================================================
    
    'Set wb = getActiveBook()
    'Set ws = getActiveSheet()
    
    If ptName = "" Then: ptName = "PivotTable"
    
    wb.Activate
    ws.Activate
    'Determine the starting and ending cells of the pivot table, in RXCY format
    ptStartCell = rangeRC(dataStartingCell)
    ptEndCell = rangeRC(findLastCell(dataStartingCell))
    
    'Specify the data source
    ptDataSource = ws.name & "!" & ptStartCell & ":" & ptEndCell
    'Insert a new sheet for the pivot table
    Set wsPivot = Sheets.Add
    'Set the pivot cachexx
    Set ptCache = wb.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=ptDataSource)
    'Set pivot table destination
    'Create the pivot table
    Set pt = ptCache.CreatePivotTable( _
        TableDestination:=wsPivot.name & "!R3C1", _
        TableName:=ptName)
    
    Sheets(wsPivot.name).Select
        

End Function


'Public Function pivotCreate(sourceDataSheetName As String, dataStartingCell As String, Optional ptName As String)
''Create a new pivot table
'
'    '=========================================================
'    Dim ptCache As PivotCache
'    Dim ptStartCell As String, ptEndCell As String
'    Dim ptDataSource As String
'
'    Dim wb As Workbook, ws As Worksheet
'    '=========================================================
'
'    Set wb = getActiveBook()
'    Set ws = getActiveSheet()
'
'    If ptName = "" Then: ptName = "PivotTable"
'
'    Sheets(sourceDataSheetName).Activate
'    'Determine the starting and ending cells of the pivot table, in RXCY format
'    ptStartCell = rangeRC(dataStartingCell)
'    ptEndCell = rangeRC(findLastCell(dataStartingCell))
'
'    'Specify the data source
'    ptDataSource = sourceDataSheetName & "!" & ptStartCell & ":" & ptEndCell
'    'Insert a new sheet for the pivot table
'    Set wsPivot = Sheets.Add
'    'Set the pivot cachexx
'    Set ptCache = wb.PivotCaches.Create( _
'        SourceType:=xlDatabase, _
'        SourceData:=ptDataSource)
'    'Set pivot table destination
'    'Create the pivot table
'    Set pt = ptCache.CreatePivotTable( _
'        TableDestination:=wsPivot.Name & "!R3C1", _
'        TableName:=ptName)
'
'    Sheets(wsPivot.Name).Select
'
'
'End Function



Public Function pivotClear() '(pt As PivotTable)
'Clear a pivot table entirely
    On Error Resume Next
    
    pt.ClearTable
    
End Function


Public Function pivotDelete()
'Deletes the sheet with the current pivot table

    On Error Resume Next
    
    wsPivot.Delete

End Function


Public Function pivotAddFilter(ptField As String) '(pt As PivotTable, ptField As String)
'In the specified pivot table, add a filter on a specific field

    On Error Resume Next
    
    With pt.PivotFields(ptField)
        .Orientation = xlPageField
        .Position = 1
    End With
    'Clear the filter
    pt.PivotFields(ptField).ClearAllFilters

End Function


Public Function pivotRemoveField(ptField As String) '(pt As PivotTable, ptField As String)
'Remove field from pivot table

    On Error Resume Next
    
    pt.PivotFields(ptField).Orientation = xlHidden
        
End Function


Public Function pivotClearFilter(ptField As String) '(pt As PivotTable, ptField As String)
'Clear all filters from a pivot table field

    On Error Resume Next
    
    pt.PivotFields(ptField).ClearAllFilters

End Function



Public Function pivotFilter(ptField As String, ptFieldFilter As Variant, showOrHide As Boolean) '(pt As PivotTable, ptField As String, ptFieldFilter As Variant, showOrHide As Boolean)
'Filter a field to only include/exclude certain values

    '===================================================
    Dim ptItem As PivotItem
    
    On Error Resume Next
    '===================================================
    
    pt.PivotFields(ptField).ClearAllFilters
    
    For Each ptItem In pt.PivotFields(ptField).PivotItems
        'Apply the filter (add or remove) to field values that match the search
        If iMatch(ptItem.name, ptFieldFilter) <> 0 Then
            ptItem.Visible = showOrHide
        Else
            ptItem.Visible = Not (showOrHide)
        End If
    Next ptItem
    
    'Enable selection of multiple fields
    pt.PivotFields(ptField).EnableMultiplePageItems = True

End Function


Public Function pivotFilterDate(ptDateField As String, ptDateFilter As String, after As Boolean) '(pt As PivotTable, ptDateField As String, ptDateFilter As String, after As Boolean)

    '===================================================
    Dim ptItem As PivotItem
    
    On Error Resume Next
    '===================================================
    
    pt.PivotFields(ptDateField).ClearAllFilters

    'Check to make sure the date filter was entered in proper format: DD/MM/YYYY (or D/M/YYYY, D/MM/YYYY, DD/M/YYYY)
    If Len(reMatch(ptDateFilter, "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$")) > 0 Then
    
        For Each ptItem In pt.PivotFields(ptDateField).PivotItems
            'Only filter if values in the pivot table are in the proper date format too
            If Len(reMatch(ptItem.name, "^[0-9]{1,2}/[0-9]{1,2}/[0-9]{4}$")) > 0 Then
                'If a date is on or after the date filter, make ot visible/hidden, depending on whether "after" was selected
                If CDate(ptItem.name) >= CDate(ptDateFilter) Then
                    ptItem.Visible = after
                Else
                    ptItem.Visible = Not (after)
                End If
            End If
            
            'Remove blanks
            If ptItem.name = "(blank)" Then: ptItem.Visible = False
            
        Next ptItem
    
    End If

    'Enable selection of multiple fields
    pt.PivotFields(ptDateField).EnableMultiplePageItems = True

End Function



Public Function pivotAddValue(ptField As String, Optional sum As Boolean) '(pt As PivotTable, ptField As String, Optional sum As Boolean)
    
    On Error Resume Next
    
    If sum Then
    'If sum is selected (True), put the sum of the selected field
        pt.AddDataField pt.PivotFields(ptField), "Sum of " & ptField, xlSum
    Else
    'If sum is not selected (False), put the count of the selected field
        pt.AddDataField pt.PivotFields(ptField), "Count of " & ptField, xlCount
    End If


End Function



Public Function pivotAddRow(ptField As String, Optional showAll As Boolean) '(pt As PivotTable, ptField As String, Optional showAll As Boolean)

    On Error Resume Next
    
    With pt.PivotFields(ptField)
        .Orientation = xlRowField
        .Position = 1
        'Show all items, including those without any observations - this will keep the data aligned
        .ShowAllItems = showAll
    End With

End Function



Public Function pivotGroupBy(byMonths As Boolean, byYears As Boolean) '(pt As PivotTable, byMonths As Boolean, byYears As Boolean)

    '=============================================
    Dim startCell As String
    
    On Error Resume Next
    '=============================================
    
    'Determine the starting cell of the pivot table (the header for row labels)
    startCell = pivotStartCell(True)
    Range(startCell).Select
    'Group by the values selected
    Selection.Group Start:=True, End:=True, Periods:=Array(False, False, False, _
        False, byMonths, False, byYears)
    
End Function



Public Function pivotUngroup() '(pt As PivotTable)

    Dim startCell As String
    
    On Error Resume Next
    
    startCell = pivotStartCell(True)
    Range(startCell).Select
    Selection.Ungroup
    
End Function



Public Function pivotCopy(Optional excludeFirstColumn As Boolean) '(pt As PivotTable, Optional excludeFirstColumn As Boolean)

    '=============================================
    Dim startCell As String, endCell As String
    
    On Error Resume Next
    '=============================================
    
    startCell = pivotStartCell()
    endCell = pivotEndCell()
    
    'If excludeFirstColumn is true, don't copy the first column
    If excludeFirstColumn Then
        'Offset the starting column by 1
        Range(Range(startCell).Offset(0, 1).Address & ":" & endCell).copy
    Else
        Range(startCell & ":" & endCell).copy
    End If
    

End Function


Public Function pivotStartCell(Optional actualDataStart As Boolean) '(pt As PivotTable, Optional actualDataStart As Boolean)

    '=============================================
    Dim ptRange As String
    
    On Error Resume Next
    '=============================================
    
    'Find the range of the pivot table. In the format "$A$1:$B$10"
    ptRange = pt.DataBodyRange.CurrentRegion.Address
    'The starting cell in the pivot table range. Searches for one or more capital letters, numbers, or dollar signs...
    'It will return the first part up to but not including the colon ("$A$1")
    pivotStartCell = reMatch(ptRange, "[$A-Z0-9]+")
    'If the actualDataStart is true, want to offset by 1 row so it selects the actual data, not the header
    If actualDataStart Then: pivotStartCell = Range(pivotStartCell).Offset(1, 0).Address
    'Remove dollar signs ("A1")
    pivotStartCell = Replace(pivotStartCell, "$", "")

End Function



Public Function pivotEndCell() '(pt As PivotTable)

    '=============================================
    Dim ptRange As String
    
    On Error Resume Next
    '=============================================
    
    'Find the range of the pivot table. In the format "$A$1:$B$10"
    ptRange = pt.DataBodyRange.CurrentRegion.Address
    
    'The ending cell in the pivot table range. Searches for a colon ,followed by one or more capital letters, numbers, or dollar signs...
    'It will return the second part, including the colon (":$B$10"), and then remove the colon ("$B$10")
    pivotEndCell = reMatch(ptRange, ":[A-Z0-9$]+", ":")

End Function


