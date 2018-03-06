Attribute VB_Name = "functions_UserForm"
Option Explicit
Option Base 1


Type requiredWorbooks
    name As String
    wbObj As Workbook
    wsObj As Worksheet
    path As String
    exists As Boolean
    labelObject As String
    labelPathObject As String
    expectedPath As String
    expectedNumCols As Variant
    expectedA1Value As String
End Type



Public Sub showMenu()
    ufUpload.Show
End Sub


Public Sub ufObjCenter(objectToCenter As Object, objectToCenterOn As Object, Optional centerHorizontal As Boolean, Optional centerVertical As Boolean)

    Dim obj1Width As Double, obj1Height As Double, obj2Width As Double, obj2Height As Double

    obj1Width = objectToCenter.Width
    obj1Height = objectToCenter.Height
    obj2Width = objectToCenterOn.Width
    obj2Height = objectToCenterOn.Height
    
    If centerHorizontal Then: objectToCenter.Left = (obj2Width - obj1Width) / 2
    If centerVertical Then: objectToCenter.Top = (obj2Height - obj1Height) / 2
    

End Sub



Public Sub ufObjPosition(obj As Object, objPositions() As Variant) 'objLeft As Double, Optional objTop As Double, Optional objWidth As Double, Optional objHeight As Double)
    
    'Left, Top, Width, Height
    
    If Val(objPositions(1)) >= 0 Then: obj.Left = Val(objPositions(1)) 'objLeft
    If Val(objPositions(2)) >= 0 Then: obj.Top = Val(objPositions(2)) 'objTop
    If Val(objPositions(3)) > 0 Then: obj.Width = Val(objPositions(3)) 'objWidth
    If Val(objPositions(4)) > 0 Then: obj.Height = Val(objPositions(4)) 'objHeight
    
End Sub



Public Function objectMatch(objToFind As Object, objsToSearch As Variant) As Long

    '============================================
    Dim i As Long
    
    On Error Resume Next
    '============================================
    
    If IsArray(objsToSearch) Then
        For i = 1 To UBound(objsToSearch)
            If objToFind.name = objsToSearch(i).name Then
                objectMatch = i
            End If
        Next i
    Else
        If objToFind.name = objsToSearch.name Then: objectMatch = 1
    End If
    
End Function
