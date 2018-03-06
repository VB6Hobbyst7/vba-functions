Attribute VB_Name = "functions_PowerPoint"
Option Explicit
Option Base 1


Public pp As PowerPoint.Application
Public ppPres As Presentation
Public ppSlide As Slide
Public ppShape As Object



Public Function pptCreateNew(Optional templateFileLocation As String, Optional saveLocally As Boolean)

    '=========================================================
    Dim templateFolder As String
    Dim copiedTemplateLocation As String
    '=========================================================
    
    Set pp = New PowerPoint.Application
    
    'If no template file is provided, create a generic presentation
    If templateFileLocation = "" Then
        Set ppPres = pp.Presentations.Add
    'If a template file is provided, use it to set the slide themes
    Else
        'Open the ORIGINAL template
        Set ppPres = pp.Presentations.Open(templateFileLocation)
        
        'Use regex to determine the FOLDER containing the template (not the path including the file)
        'This code will remove the file name ([filename].pptx)
        templateFolder = reMatch(templateFileLocation, ".+.ppt[x]?", "\\[^\\]+.ppt[x]?")
        
        'Set the copy location depending on whether local copy was selected or not
        If saveLocally Then
            copiedTemplateLocation = getMyDocuments()
        Else
            copiedTemplateLocation = templateFolder
        End If
        'Set the copied template location - add the seconds (represented as a decimal number) in the current time - basically a random number
        copiedTemplateLocation = copiedTemplateLocation & "\" & "temp-" & reMatch(CDbl(Now()), ".+", "[0-9]+\.") & ".pptx"
        'Save the original template to the new template location
        ppPres.SaveAs copiedTemplateLocation, ppSaveAsDefault
        'Close the ORIGINAL template
        ppPres.Close
        'Open the COPIED template
        Set ppPres = pp.Presentations.Open(copiedTemplateLocation)
    End If
    
End Function


Public Function pptAddSlide(Optional layoutName As String, Optional slideIndex As Long)

    '==============================================
    Dim slideLayout As CustomLayout
    '=============================================
    
    Set slideLayout = pptGetLayout(layoutName)
    
    If slideIndex > 0 Then
        slideIndex = min(Array(slideIndex, pptCountSlides() + 1), True)
    Else
        slideIndex = pptCountSlides() + 1
    End If
    

    Set ppSlide = ppPres.Slides.AddSlide(slideIndex, slideLayout)
    
End Function


Public Function pptSlideTitle(titleLine1 As String, Optional fontSizeLine1 As Long, Optional titleLine2 As String, Optional fontSizeLine2 As Long)

    If fontSizeLine1 = 0 Then: fontSizeLine1 = 36
    
    With ppSlide.Shapes.Title.TextFrame.TextRange
        With .InsertAfter(titleLine1)
            .Font.Size = fontSizeLine1
        End With
        
        If titleLine2 <> "" Then
            .InsertAfter (vbNewLine)
            If fontSizeLine2 = 0 Then: fontSizeLine2 = 24
            With .InsertAfter(titleLine2)
                .Font.Size = fontSizeLine2
            End With
        End If
    End With
    
        
End Function


Public Function pptSave(saveLocation As String, fileName As String, Optional deleteWorkingCopy As Boolean)

    '===================================================
    Dim fullSavePath As String
    Dim workingCopyPath As String
    Dim i As Integer
    '===================================================
    
    'fullSavePath = saveLocation & "\" & fileName & ".pptx"
    workingCopyPath = ppPres.FullName
    
    'If there is an active presentation, save it
   If Not (ppPres Is Nothing) Then
        fileName = uniqueFileName(saveLocation, fileName, "pptx")
        fullSavePath = saveLocation & "\" & fileName
        ppPres.SaveAs (fullSavePath)
        'Deletes the working copy of the presentation
        If deleteWorkingCopy = True Then: Kill workingCopyPath
    End If
    
End Function


Public Function pptCountSlides() As Long

    pptCountSlides = ppPres.Slides.Count
    
End Function


Public Function pptGetLayout(Optional layoutName As String) As CustomLayout
'Get the layout of the template file

    '==============================================
    Dim cl As CustomLayout
    '==============================================
    
    'Search through each layout type contained in the template
    For Each cl In ppPres.SlideMaster.CustomLayouts
        'If the layout type matches the type requested, apply it to the current slide
        If LCase(cl.name) = LCase(layoutName) Then
            Set pptGetLayout = cl
        End If
    Next cl
    
    'If no matching layout is found, apply the default layout to the slide
    If pptGetLayout Is Nothing Then
        Set pptGetLayout = ppPres.SlideMaster.CustomLayouts(1)
    End If

End Function


Public Function pptPaste(Optional positionLeft As Double, Optional positionTop As Double, Optional shapeHeight As Double, Optional shapeWidth As Double)

    On Error Resume Next
    
    'Paste as an image
    Set ppShape = ppSlide.Shapes.PasteSpecial(2)
    'Turn off aspect ratio locking so that height and width can be adjusted independently
    ppShape.LockAspectRatio = msoFalse
    
    '***********************Need to add error checking and stuff here
    '**************************************************************
    'Set the position
'xxxxxxxxxxxxx add in pixels and inches here eventually xxxxxxxxxxxxxxxxx
    If positionLeft <> 0 Then: ppShape.Left = positionLeft
    If positionTop <> 0 Then: ppShape.Top = positionTop

    'Set the size
    If shapeHeight > 0 Then: ppShape.Height = shapeHeight
    If shapeWidth > 0 Then: ppShape.Width = shapeWidth

'''''    If positionLeft <> 0 Then: ppShape.Left = inchesToPixels(positionLeft)
'''''    If positionTop <> 0 Then: ppShape.Top = inchesToPixels(positionTop)
'''''
'''''    'Set the size
'''''    If shapeHeight > 0 Then: ppShape.Height = inchesToPixels(shapeHeight)
'''''    If shapeWidth > 0 Then: ppShape.Width = inchesToPixels(shapeWidth)

    'Clears the clipboard
    Application.CutCopyMode = False
    
End Function


Public Function inchesToPixels(inches As Double) As Double

    Dim slideInchesH As Double, slideInchesW As Double
    Dim slidePixelsH As Double, slidePixelsW As Double
    
    slidePixelsH = 720
    slidePixelsW = 960
    slideInchesH = 7.5 'ppPres.PageSetup.slideheight
    slideInchesW = 10 'ppPres.PageSetup.slideWidth
    
    inchesToPixels = inches * (slidePixelsH / slideInchesH)

End Function


Public Function pptAlignCenter(Optional horizontal As Boolean, Optional vertical As Boolean)

    '======================================================
    Dim slideH As Double, slideW As Double, shapeH As Double, shapeW As Double
    Dim shapeLeft As Double, shapeTop As Double
    
    On Error Resume Next
    '======================================================
    
    If horizontal = False And vertical = False Then
        horizontal = True
    End If
    

    If horizontal = True Then
        slideW = ppPres.PageSetup.slideWidth
        shapeW = ppShape.Width
        shapeLeft = (slideW - shapeW) / 2
        ppShape.Left = shapeLeft
    End If
    
    If vertical = True Then
        slideH = ppPres.PageSetup.slideheight
        shapeH = ppShape.Height
        shapeTop = (slideH - shapeH) / 2
        ppShape.Top = shapeTop
    End If

End Function



Public Function pptResize(Optional newHeight As String, Optional newWidth As String, Optional maintainRatio As Boolean)
    
    '=================================================================
    Dim heightMultiplier As Double, widthMultiplier As Double
    
    On Error Resume Next
    '=================================================================
    
    '******************************************************************
    '*********Can I make the 2 duplicate if statements into a function??
    '******************************************************************
    
    'If resized height is entered as a percent...
    If Len(reMatch(newHeight, "^[0-9]+[.]?[0-9]+[%]$")) > 0 Then
        '...Remove the percent sign and convert to a decimal (divide by 100)
        heightMultiplier = CDbl(reMatch(newHeight, "[0-9]+[.]?[0-9]+")) / 100
    'If resize is entered as a number, convert to a decimal - the ratio of new size to current size
    ElseIf Len(reMatch(newHeight, "^[0-9]+[.]?[0-9]+$")) Then
        heightMultiplier = CDbl(newHeight) / ppShape.Height
    End If
    
    'If resized width is entered as a percent...
    If Len(reMatch(newWidth, "^[0-9]+[.]?[0-9]+[%]$")) > 0 Then
        'Remove the percent sign and convert to a decimal (divide by 100)
        widthMultiplier = CDbl(reMatch(newWidth, "[0-9]+[.]?[0-9]+")) / 100
    'If resize is entered as a number, convert to a decimal - the ratio of new size to current size
    ElseIf Len(reMatch(newWidth, "^[0-9]+[.]?[0-9]+$")) Then
        widthMultiplier = CDbl(newWidth) / ppShape.Width
    End If
    

    'If a new height but NOT new width was entered
    If heightMultiplier > 0 And widthMultiplier = 0 Then
    
        'If aspect ratio should be maintained, adjust the width accordingly
        If maintainRatio = True Then
            ppShape.Width = ppShape.Width * heightMultiplier
        End If
        'Multiply the current height by the height multiplier
        ppShape.Height = ppShape.Height * heightMultiplier

    'If a new width but NOT a new height was entered
    ElseIf heightMultiplier = 0 And widthMultiplier > 0 Then
    
        'If aspect ratio should be maintained, adjust the width accordingly
        If maintainRatio = True Then
            ppShape.Height = ppShape.Height * widthMultiplier
        End If
        'Multiply the current height by the height multiplier
        ppShape.Width = ppShape.Width * widthMultiplier
        
    'If both new height and width were entered, don't maintain aspect ratio (assume user entered True by accident)
    ElseIf heightMultiplier > 0 And widthMultiplier > 0 Then
        ppShape.Height = ppShape.Height * heightMultiplier
        ppShape.Width = ppShape.Width * widthMultiplier
    End If
    

End Function
