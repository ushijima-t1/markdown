Attribute VB_Name = "SlideGenerator"
'==============================================================================
' PowerPoint Slide Generation System - Clean Implementation
' Uses JsonConverter.bas for reliable JSON parsing
'
' REQUIRED REFERENCE: Microsoft Scripting Runtime
'==============================================================================

Option Explicit

' Global variables
Private processLog As Collection

'------------------------------------------------------------------------------
' Main entry point - Generate slides from JSON data
'------------------------------------------------------------------------------
Public Sub GenerateSlidesFromJSON()
    On Error GoTo ErrorHandler
    
    ' Initialize
    Set processLog = New Collection
    AddLog "Process started: " & ActivePresentation.Name
    
    ' 1. Get JSON data from slide 1 notes
    Dim jsonData As String
    jsonData = GetJSONFromSlideNotes()
    
    If Len(jsonData) = 0 Then
        MsgBox "No JSON data found in slide 1 notes", vbExclamation
        Exit Sub
    End If
    
    AddLog "JSON data retrieved (" & Len(jsonData) & " characters)"
    
    ' 2. Parse JSON using JsonConverter
    Dim slideDataArray As Object
    Set slideDataArray = JsonConverter.ParseJson(jsonData)
    
    If slideDataArray Is Nothing Then
        MsgBox "JSON parsing failed. Please check JSON format.", vbCritical
        Exit Sub
    End If
    
    AddLog "JSON parsed successfully. Type: " & TypeName(slideDataArray)
    
    ' 3. Process slides based on data type
    If TypeName(slideDataArray) = "Collection" Then
        ProcessSlideCollection slideDataArray
    Else
        MsgBox "Expected JSON array format, got: " & TypeName(slideDataArray), vbExclamation
        Exit Sub
    End If
    
    AddLog "All slides generated successfully"
    MsgBox "Slide generation completed", vbInformation
    Exit Sub
    
ErrorHandler:
    AddLog "Critical error: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

'------------------------------------------------------------------------------
' Get JSON data from slide 1 notes
'------------------------------------------------------------------------------
Private Function GetJSONFromSlideNotes() As String
    On Error GoTo ErrorHandler
    
    Dim pres As Presentation
    Dim firstSlide As Slide
    Dim notesText As String
    
    Set pres = ActivePresentation
    AddLog "Starting JSON retrieval from slide notes..."
    
    If pres.Slides.Count = 0 Then
        AddLog "Error: No slides found in presentation"
        GetJSONFromSlideNotes = ""
        Exit Function
    End If
    
    AddLog "Total slides in presentation: " & pres.Slides.Count
    Set firstSlide = pres.Slides(1)
    AddLog "Successfully accessed slide 1"
    
    ' Check if slide has notes page
    On Error Resume Next
    Dim hasNotes As Boolean
    hasNotes = firstSlide.HasNotesPage
    If Err.Number <> 0 Then
        AddLog "Error checking HasNotesPage: " & Err.Description
        Err.Clear
        hasNotes = False
    End If
    On Error GoTo ErrorHandler
    
    If hasNotes Then
        AddLog "Slide has notes page"
        
        ' Get notes page and check shapes
        On Error Resume Next
        Dim notesPage As SlideRange
        Set notesPage = firstSlide.NotesPage
        If Err.Number <> 0 Then
            AddLog "Error getting NotesPage: " & Err.Description
            Err.Clear
        Else
            AddLog "Notes page shapes count: " & notesPage.Shapes.Count
            
            ' Look for placeholder with PlaceholderFormat.Type = ppPlaceholderBody
            AddLog "Searching for notes placeholder..."
            Dim shape As Shape
            Dim shapeIndex As Long
            shapeIndex = 1
            
            For Each shape In notesPage.Shapes
                AddLog "Checking shape " & shapeIndex & " - Type: " & shape.Type
                If shape.Type = msoPlaceholder Then
                    AddLog "Shape " & shapeIndex & " is a placeholder"
                    If shape.PlaceholderFormat.Type = ppPlaceholderBody Then
                        AddLog "Found notes body placeholder at shape " & shapeIndex
                        If shape.HasTextFrame Then
                            AddLog "Notes placeholder has text frame"
                            If shape.TextFrame.HasText Then
                                notesText = shape.TextFrame.TextRange.Text
                                AddLog "Successfully retrieved notes text from body placeholder"
                                Exit For
                            Else
                                AddLog "Notes placeholder text frame is empty"
                            End If
                        Else
                            AddLog "Notes placeholder has no text frame"
                        End If
                    Else
                        AddLog "Shape " & shapeIndex & " placeholder type: " & shape.PlaceholderFormat.Type & " (not body)"
                    End If
                End If
                
                If Err.Number <> 0 Then
                    AddLog "Error processing shape " & shapeIndex & ": " & Err.Description
                    Err.Clear
                End If
                shapeIndex = shapeIndex + 1
            Next shape
            
            ' Fallback: try any shape with text if no placeholder found
            If Len(notesText) = 0 Then
                AddLog "No body placeholder found, trying any text shape..."
                For Each shape In notesPage.Shapes
                    If shape.HasTextFrame Then
                        If shape.TextFrame.HasText Then
                            Dim testText As String
                            testText = shape.TextFrame.TextRange.Text
                            If Len(Trim(testText)) > 10 Then
                                notesText = testText
                                AddLog "Found notes text using fallback method"
                                Exit For
                            End If
                        End If
                    End If
                Next shape
            End If
        End If
        On Error GoTo ErrorHandler
    Else
        AddLog "Slide has no notes page"
    End If
    
    AddLog "Notes text length: " & Len(notesText)
    
    ' Simply return the notes text - let JsonConverter handle the parsing
    GetJSONFromSlideNotes = notesText
    
    Exit Function
    
ErrorHandler:
    GetJSONFromSlideNotes = ""
End Function

'------------------------------------------------------------------------------
' Process slide collection from JSON
'------------------------------------------------------------------------------
Private Sub ProcessSlideCollection(slideDataArray As Object)
    On Error GoTo ErrorHandler
    
    Dim slideCount As Long
    slideCount = slideDataArray.Count
    AddLog "Processing " & slideCount & " slides"
    
    Dim i As Long
    For i = 1 To slideCount
        AddLog "--- Processing slide " & i & " of " & slideCount & " ---"
        
        Dim slideData As Object
        Set slideData = slideDataArray(i)
        
        If Not slideData Is Nothing Then
            ProcessSingleSlide slideData, i
        Else
            AddLog "Warning: Slide data " & i & " is Nothing"
        End If
    Next i
    
    Exit Sub
    
ErrorHandler:
    AddLog "Error in ProcessSlideCollection: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Process single slide data
'------------------------------------------------------------------------------
Private Sub ProcessSingleSlide(slideData As Object, slideNumber As Long)
    On Error GoTo ErrorHandler
    
    ' Get slide_id and slide_content
    Dim slideId As Variant
    Dim slideContent As Object
    
    slideId = slideData("slide_id")
    Set slideContent = slideData("slide_content")
    
    AddLog "Slide " & slideNumber & " - Template ID: " & slideId
    
    ' Validate slide_id
    If IsEmpty(slideId) Or Not IsNumeric(slideId) Then
        AddLog "Error: Invalid slide_id for slide " & slideNumber
        Exit Sub
    End If
    
    If slideContent Is Nothing Then
        AddLog "Error: No slide_content for slide " & slideNumber
        Exit Sub
    End If
    
    ' Create new slide from template
    Dim newSlide As Slide
    Set newSlide = CreateSlideFromTemplate(CLng(slideId))
    
    If Not newSlide Is Nothing Then
        ' Replace variables in the new slide
        ReplaceVariablesInSlide newSlide, slideContent
        AddLog "Slide " & slideNumber & " created successfully"
    Else
        AddLog "Error: Failed to create slide " & slideNumber
    End If
    
    Exit Sub
    
ErrorHandler:
    AddLog "Error processing slide " & slideNumber & ": " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Create slide from template using Duplicate + MoveTo
'------------------------------------------------------------------------------
Private Function CreateSlideFromTemplate(templateId As Long) As Slide
    On Error GoTo ErrorHandler
    
    Dim pres As Presentation
    Set pres = ActivePresentation
    
    ' Validate template exists
    If templateId < 1 Or templateId > pres.Slides.Count Then
        AddLog "Error: Template slide " & templateId & " does not exist"
        Set CreateSlideFromTemplate = Nothing
        Exit Function
    End If
    
    ' Duplicate template slide
    Dim templateSlide As Slide
    Set templateSlide = pres.Slides(templateId)
    
    Dim newSlideRange As SlideRange
    Set newSlideRange = templateSlide.Duplicate
    
    ' Move to end of presentation
    newSlideRange.MoveTo pres.Slides.Count
    
    ' Refresh presentation reference and return the last slide
    Set pres = ActivePresentation
    Set CreateSlideFromTemplate = pres.Slides(pres.Slides.Count)
    
    AddLog "Template " & templateId & " duplicated to position " & pres.Slides.Count
    Exit Function
    
ErrorHandler:
    AddLog "Error creating slide from template " & templateId & ": " & Err.Description
    Set CreateSlideFromTemplate = Nothing
End Function

'------------------------------------------------------------------------------
' Replace variables in slide content
'------------------------------------------------------------------------------
Private Sub ReplaceVariablesInSlide(slide As Slide, content As Object)
    On Error GoTo ErrorHandler
    
    AddLog "Starting variable replacement"
    
    ' Process all shapes on the slide
    Dim shape As Shape
    Dim shapeCount As Long
    shapeCount = 0
    
    For Each shape In slide.Shapes
        shapeCount = shapeCount + 1
        
        ' Check if shape is a group
        If shape.Type = msoGroup Then
            ProcessGroupedShapes shape, content
            AddLog "Shape " & shapeCount & " is a group - processing group items"
        ' Process regular text frames
        ElseIf shape.HasTextFrame Then
            If shape.TextFrame.HasText Then
                Dim originalText As String
                originalText = shape.TextFrame.TextRange.Text
                
                ' Check if text contains any _list placeholders
                If InStr(originalText, "_list") > 0 Then
                    AddLog "Shape " & shapeCount & " contains _list placeholder: " & Left(originalText, 50) & "..."
                End If
                
                Dim newText As String
                newText = ReplaceVariablesInText(originalText, content)
                
                If newText <> originalText Then
                    ' Save original font properties before text replacement
                    Dim originalFontName As String
                    Dim originalFontSize As Single
                    Dim originalFontBold As Boolean
                    Dim originalFontItalic As Boolean
                    Dim originalFontColor As Long
                    
                    On Error Resume Next
                    With shape.TextFrame.TextRange.Font
                        originalFontName = .Name
                        originalFontSize = .Size
                        originalFontBold = .Bold
                        originalFontItalic = .Italic
                        originalFontColor = .Color.RGB
                    End With
                    On Error GoTo ErrorHandler
                    
                    ' Replace text
                    shape.TextFrame.TextRange.Text = newText
                    
                    ' Restore original font properties
                    On Error Resume Next
                    With shape.TextFrame.TextRange.Font
                        .Name = originalFontName
                        .Size = originalFontSize
                        .Bold = originalFontBold
                        .Italic = originalFontItalic
                        .Color.RGB = originalFontColor
                    End With
                    On Error GoTo ErrorHandler
                    
                    AddLog "Shape " & shapeCount & " text updated with font preservation"
                Else
                    ' Only log if text contains variables but wasn't updated
                    If InStr(originalText, "{") > 0 Then
                        AddLog "Shape " & shapeCount & " has variables but no replacement: " & Left(originalText, 50) & "..."
                    End If
                End If
            Else
                AddLog "Shape " & shapeCount & " has text frame but no text"
            End If
        Else
            AddLog "Shape " & shapeCount & " has no text frame (Type: " & shape.Type & ")"
        End If
        
        ' Process SmartArt (skip if already processed in group)
        If shape.Type <> msoGroup And shape.HasSmartArt Then
            ProcessSmartArtVariables shape.SmartArt, content
            AddLog "Shape " & shapeCount & " SmartArt processed"
        End If
    Next shape
    
    ' Process tables if any (search through shapes for tables)
    On Error Resume Next
    Dim tableShape As Shape
    For Each tableShape In slide.Shapes
        If tableShape.HasTable Then
            ProcessTableVariables tableShape.Table, content
        End If
    Next tableShape
    On Error GoTo ErrorHandler
    
    AddLog "Variable replacement completed"
    Exit Sub
    
ErrorHandler:
    AddLog "Error in variable replacement: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Replace variables in text using content dictionary
'------------------------------------------------------------------------------
Private Function ReplaceVariablesInText(text As String, content As Object) As String
    Dim result As String
    result = text
    
    ' Try to access content properties
    On Error Resume Next
    
    ' Check if it's a Scripting.Dictionary
    If TypeName(content) = "Dictionary" Then
        ' Process Dictionary keys
        Dim key As Variant
        For Each key In content.Keys
            Dim placeholder As String
            placeholder = "{" & key & "}"
            
            ' Replace placeholder with value if it exists in text
            If InStr(result, placeholder) > 0 Then
                Dim value As Variant
                value = content(key)
                
                If Err.Number = 0 And Not IsEmpty(value) Then
                    ' Check if this is a _list variable
                    If InStr(key, "_list") > 0 Then
                        ' Process comma-separated string as bullet list
                        Dim listString As String
                        listString = CStr(value)
                        
                        ' Debug: Log the raw string content
                        AddLog "DEBUG: Raw _list string: '" & listString & "'"
                        AddLog "DEBUG: String length: " & Len(listString)
                        
                        ' Split by comma and create bullet list
                        Dim items() As String
                        items = Split(listString, ",")
                        
                        AddLog "DEBUG: Split into " & (UBound(items) + 1) & " items"
                        
                        Dim bulletText As String
                        Dim i As Long
                        For i = 0 To UBound(items)
                            If i > 0 Then bulletText = bulletText & vbCrLf
                            
                            ' Clean the item text
                            Dim cleanItem As String
                            cleanItem = Trim(items(i))
                            
                            ' Debug: Log each item before and after cleaning
                            AddLog "DEBUG: Item " & i & " before cleaning: '" & items(i) & "'"
                            
                            ' Remove various bullet characters and corrupted characters
                            cleanItem = Replace(cleanItem, "窶｢", "")  ' Remove corrupted bullet
                            cleanItem = Replace(cleanItem, "窶", "")   ' Remove corrupted character
                            cleanItem = Replace(cleanItem, "｢", "")    ' Remove corrupted bracket
                            cleanItem = Replace(cleanItem, "・", "")   ' Remove Japanese bullet
                            cleanItem = Replace(cleanItem, "•", "")    ' Remove bullet
                            cleanItem = Replace(cleanItem, "◦", "")    ' Remove hollow bullet
                            cleanItem = Replace(cleanItem, "‣", "")    ' Remove triangular bullet
                            cleanItem = Replace(cleanItem, "▪", "")    ' Remove square bullet
                            cleanItem = Replace(cleanItem, "▫", "")    ' Remove hollow square bullet
                            cleanItem = Replace(cleanItem, Chr(226) & Chr(128) & Chr(162), "")  ' UTF-8 bullet
                            cleanItem = Replace(cleanItem, vbTab, "")  ' Remove tabs
                            cleanItem = Replace(cleanItem, Chr(160), " ")  ' Replace non-breaking space with regular space
                            
                            ' Remove leading/trailing spaces and multiple spaces
                            Do While InStr(cleanItem, "  ") > 0
                                cleanItem = Replace(cleanItem, "  ", " ")
                            Loop
                            cleanItem = Trim(cleanItem)
                            
                            AddLog "DEBUG: Item " & i & " after cleaning: '" & cleanItem & "'"
                            
                            If Len(cleanItem) > 0 Then
                                bulletText = bulletText & cleanItem  ' No bullet character - PowerPoint will add it
                            End If
                        Next i
                        
                        result = Replace(result, placeholder, bulletText)
                        AddLog "Replaced " & placeholder & " with bullet list (" & (UBound(items) + 1) & " items)"
                    Else
                        ' Normal variable replacement
                        If Not IsObject(value) Then
                            result = Replace(result, placeholder, CStr(value))
                        End If
                    End If
                End If
                Err.Clear
            End If
        Next key
    Else
        ' For non-Dictionary objects, try to dynamically detect placeholders in text
        
        ' Find all placeholders in text
        Dim startPos As Long
        Dim endPos As Long
        Dim varName As String
        Dim searchPos As Long
        searchPos = 1
        
        Do
            startPos = InStr(searchPos, result, "{")
            If startPos > 0 Then
                endPos = InStr(startPos, result, "}")
                If endPos > startPos Then
                    ' Extract variable name
                    varName = Mid(result, startPos + 1, endPos - startPos - 1)
                    
                    
                    ' Try to get value from content
                    Dim varValue As Variant
                    varValue = content(varName)
                    
                    If Err.Number = 0 And Not IsEmpty(varValue) Then
                        Dim varPlaceholder As String
                        varPlaceholder = "{" & varName & "}"
                        
                        ' Check if this is a _list variable
                        If InStr(varName, "_list") > 0 Then
                            ' Process comma-separated string as bullet list
                            Dim listStringFallback As String
                            listStringFallback = CStr(varValue)
                            
                            ' Split by comma and create bullet list
                            Dim itemsFallback() As String
                            itemsFallback = Split(listStringFallback, ",")
                            
                            Dim bulletTextFallback As String
                            Dim j As Long
                            For j = 0 To UBound(itemsFallback)
                                If j > 0 Then bulletTextFallback = bulletTextFallback & vbCrLf
                                
                                ' Clean the item text
                                Dim cleanItemFallback As String
                                cleanItemFallback = Trim(itemsFallback(j))
                                
                                ' Remove various bullet characters and corrupted characters
                                cleanItemFallback = Replace(cleanItemFallback, "窶｢", "")  ' Remove corrupted bullet
                                cleanItemFallback = Replace(cleanItemFallback, "窶", "")   ' Remove corrupted character
                                cleanItemFallback = Replace(cleanItemFallback, "｢", "")    ' Remove corrupted bracket
                                cleanItemFallback = Replace(cleanItemFallback, "・", "")   ' Remove Japanese bullet
                                cleanItemFallback = Replace(cleanItemFallback, "•", "")    ' Remove bullet
                                cleanItemFallback = Replace(cleanItemFallback, "◦", "")    ' Remove hollow bullet
                                cleanItemFallback = Replace(cleanItemFallback, "‣", "")    ' Remove triangular bullet
                                cleanItemFallback = Replace(cleanItemFallback, "▪", "")    ' Remove square bullet
                                cleanItemFallback = Replace(cleanItemFallback, "▫", "")    ' Remove hollow square bullet
                                cleanItemFallback = Replace(cleanItemFallback, Chr(226) & Chr(128) & Chr(162), "")  ' UTF-8 bullet
                                cleanItemFallback = Replace(cleanItemFallback, vbTab, "")  ' Remove tabs
                                cleanItemFallback = Replace(cleanItemFallback, Chr(160), " ")  ' Replace non-breaking space with regular space
                                
                                ' Remove leading/trailing spaces and multiple spaces
                                Do While InStr(cleanItemFallback, "  ") > 0
                                    cleanItemFallback = Replace(cleanItemFallback, "  ", " ")
                                Loop
                                cleanItemFallback = Trim(cleanItemFallback)
                                
                                If Len(cleanItemFallback) > 0 Then
                                    bulletTextFallback = bulletTextFallback & cleanItemFallback  ' No bullet character - PowerPoint will add it
                                End If
                            Next j
                            
                            result = Replace(result, varPlaceholder, bulletTextFallback)
                            AddLog "Replaced " & varPlaceholder & " with bullet list (" & (UBound(itemsFallback) + 1) & " items)"
                        Else
                            ' Normal variable replacement
                            If Not IsObject(varValue) Then
                                result = Replace(result, varPlaceholder, CStr(varValue))
                            End If
                        End If
                    End If
                    
                    Err.Clear
                    searchPos = endPos + 1
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop
    End If
    
    On Error GoTo 0
    
    ReplaceVariablesInText = result
End Function

'------------------------------------------------------------------------------
' Process table variables
'------------------------------------------------------------------------------
Private Sub ProcessTableVariables(table As Object, content As Object)
    On Error Resume Next
    
    Dim row As Long
    Dim col As Long
    
    For row = 1 To table.Rows.Count
        For col = 1 To table.Columns.Count
            Dim cell As Object
            Set cell = table.Cell(row, col)
            
            If cell.Shape.HasTextFrame Then
                If cell.Shape.TextFrame.HasText Then
                    Dim originalText As String
                    originalText = cell.Shape.TextFrame.TextRange.Text
                    
                    Dim newText As String
                    newText = ReplaceVariablesInText(originalText, content)
                    
                    If newText <> originalText Then
                        ' Save original font properties for table cell
                        Dim cellFontName As String
                        Dim cellFontSize As Single
                        Dim cellFontBold As Boolean
                        Dim cellFontItalic As Boolean
                        Dim cellFontColor As Long
                        
                        On Error Resume Next
                        With cell.Shape.TextFrame.TextRange.Font
                            cellFontName = .Name
                            cellFontSize = .Size
                            cellFontBold = .Bold
                            cellFontItalic = .Italic
                            cellFontColor = .Color.RGB
                        End With
                        
                        ' Replace text
                        cell.Shape.TextFrame.TextRange.Text = newText
                        
                        ' Restore font properties
                        With cell.Shape.TextFrame.TextRange.Font
                            .Name = cellFontName
                            .Size = cellFontSize
                            .Bold = cellFontBold
                            .Italic = cellFontItalic
                            .Color.RGB = cellFontColor
                        End With
                        On Error GoTo 0
                    End If
                End If
            End If
        Next col
    Next row
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Process SmartArt variables
'------------------------------------------------------------------------------
Private Sub ProcessSmartArtVariables(smartArt As Object, content As Object)
    On Error Resume Next
    
    ' Process all nodes in the SmartArt
    Dim node As Object
    For Each node In smartArt.AllNodes
        ' Process text in each node
        If node.TextFrame2.HasText Then
            Dim originalText As String
            originalText = node.TextFrame2.TextRange.Text
            
            Dim newText As String
            newText = ReplaceVariablesInText(originalText, content)
            
            If newText <> originalText Then
                ' Save original font properties for SmartArt node
                Dim smartFontName As String
                Dim smartFontSize As Single
                Dim smartFontBold As Boolean
                Dim smartFontItalic As Boolean
                
                On Error Resume Next
                With node.TextFrame2.TextRange.Font
                    smartFontName = .Name
                    smartFontSize = .Size
                    smartFontBold = .Bold
                    smartFontItalic = .Italic
                End With
                
                ' Replace text
                node.TextFrame2.TextRange.Text = newText
                
                ' Restore font properties
                With node.TextFrame2.TextRange.Font
                    .Name = smartFontName
                    .Size = smartFontSize
                    .Bold = smartFontBold
                    .Italic = smartFontItalic
                End With
                On Error GoTo 0
                
                AddLog "SmartArt node text updated: " & Left(originalText, 20) & "..."
            End If
        End If
        
        ' Process sub-nodes if they exist
        If node.Nodes.Count > 0 Then
            ProcessSmartArtSubNodes node.Nodes, content
        End If
    Next node
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Process SmartArt sub-nodes recursively
'------------------------------------------------------------------------------
Private Sub ProcessSmartArtSubNodes(nodes As Object, content As Object)
    On Error Resume Next
    
    Dim node As Object
    For Each node In nodes
        ' Process text in each sub-node
        If node.TextFrame2.HasText Then
            Dim originalText As String
            originalText = node.TextFrame2.TextRange.Text
            
            Dim newText As String
            newText = ReplaceVariablesInText(originalText, content)
            
            If newText <> originalText Then
                node.TextFrame2.TextRange.Text = newText
                AddLog "SmartArt sub-node text updated: " & Left(originalText, 20) & "..."
            End If
        End If
        
        ' Process further sub-nodes if they exist
        If node.Nodes.Count > 0 Then
            ProcessSmartArtSubNodes node.Nodes, content
        End If
    Next node
    
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------
' Process grouped shapes recursively
'------------------------------------------------------------------------------
Private Sub ProcessGroupedShapes(groupShape As Shape, content As Object)
    On Error GoTo ErrorHandler
    
    Dim i As Long
    Dim subShape As Shape
    
    AddLog "Processing group with " & groupShape.GroupItems.Count & " items"
    
    ' Process each shape in the group
    For i = 1 To groupShape.GroupItems.Count
        Set subShape = groupShape.GroupItems(i)
        
        ' If this is another group, process recursively
        If subShape.Type = msoGroup Then
            ProcessGroupedShapes subShape, content
            AddLog "Group item " & i & " is a nested group - processing recursively"
        Else
            ' Process text frame if exists
            If subShape.HasTextFrame Then
                If subShape.TextFrame.HasText Then
                    Dim originalText As String
                    originalText = subShape.TextFrame.TextRange.Text
                    
                    ' Check if text contains any _list placeholders
                    If InStr(originalText, "_list") > 0 Then
                        AddLog "Group item " & i & " contains _list placeholder: " & Left(originalText, 50) & "..."
                    End If
                    
                    Dim newText As String
                    newText = ReplaceVariablesInText(originalText, content)
                    
                    If newText <> originalText Then
                        ' Save original font properties before text replacement
                        Dim originalFontName As String
                        Dim originalFontSize As Single
                        Dim originalFontBold As Boolean
                        Dim originalFontItalic As Boolean
                        Dim originalFontColor As Long
                        
                        On Error Resume Next
                        With subShape.TextFrame.TextRange.Font
                            originalFontName = .Name
                            originalFontSize = .Size
                            originalFontBold = .Bold
                            originalFontItalic = .Italic
                            originalFontColor = .Color.RGB
                        End With
                        On Error GoTo ErrorHandler
                        
                        ' Replace text
                        subShape.TextFrame.TextRange.Text = newText
                        
                        ' Restore original font properties
                        On Error Resume Next
                        With subShape.TextFrame.TextRange.Font
                            .Name = originalFontName
                            .Size = originalFontSize
                            .Bold = originalFontBold
                            .Italic = originalFontItalic
                            .Color.RGB = originalFontColor
                        End With
                        On Error GoTo ErrorHandler
                        
                        AddLog "Group item " & i & " text updated with font preservation"
                    Else
                        ' Only log if text contains variables but wasn't updated
                        If InStr(originalText, "{") > 0 Then
                            AddLog "Group item " & i & " has variables but no replacement: " & Left(originalText, 50) & "..."
                        End If
                    End If
                Else
                    AddLog "Group item " & i & " has text frame but no text"
                End If
            Else
                AddLog "Group item " & i & " has no text frame (Type: " & subShape.Type & ")"
            End If
            
            ' Process SmartArt in grouped shape
            If subShape.HasSmartArt Then
                ProcessSmartArtVariables subShape.SmartArt, content
                AddLog "Group item " & i & " SmartArt processed"
            End If
            
            ' Process tables in grouped shape
            On Error Resume Next
            If subShape.HasTable Then
                ProcessTableVariables subShape.Table, content
                AddLog "Group item " & i & " table processed"
            End If
            On Error GoTo ErrorHandler
        End If
    Next i
    
    AddLog "Finished processing group"
    Exit Sub
    
ErrorHandler:
    AddLog "Error processing grouped shapes: " & Err.Description
End Sub

'------------------------------------------------------------------------------
' Add log entry
'------------------------------------------------------------------------------
Private Sub AddLog(message As String)
    Dim timestamp As String
    Dim logMessage As String
    
    timestamp = Format(Now, "hh:mm:ss")
    logMessage = "[" & timestamp & "] " & message
    
    Debug.Print logMessage
    processLog.Add logMessage
End Sub

'------------------------------------------------------------------------------
' Show process log
'------------------------------------------------------------------------------
Public Sub ShowProcessLog()
    If processLog Is Nothing Or processLog.Count = 0 Then
        MsgBox "No log entries available", vbInformation
        Exit Sub
    End If
    
    Dim logText As String
    Dim i As Long
    Dim startIndex As Long
    
    ' Show latest 20 entries
    startIndex = IIf(processLog.Count > 20, processLog.Count - 19, 1)
    
    logText = "Process Log (Latest 20 entries):" & vbNewLine & vbNewLine
    For i = startIndex To processLog.Count
        logText = logText & processLog(i) & vbNewLine
    Next i
    
    MsgBox logText, vbInformation, "Process Log"
End Sub