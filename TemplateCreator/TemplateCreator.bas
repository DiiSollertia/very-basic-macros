Attribute VB_Name = "Module2"
Sub CreateDocVarAndUpdateFields()
On Error GoTo ErrorHandler
'Connect using Early Binding.
'Remember to set the reference to the Word Object Library
'In VBE Editor Tools -> References -> Microsoft Word x.xx Object Library
Dim WordApp As Object, WordDoc As Object, Sheet As Object
'Initiate Word session
Set WordApp = CreateObject("Word.Application")
'New Apps will be hidden by default, so make visible if needed for debugging
WordApp.Visible = True
Set Sheet = Worksheets("Utility")
Debug.Print "Debug String 1: " & Sheet.Range("D2") 'Debug String 1
Set WordDoc = WordApp.Documents.Open(Sheet.Range("D2").Value)
Debug.Print "Debug String 2: " & TypeName(WordDoc) 'Debug String 2

With WordDoc
'Clear all pre-existing DocumentVariables
Dim i As Long
Debug.Print "Debug String 3: " & .Variables.Count 'Debug String 3
For i = .Variables.Count To 1 Step -1
Debug.Print "Debug String 4: " & .Variables(i).Name & "=" & .Variables(i).Value 'Debug String 4
    .Variables(i).Delete
Next i

'Initiate DocVar with Values provided
i = 2
For Each DocVar In Sheet.Range("A2", Sheet.Range("A2").End(xlDown))
    Debug.Print "Debug String 5: " & DocVar & "=" & Sheet.Range("B" & i).Value 'Debug String 5
    WordDoc.Variables(DocVar).Value = Sheet.Range("B" & i).Value
    i = i + 1
Next

'Update all fields and references in document regardless of location
'=========================
'Macro created 2019 by Lene Fredborg, DocTools - www.thedoctools.com
'Revised August 2020 by Lene Fredborg
'THIS MACRO IS COPYRIGHT. YOU ARE WELCOME TO USE THE MACRO BUT YOU MUST KEEP THE LINE ABOVE.
'YOU ARE NOT ALLOWED TO PUBLISH THE MACRO AS YOUR OWN, IN WHOLE OR IN PART.
'=========================
'The macro updates all fields in the activedocument no matter where the fields are found
'Includes fields in headers, footers, footnotes, endnotes, shapes, etc.
'=========================
'Modified by xinzhen.khoo to be deployed from Excel document
Dim rngStory As Word.Range
Dim oShape As Word.Shape
Dim oShape_2 As Word.Shape
Dim oTOC As Word.TableOfContents
Dim oTOF As Word.TableOfFigures
Dim oTOA As Word.TableOfAuthorities

'Turn off screen updating for better performance
WordApp.ScreenUpdating = False

'Prevent alert when updating footnotes/endnotes/comments story
WordApp.DisplayAlerts = Word.wdAlertsNone

'Iterate through all stories and update fields
For Each rngStory In .StoryRanges
    If Not rngStory Is Nothing Then
        'Update fields directly in story
        rngStory.Fields.Update
        
        If rngStory.StoryType <> wdMainTextStory Then
            'Update fields in shapes and drawing canvases with shapes
            For Each oShape In rngStory.ShapeRange
                With oShape.TextFrame
                    If .HasText Then
                        .TextRange.Fields.Update
                    End If
                    
                    'In case of a drawing canvas
                    'May contain other shapes that may contain fields
                    If oShape.Type = msoCanvas Then
                        For Each oShape_2 In oShape.CanvasItems
                            With oShape_2.TextFrame
                                If .HasText Then
                                    .TextRange.Fields.Update
                                End If
                            End With
                        Next oShape_2
                    End If
                    
                End With
            Next oShape
        End If

        'Handle e.g. multiple sections with unlinked headers/footers or linked text boxes
        If rngStory.StoryType <> wdMainTextStory Then
            While Not (rngStory.NextStoryRange Is Nothing)
                Set rngStory = rngStory.NextStoryRange
                rngStory.Fields.Update
            
                'Update fields in shapes and drawing canvases with shapes
                For Each oShape In rngStory.ShapeRange
                    With oShape.TextFrame
                        If .HasText Then
                            .TextRange.Fields.Update
                        End If
                        
                        'In case of a drawing canvas
                        'May contain other shapes that may contain fields
                        If oShape.Type = msoCanvas Then
                            For Each oShape_2 In oShape.CanvasItems
                                With oShape_2.TextFrame
                                    If .HasText Then
                                        .TextRange.Fields.Update
                                    End If
                                End With
                            Next oShape_2
                        End If
                        
                    End With
                Next oShape
            Wend
        End If
    End If
Next rngStory
'=========================

'Update any TOC, TOF, TOA
For Each oTOC In .TablesOfContents
    oTOC.Update
Next oTOC
For Each oTOF In .TablesOfFigures
    oTOF.Update
Next oTOF
For Each oTOA In .TablesOfAuthorities
    oTOA.Update
Next oTOA
'=========================

MsgBox "Task completed, closing Word Document..."
'Close application and release memory
.Close
End With

Application.ScreenUpdating = True
Application.DisplayAlerts = Word.wdAlertsAll
WordApp.Quit
Set rngStory = Nothing
Set WordDoc = Nothing
Set WordApp = Nothing
Exit Sub

'Error Handling
ErrorHandler:
If Err.Number <> 0 Then
    Msg = "Error # " & Str(Err.Number) & " was generated by " _
    & Err.Source & Chr(13) & "Error Line: " & Erl & Chr(13) & Err.Description
    MsgBox Msg, , "Error", Err.HelpFile, Err.HelpContext
End If
Resume Next
End Sub
