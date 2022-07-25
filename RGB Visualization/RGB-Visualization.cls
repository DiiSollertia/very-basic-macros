VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo bm_Safe_Exit
    Application.EnableEvents = False
    Dim rng As Range, clr As String
    For Each rng In Target
        'Format of values = XXX,XXX,XXX
        If Len(rng.Value2) = 11 Then
            clr = rng.Value2
            'Change color based on RGB values
            rng.Interior.Color = _
              RGB(Left(clr, 3), _
                  Mid(clr, 5, 3), _
                  Right(clr, 3))
        ElseIf Len(rng.Value2) = 0 Then
            'Reset color after deleting content
            rng.ClearFormats
        End If
    Next rng

bm_Safe_Exit:
    Application.EnableEvents = True
End Sub

