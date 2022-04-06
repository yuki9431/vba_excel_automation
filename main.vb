' Excel | Move to A1 and display zoom 100%
Sub moveA1_zoom100()
    Dim sheet As Worksheet
    Set sheet = ActiveSheet
    For Each sheet In Worksheet
        sheet.Activate
        ActiveWindow.Zoom = 100
        sheet.Range("A1").Select
    Next
Sheets(1).Select
End Sub
