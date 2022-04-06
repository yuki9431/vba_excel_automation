' Excel | Move to A1 and display zoom 100%
Sub moveA1_zoom100()
    Dim sheet As Worksheet

    For Each sheet In Worksheet
        sheet.Activate
        sheet.Range("A1").Activate
        ActiveWindow.Zoom = 100
    Next
Sheets(1).Activate
End Sub
