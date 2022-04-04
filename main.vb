' Excel | Move to A1 and display zoom 100%
Sub moveA1_zoom100()
    Dim cell As Worksheet
    Dim sheet As Worksheet

    Set sheet = ActiveSheet

    For Each cell In Worksheet
        cell.Select
        ActiveWindow.Zoom = 100
        cell.Range("A1").Select
    Next
Sheets(1).Select
End Sub
