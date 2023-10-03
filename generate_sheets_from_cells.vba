Sub Addsheetsfromselection()
    Dim CurSheet As Worksheet
    Dim Source As Range
    Dim c As Range
    
    Set CurSheet = ActiveSheet
    Set Source = Selection.Cells
    Application.ScreenUpdating = False
    
    For Each c In Source
        sName = Trim(c.Text)
        If Len(sName) > 0 Then
            Worksheets.Add After:=Worksheets(Worksheets.Count)
            ActiveSheet.Name = sName
        End If
    Next c
    CurSheet.Activate
    Application.ScreenUpdating = True
End Sub
