Sub HideInstructions_Click()
    ActiveSheet.Rows("2:15").Select
    Selection.EntireRow.Hidden = True
End Sub
