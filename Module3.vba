Sub ShowInstructions_Click()
   ActiveSheet.Rows("2:15").Select
   Selection.EntireRow.Hidden = False
End Sub
