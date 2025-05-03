Private Sub Workbook_Open()
Call CopyDataByDateRange(False)
Call CalculateDailyAverages
End Sub

Public Sub doit()
Call CopyDataByDateRange(True)
Call CalculateDailyAverages
End Sub