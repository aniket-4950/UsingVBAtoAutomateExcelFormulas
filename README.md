# UsingVBAtoAutomateExcelFormulas

Code with comments:
Public Sub FunWithFormula()
    Dim ws As Worksheet
    Dim lastCell As String
    For Each ws In Worksheets 'looping through the sheets
        Worksheets(ws.Name).Select 'selecting the current worksheet
        Range("F2").Select 'selecting the first value in the total column
        Selection.End(xlDown).Select 'moving all thw way down to last cell and making it dynamic for each and every worksheet
        lastCell = ActiveCell.Address(False, False) '(false,false) for a relative reference else it would have taken the absolute reference
        ActiveCell.Offset(1, 0).Select 'moving down the last active cell to store the sum
        ActiveCell.Value = "=sum(F2:" & lastCell & ")" 'creating the formula for summing up the total
    Next ws 'iterating to the next worksheet
End Sub
