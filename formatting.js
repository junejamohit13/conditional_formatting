Sub ApplyConditionalFormattingToAllCells()
    Dim ws As Worksheet
    Dim rng As Range
    
    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Set the range to the used range of each sheet
        Set rng = ws.UsedRange
        
        ' Clear any existing conditional formatting
        rng.FormatConditions.Delete
        
        ' Add new conditional formatting to highlight cells with formulas
        With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=ISFORMULA(INDIRECT(""rc"",FALSE))")
            .Font.Color = RGB(255, 0, 0) ' Red font color
        End With
    Next ws
End Sub
