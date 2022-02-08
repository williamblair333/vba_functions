'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Delete_Row_If_Cell_Contains_String()

    Application.ScreenUpdating = False
    Call Text2Columns
    Call AutoFit
    Range("B1").Select

    While ActiveCell.Value <> " !end of configuration"
        Call FindMatch
    Wend
    Range("A1").Select
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Text2Columns()
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=True, OtherChar _
        :="]", FieldInfo:=Array(Array(1, 1), Array(2, 1)), TrailingMinusNumbers:=True
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FindMatch()
    If ActiveCell.Value Like "*]*" Or ActiveCell.Value Like "*!*" Or IsEmpty(ActiveCell.Value) Then
        Rows(ActiveCell.Row).Select
        Selection.Delete Shift:=xlUp
        ActiveCell.Offset(rowOffset:=0, columnOffset:=1).Activate
    
    Else
        ActiveCell.Offset(rowOffset:=1, columnOffset:=0).Activate
    
    End If
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AutoFit()
    Cells.Select
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ReportSetup()
    Sheets("Sheet1").Select
    Sheets("Sheet1").Copy After:=Sheets(1)
    Sheets("Sheet1 (2)").Select
    Sheets("Sheet1 (2)").Name = "report_copy"
    Sheets("Sheet1").Select
    Sheets("Sheet1").Name = "report"
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub Report_Restore()
    Application.DisplayAlerts = False
    Sheets("report").Select
    ActiveWindow.SelectedSheets.Delete
    
    Sheets("report_copy").Select
    Sheets("report_copy").Name = "report"
    
    Sheets("report").Select
    Sheets("report").Copy After:=Sheets(1)
    Sheets("report (2)").Select
    Sheets("report (2)").Name = "report_copy"
    Sheets("report").Select
    Range("A1").Select
    
End Sub
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''