Sub WellsFargo_1(ByVal ws_name As String)

    If (ActiveWorkbook.Worksheets(ws_name).Columns("A:A").Insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove)) Then
        
    End If
      
    ActiveWorkbook.Worksheets(ws_name).Cells(1, 1).Value = "State"
    
    FillState (ws_name)
    
    FixHeaders (ws_name)
    
    ChangeToCurrency (ws_name)
    
End Sub

Sub ChangeToCurrency(ByVal ws_name As String)
    Dim totalRows As Integer
    totalRows = ActiveWorkbook.Worksheets(ws_name).Cells(ActiveWorkbook.Worksheets(ws_name).Rows.Count, 2).End(xlUp).Row
    Dim totalColumns As Integer
    totalColumns = ActiveWorkbook.Worksheets(ws_name).UsedRange.Columns.Count
    
    For i = 2 To totalRows
        For j = 3 To totalColumns
            ActiveWorkbook.Worksheets(ws_name).Cells(i, j).Style = "Currency"
        Next j
    Next i
End Sub

Sub FillState(ByVal ws_name As String)
    Dim stateName As String
    
    stateName = Split(ws_name, "_")(0)
    Dim totalRows As Integer
    totalRows = ActiveWorkbook.Worksheets(ws_name).Cells(ActiveWorkbook.Worksheets(ws_name).Rows.Count, 2).End(xlUp).Row
    For i = 2 To totalRows
        ActiveWorkbook.Worksheets(ws_name).Cells(i, 1).Value = stateName
    Next i
End Sub

Sub FixHeaders(ByVal ws_name As String)
    Dim totalColumns As Integer
    Dim year As String
    totalColumns = ActiveWorkbook.Worksheets(ws_name).Cells(1, ActiveWorkbook.Worksheets(ws_name).Columns.Count).End(xlToLeft).Column
    
    For i = 3 To 7
        year = Mid(ActiveWorkbook.Worksheets(ws_name).Cells(1, i), 22, 4)
        ActiveWorkbook.Worksheets(ws_name).Cells(1, i).Value = year
    Next i
End Sub

Sub PopulateSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        WellsFargo_1 (ws.Name)
    Next
    
    MsgBox ("All Done!")
End Sub
