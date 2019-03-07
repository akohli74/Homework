'Easy
Sub VolumeCombiner(ByVal ws_name As String)
    Dim current_ticker As String
    Dim current_volume As Variant
    Dim total_rows As Long
    Dim total_count As Long
    
    ActiveWorkbook.Worksheets(ws_name).Cells(1, 10).Value = "Ticker"
    ActiveWorkbook.Worksheets(ws_name).Cells(1, 11).Value = "Total Volume"
    
    total_rows = ActiveWorkbook.Worksheets(ws_name).Cells(ActiveWorkbook.Worksheets(ws_name).Rows.Count, 1).End(xlUp).Row
    current_ticker = ActiveWorkbook.Worksheets(ws_name).Cells(2, 1).Value
    current_volume = 0
    total_count = 2
    
    For i = 2 To total_rows
        If current_ticker <> ActiveWorkbook.Worksheets(ws_name).Cells(i, 1).Value Then
            ActiveWorkbook.Worksheets(ws_name).Cells(total_count, 10).Value = current_ticker
            ActiveWorkbook.Worksheets(ws_name).Cells(total_count, 11).Value = current_volume
            current_volume = 0
            total_count = total_count + 1
            current_ticker = ActiveWorkbook.Worksheets(ws_name).Cells(i, 1).Value
        End If
        current_volume = current_volume + ActiveWorkbook.Worksheets(ws_name).Cells(i, 7).Value
    Next i
End Sub

'Moderate
Sub StatsTaker(ByVal ws_name As String)
    If Not ActiveWorkbook.Worksheets(ws_name).Columns("K:K").Insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove) Then
            MsgBox ("Failed to insert column.")
    Else
            ActiveWorkbook.Worksheets(ws_name).Cells(1, 11).Value = "Yearly Change"
    End If
    If Not ActiveWorkbook.Worksheets(ws_name).Columns("K:K").Insert(Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove) Then
            MsgBox ("Failed to insert column.")
    Else
            ActiveWorkbook.Worksheets(ws_name).Cells(1, 12).Value = "Percent Change"
    End If

    Dim current_ticker As String
    Dim current_ticker_opening_price As Double
    Dim current_ticker_closing_price As Double
    Dim current_yearly_change As Double
    Dim current_percent_change As Double
    Dim current_ticker_date_raw As String
    Dim current_ticker_date As Variant
    Dim ticker_count As Long
    
    ticker_count = 1
    Dim total_rows As Long
    total_rows = ActiveWorkbook.Worksheets(ws_name).Cells(ActiveWorkbook.Worksheets(ws_name).Rows.Count, 1).End(xlUp).Row

    For i = 2 To total_rows
        If current_ticker <> ActiveWorkbook.Worksheets(ws_name).Cells(i, 1).Value Then
            current_ticker = ActiveWorkbook.Worksheets(ws_name).Cells(i, 1).Value
            current_ticker_opening_price = ActiveWorkbook.Worksheets(ws_name).Cells(i, 3).Value
            ticker_count = ticker_count + 1
        End If

        current_ticker_date_raw = ActiveWorkbook.Worksheets(ws_name).Cells(i, 2).Value
        current_ticker_date = CDate(Left(current_ticker_date_raw, 4) & "-" & Mid(current_ticker_date_raw, 5, 2) & "-" & Right(current_ticker_date_raw, 2))
        
        If current_ticker_date = CDate("30 Dec " + ws_name) Then
            current_ticker_closing_price = ActiveWorkbook.Worksheets(ws_name).Cells(i, 6).Value
            current_yearly_change = current_ticker_closing_price - current_ticker_opening_price
            ActiveWorkbook.Worksheets(ws_name).Cells(ticker_count, 11).Value = current_yearly_change
            If current_ticker_opening_price = 0 Then
                current_percent_change = current_yearly_change * 100
            Else
                current_percent_change = (current_yearly_change / current_ticker_opening_price) * 100
            End If
            ActiveWorkbook.Worksheets(ws_name).Cells(ticker_count, 12).Value = current_percent_change
        End If
        
    Next i
End Sub

'Challenge
Sub PopulateSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
       VolumeCombiner (ws.Name)
        StatsTaker (ws.Name)
    Next
    
    MsgBox ("All Done!")
End Sub

