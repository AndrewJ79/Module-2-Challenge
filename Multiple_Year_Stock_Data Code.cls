VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim row As Long
    Dim LastRow As Long
    Dim open_price As Double
    Dim volume As Double
    Dim ticker As String
    Dim i As Long
    Dim ticker_column As Long
    Dim tinker_column As Long
    
    ' Define the ticker_column and tinker_column with appropriate indices
    ticker_column = 1
    tinker_column = 1

    For Each ws In ActiveWorkbook.Worksheets
        row = 2
        LastRow = ws.Cells(Rows.Count, "A").End(xlUp).row
        
        ' Set headers
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        open_price = ws.Cells(2, tinker_column + 2).Value
        volume = 0

        For i = 2 To LastRow
            If ws.Cells(i + 1, ticker_column).Value <> ws.Cells(i, ticker_column).Value Then
                ' Same ticker
                ticker = ws.Cells(i, ticker_column).Value
                ws.Cells(row, ticker_column + 8).Value = ticker
                
                volume = volume + ws.Cells(i, ticker_column + 6).Value
                ws.Cells(row, ticker_column + 11).Value = volume
                
                close_price = ws.Cells(i, ticker_column + 5).Value
                qtly_change = close_price - open_price
                percent_change = qtly_change / open_price
                
                ws.Cells(row, ticker_column + 9).Value = qtly_change
                ws.Cells(row, ticker_column + 10).Value = percent_change
                ws.Cells(row, ticker_column + 10).NumberFormat = "0.00%"
                
                row = row + 1
                open_price = ws.Cells(i + 1, tinker_column + 2).Value ' Adjust as needed
            Else
                ' Different ticker; accumulate volume and write if applicable
                volume = volume + ws.Cells(i, ticker_column + 6).Value
                
                If i = LastRow Or ws.Cells(i + 1, ticker_column).Value <> ws.Cells(i, ticker_column).Value Then
                    ws.Cells(row, ticker_column + 11).Value = volume
                    volume = 0 ' Reset for next ticker
                End If
            End If
        Next i
                qtly_change_lastrow = ws.Cells(Rows.Count, "J").End(xlUp).row
        
        For j = 2 To qtly_change_lastrow
            If ws.Cells(j, 10).Value > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
        Next j
        
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        For k = 2 To qtly_change_lastrow
            If ws.Cells(k, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & qtly_change_lastrow)) Then
                ws.Cells(2, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(2, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(2, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & qtly_change_lastrow)) Then
                ws.Cells(3, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(3, 17).Value = ws.Cells(k, 11).Value
                ws.Cells(3, 17).NumberFormat = "0.00%"
            ElseIf ws.Cells(k, ticker_column + 11).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & qtly_change_lastrow)) Then
                ws.Cells(4, 16).Value = ws.Cells(k, 9).Value
                ws.Cells(4, 17).Value = ws.Cells(k, 12).Value
            End If
        Next k
        
       
        
    
    
   
    Next ws
        MsgBox ("Test Complete")
End Sub

        
