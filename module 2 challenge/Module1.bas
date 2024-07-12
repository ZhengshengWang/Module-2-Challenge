Attribute VB_Name = "Module1"
Sub testing1()
    Dim i As Long
    Dim Ticker_Name As String
    Dim Ticker_Total As Double
    Dim Summary_Table_Row As Integer
    Dim WorksheetName As String
    Dim LastRow As Long
    Dim open_temp As Double
    Dim close_temp As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim maxChange As Double
    Dim maxChangeTick As String
    Dim minChange As Double
    Dim minChangeTick As String
    Dim maxVolume As Double
    Dim maxVolumeTick As String
    Ticker_Total = 0
    Dim ws As Worksheet
    For Each ws In Worksheets
    Summary_Table_Row = 2
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "max change"
    ws.Cells(3, 15).Value = "min change"
    ws.Cells(4, 15).Value = "max Volume"
    ws.Cells(1, 16).Value = "Ticker Name"
    ws.Cells(1, 17).Value = "Value"
    open_temp = ws.Cells(2, 3).Value
    maxChange = 0
    minChange = 0
    maxVolume = 0
    
    For i = 2 To LastRow
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        Ticker_Name = ws.Cells(i, 1).Value
        close_temp = ws.Cells(i, 6).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Or i = LastRow Then
            ws.Cells(Summary_Table_Row, 9).Value = ws.Cells(i, 1).Value
            quarterlyChange = close_temp - open_temp
            ws.Cells(Summary_Table_Row, 10).Value = quarterlyChange
            If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(Summary_Table_Row, 10).Value < 0 Then
                ws.Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            End If
            percentChange = quarterlyChange / open_temp
            ws.Cells(Summary_Table_Row, 11).Value = percentChange
            ws.Cells(Summary_Table_Row, 11).NumberFormat = "0.00%"
            ws.Cells(Summary_Table_Row, 12).Value = Ticker_Total
            Summary_Table_Row = Summary_Table_Row + 1
            If percentChange > maxChange Then
                maxChange = percentChange
                maxChangeTick = ws.Cells(i, 1).Value
            End If
            If percentChange < minChange Then
                minChange = percentChange
                minChangeTick = ws.Cells(i, 1).Value
            End If
            If Ticker_Total > maxVolume Then
                maxVolume = Ticker_Total
                maxVolumeTick = ws.Cells(i, 1).Value
            End If
            open_temp = ws.Cells(i + 1, 3).Value
            Ticker_Total = 0
            close_temp = 0
        End If
    Next i
    
    
    ws.Cells(2, 16).Value = maxChangeTick
    ws.Cells(3, 16).Value = minChangeTick
    ws.Cells(4, 16).Value = maxVolumeTick
    ws.Cells(2, 17).Value = maxChange
    ws.Cells(3, 17).Value = minChange
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).Value = maxVolume
    Next ws
    
End Sub
