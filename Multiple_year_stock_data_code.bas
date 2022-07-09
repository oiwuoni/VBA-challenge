Attribute VB_Name = "Module1"
Sub Multiple_Year_Stock_Data()

   
    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim stockVolume As Double
        Dim tickers As String
        Dim accIndex As Integer
        Dim YearlyChange As Double
        Dim start As Long
        Dim percentChange As Double
        start = 2
        accIndex = 1
    
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        RowCount = Cells(Rows.Count, 1).End(xlUp).Row
        For i = 2 To RowCount
        
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                ' Start the accumulation index
                accIndex = accIndex + 1
                ' Declare the ticker values
                tickers = Cells(i, 1).Value
                Cells(accIndex, 9).Value = tickers
                ' Calculate YearlyChange
                YearlyChange = Cells(i, 6).Value - Cells(start, 3).Value
                ' Declare the YearlyChange value
                Cells(accIndex, 10).Value = YearlyChange
                ' Calculate percentChange
                percentChange = YearlyChange / Cells(start, 3).Value
                ' start of the next stock ticker
                    start = i + 1
                ' print the results
                    Cells(accIndex, 10).NumberFormat = "0.00"
                    Cells(accIndex, 11).Value = percentChange
                    Cells(accIndex, 11).NumberFormat = "0.00%"
                stockVolume = stockVolume + Cells(i, 7).Value
                Cells(accIndex + 1, 12).Value = stockVolume
                stockVolume = 0
            End If
        
            If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
                ' Calculate the stockVolume for each ticker
                stockVolume = stockVolume + Cells(i, 7).Value
                Cells(accIndex + 1, 12).Value = stockVolume
            End If
        
        
        Next i
    
    Next ws
        





End Sub
