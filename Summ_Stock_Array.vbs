Attribute VB_Name = "Module3"
Type InfoStocks
   Ticker As String
   Value As Double
End Type
Sub StockSummary_Array()

Dim Current_WS As Worksheet
Dim InMemStock As Variant
Dim ORow As Long
Dim OpenValue As Double, ValChange As Double, PercChange As Double
Dim Volume As Single
Dim GrtStocks(2) As InfoStocks

Dim i As Long

' Loop through all of the worksheets in the active workbook.
    For Each Current_WS In Worksheets
        Current_WS.Activate

    'Output headers
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly Change"
        Range("L1").Value = "Percent Change%"
        Range("M1").Value = "Total Stock Volume"
        Range("Q1").Value = "Ticker"
        Range("R1").Value = "Value"
        Range("P2").Value = "Greatest % Increase"
        Range("P3").Value = "Greatest % Decrease"
        Range("P4").Value = "Greatest Total Volume"
    
    'Insert a final value to avoid going out of range in InMemStock(i + 1, 1) <> InMemStock(i, 1) check
        Range("A" & Cells(Rows.Count, 1).End(xlUp).Row + 1).Value = "END"
        
    'Read Stock info into memory array for processing
        Range("A1").Select
        Range(Selection, Selection.End(xlDown)).Select
        InMemStock = Range(Selection, Selection.End(xlToRight))
        
        ORow = 2
        OpenValue = InMemStock(2, 3)  ' Open value of first stock
        Volume = 0
        
        For i = 2 To UBound(InMemStock, 1) - 1
        ' If the next ticker is different from the current ticker, summarize it
            If (InMemStock(i + 1, 1) <> InMemStock(i, 1)) Then
            
            ' Ticker name
                Range("J" & ORow).Value = InMemStock(i, 1)
                
            ' Value Change
                ValChange = InMemStock(i, 6) - OpenValue
                Range("K" & ORow).Value = ValChange
            'Format depending if value change is positive or negative
                If ValChange < 0 Then
                    Range("K" & ORow).Interior.ColorIndex = 3     ' Red
                Else
                    Range("K" & ORow).Interior.ColorIndex = 4     ' Green
                End If
                
            ' Percentage of Change
                If OpenValue = 0 Then ' Special case: open Value was 0
                    Range("L" & ORow).Value = ""
                    PercChange = 0
                Else
                    PercChange = (InMemStock(i, 6) / OpenValue) - 1
                    Range("L" & ORow).Value = PercChange
                End If
            ' Format as percentage
                Range("L" & ORow).NumberFormat = "0.00%"
                
            ' Volume
                Volume = Volume + InMemStock(i, 7)
                Range("M" & ORow).Value = Volume
            ' Format Volume
                Range("M" & ORow).NumberFormat = "#,##0"
    
            ' Verify if values should go in greatStocks Array
                If (PercChange > GrtStocks(0).Value) Then
                    GrtStocks(0).Ticker = InMemStock(i, 1)
                    GrtStocks(0).Value = PercChange
                End If
                If (PercChange < GrtStocks(1).Value) Then
                    GrtStocks(1).Ticker = InMemStock(i, 1)
                    GrtStocks(1).Value = PercChange
                End If
                If (Volume > GrtStocks(2).Value) Then
                    GrtStocks(2).Ticker = InMemStock(i, 1)
                    GrtStocks(2).Value = Volume
                End If
    
            'Prepare for next stock group
                OpenValue = InMemStock(i + 1, 3)
                Volume = 0
                ORow = ORow + 1
                
            Else
                Volume = Volume + InMemStock(i, 7)
            End If
        Next i
    
    ' Report great Stocks
        ' Greatest % increase
        Range("Q2").Value = GrtStocks(0).Ticker
        Range("R2").Value = GrtStocks(0).Value
        Range("R2").NumberFormat = "0.00%"
        
        'Greatest % Decrease
        Range("Q3").Value = GrtStocks(1).Ticker
        Range("R3").Value = GrtStocks(1).Value
        Range("R3").NumberFormat = "0.00%"
        
        'Greatest total volume
        Range("Q4").Value = GrtStocks(2).Ticker
        Range("R4").Value = GrtStocks(2).Value
        Range("R4").NumberFormat = "#,##0"
        
    ' Format Results and set focus to A1 cell in that worksheet
        Columns("J:R").AutoFit
        ActiveSheet.Cells(1, 1).Select
        
    ' Clean Great Stock Array for next worksheet
        GrtStocks(0).Ticker = ""
        GrtStocks(0).Value = 0
        GrtStocks(1).Ticker = ""
        GrtStocks(1).Value = 0
        GrtStocks(2).Ticker = ""
        GrtStocks(2).Value = 0
        
    ' Delete inserted value
        Range("A" & Cells(Rows.Count, 1).End(xlUp).Row).Value = ""
    Next
End Sub
