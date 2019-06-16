Attribute VB_Name = "Module2"
Type InfoStocks2
   Ticker As String
   Value As Double
End Type

Sub StockSummary_Excel()
Dim SrcRange As Range, TrgRange As Range
Dim FirstCell As Range, LastCell As Range
Dim TRow As Variant
Dim i As Long
Dim OpenValue As Double, CloseValue As Double, Volume As Double
Dim ValChange As Double, PercChange As Double
Dim TickerSymb As String
Dim GrtStocks_exc(2) As InfoStocks2


'Output headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Yearly Change"
    Cells(1, 12).Value = "Percent Change%"
    Cells(1, 13).Value = "Total Stock Volume"
    Cells(1, 17).Value = "Ticker"
    Cells(1, 18).Value = "Value"
    Cells(2, 16).Value = "Greatest % Increase"
    Cells(3, 16).Value = "Greatest % Decrease"
    Cells(4, 16).Value = "Greatest Total Volume"

' Copy the ticker column and get the unique values in it
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("J2").Select
    ActiveSheet.Paste
    Range(Selection, Selection.End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo
   
' Define target range once the duplicates have been removed
    Range("J2").Select
    Set TrgRange = Range(Selection, Selection.End(xlDown))
    
' Define source range
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Set SrcRange = Range(Selection, Selection.End(xlToRight))
    
    i = 2
    For Each TRow In TrgRange.Rows
        Set FirstCell = SrcRange.Find(TRow, , xlValues, xlWhole, , xlNext, , , False)
        Set LastCell = SrcRange.Find(TRow, , xlValues, xlWhole, , xlPrevious, , , False)
        OpenValue = Range("C" & FirstCell.Row).Value
        CloseValue = Range("F" & LastCell.Row).Value
        
        TickerSymb = Range("A" & FirstCell.Row).Value
        Range("J" & i).Value = TickerSymb
        ValChange = CloseValue - OpenValue
        Range("K" & i).Value = ValChange
        If ValChange < 0 Then
            Range("K" & i).Interior.ColorIndex = 3 ' Red cell if change value is negative
        Else
            Range("K" & i).Interior.ColorIndex = 4 ' Green cell if change value is positive
        End If

    ' Special case: open and close Value were 0
        If OpenValue = 0 Then
            Range("L" & i).Value = ""
        Else
            PercChange = (CloseValue / OpenValue) - 1
            Range("L" & i).Value = PercChange
        End If
        Range("L" & i).NumberFormat = "0.00%"
        
        Volume = Application.Sum(Range(Cells(FirstCell.Row, 7), Cells(LastCell.Row, 7)))
        Range("M" & i).Value = Volume
        Range("M" & i).NumberFormat = "#,##0"

    ' Verify if values should go in greatStocks Array
        If (PercChange > GrtStocks_exc(0).Value) Then
            GrtStocks_exc(0).Ticker = TickerSymb
            GrtStocks_exc(0).Value = PercChange
        End If
        If (PercChange < GrtStocks_exc(1).Value) Then
            GrtStocks_exc(1).Ticker = TickerSymb
            GrtStocks_exc(1).Value = PercChange
        End If
        If (Volume > GrtStocks_exc(2).Value) Then
            GrtStocks_exc(2).Ticker = TickerSymb
            GrtStocks_exc(2).Value = Volume
        End If

    ' Get ready for next iteration
        i = i + 1
        Volume = 0
    Next TRow

' Report great Stocks
    ' Greatest % increase
    Cells(2, 17).Value = GrtStocks_exc(0).Ticker
    Cells(2, 18).Value = GrtStocks_exc(0).Value
    Cells(2, 18).NumberFormat = "0.00%"
    
    'Greatest % Decrease
    Cells(3, 17).Value = GrtStocks_exc(1).Ticker
    Cells(3, 18).Value = GrtStocks_exc(1).Value
    Cells(3, 18).NumberFormat = "0.00%"
    
    'Greatest total volume
    Cells(4, 17).Value = GrtStocks_exc(2).Ticker
    Cells(4, 18).Value = GrtStocks_exc(2).Value
    Cells(4, 18).NumberFormat = "#,##0"
    
' Format Results and set focus to A1 cell in that worksheet
    Columns("J:R").AutoFit
    ActiveSheet.Cells(1, 1).Select

End Sub
