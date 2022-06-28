Attribute VB_Name = "Module1"
'This code will output the following information:
    'Yearly change for specific stock by volume
    'Yearly change for specific stock by percentage
    'total stock volume
'This code will also:
    'Format perecentage change for positive and negative values
    'Create an array to display said data
    'Apply data amongst all worksheets

Sub Stocktracker()

'Set variables for Worksheets and cells to display data

For Each WS In Worksheets
    
'Variables
Dim StockTicker As String
Dim YearlyChange As Double
Dim PercentChange As Double
Dim Volume As Double
    
Dim StockOpen As Double
Dim StockClose As Double
    
Dim lastrow As Double
lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'Table Headers
WS.Range("I1") = "Stock Ticker"
WS.Range("J1") = "Yearly Change"
WS.Range("K1") = "Percent Change"
WS.Range("L1") = "Total Stock Volume"

'Set to Zero
Volume = 0
Dim Summary_Table_Row As Double
Summary_Table_Row = 2

For I = 2 To lastrow

'Set Total Ticker and Total Volume data


If WS.Cells(I + 1, 1).Value <> WS.Cells(I, 1) Then
    StockTicker = WS.Cells(I, 1).Value
    Volume = Volume + WS.Cells(I, 7).Value

WS.Range("I" & Summary_Table_Row).Value = StockTicker
WS.Range("L" & Summary_Table_Row).Value = Volume

Volume = 0

StockClose = WS.Cells(I, 6)

'Calculate Yearly and Percent Change

    If StockOpen = 0 Then
    YearlyChange = 0
    PercentChange = 0
    
    Else:
    YearlyChange = StockClose - StockOpen
    PercentChange = (StockClose - StockOpen) / StockOpen
    
    End If
    
'Set Yearly change, Percent Change, and Percent on the table
    
WS.Range("J" & Summary_Table_Row).Value = YearlyChange
WS.Range("K" & Summary_Table_Row).Value = PercentChange
WS.Range("K" & Summary_Table_Row).Style = "Percent"
WS.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

Summary_Table_Row = Summary_Table_Row + 1

ElseIf WS.Cells(I - 1, 1).Value <> WS.Cells(I, 1) Then
    StockOpen = WS.Cells(I, 3)
    
Else: Volume = Volume + WS.Cells(I, 7).Value

End If

Next I

'Format Percentage change

For I = 2 To lastrow

'Green for Positive Value
If WS.Range("J" & I).Value > 0 Then
        WS.Range("J" & I).Interior.ColorIndex = 4
'Red for Negative Value
ElseIf WS.Range("J" & I).Value < 0 Then
        WS.Range("J" & I).Interior.ColorIndex = 3
        
End If

Next I
    
Next WS


End Sub


