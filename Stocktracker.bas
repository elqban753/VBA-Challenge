Attribute VB_Name = "Module1"
'This code will output the following information:
    'Yearly change for specific stock by volume
    'Yearly change for specific stock by percentage
    'total stock volume
    
    
Sub Stocktracker():

    Stock = ""
    
    totalvolume = 0
    
    summarytablerow = 2
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    For Row = 2 To lastRow
    
    If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
    
        Stock = Cells(Row, 1).Value
        
        totalvolume = totalvolume + Cells(Row, 7).Value
        
        Cells(summarytablerow, 9).Value = Stock
        
        Cells(summarytablerow, 12).Value = totalvolume
        
        summarytablerow = summarytablerow + 1
        
        totalvolume = 0
        
    Else
        
        totalvolume = totalvolume + Cells(Row, 7).Value
        
    End If
    
    Next Row

        
End Sub
