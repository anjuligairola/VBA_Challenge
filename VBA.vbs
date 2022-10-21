Sub stockdata()
    Const CLOSE_COLUMN As Integer = 6
    Const OPEN_COLUMN As Integer = 3
    Const TICKER_COLUMN As Integer = 1
    Const VOLUME_COLUMN As Integer = 7
    
    Dim ticker As String
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalstockvolume As Double
    Dim summarytablerow As Integer
    Dim Row As Long
    summarytablerow = 2

    For Row = 2 To 753001
        If Cells(Row - 1, 1).Value <> Cells(Row, 1).Value Then 'first row of the ticker'
            Opening_price = Cells(Row, OPEN_COLUMN).Value
            totalstockvolume = 0
        End If
        
        totalstockvolume = totalstockvolume + Cells(Row, VOLUME_COLUMN).Value
        
        If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then 'last row of the ticker'
            Closing_price = Cells(Row, CLOSE_COLUMN).Value
            yearlychange = Closing_price - Opening_price
            percentchange = (yearlychange / Opening_price) * 100
            
            ticker = Cells(Row, TICKER_COLUMN).Value
            totalstockvolume = totalstockvolume + Cells(Row, VOLUME_COLUMN).Value
            
            Cells(summarytablerow, 10).Value = ticker
            Cells(summarytablerow, 11).Value = yearlychange
            If yearlychange < 0 Then
                Cells(summarytablerow, 11).Interior.ColorIndex = 3
            Else
                Cells(summarytablerow, 11).Interior.ColorIndex = 4
            End If
            Cells(summarytablerow, 12).Value = percentchange
            Cells(summarytablerow, 13).Value = totalstockvolume
            
           
            summarytablerow = summarytablerow + 1
         End If
    Next Row
    
    'Bonus
    

Dim maxpercentincrease As Double
Dim maxpercentdecrease As Double


Set range1 = Sheet1.Range("L2:L3001")
maxpercentincrease = Application.WorksheetFunction.Max(range1)
Cells(2, 20).Value = maxpercentincrease
maxpercentdecrease = Application.WorksheetFunction.Min(range1)
Cells(3, 20).Value = maxpercentdecrease

Dim maxtotalvolume As Double
Set range2 = Sheet1.Range("M2:M3001")
maxtotalvolume = Application.WorksheetFunction.Max(range2)
Cells(4, 20).Value = maxtotalvolume

End Sub


