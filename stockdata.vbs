Sub wall_street():

    ' Set variable for holding ticker symbol
    Dim ticker As String
    
    ' Set variable for holding opening price
    Dim openPrice As Double
    openPrice = 0
    
    ' Set variable for closing price
    Dim closingPrice As Double
    closingPrice = 0
    
    ' Set variable for holding yearly change price
    Dim priceChange As Double
    priceChange = 0
    
    ' Set variable for holding percent change price
    Dim percentChange As Double
    percentChange = 0
    
    ' Set variable for holding total volume of stock
    Dim volume As Double
    volume = 0
    
    'Set variable for last row
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Keep track of location for each stock ticker in summary table
    Dim summaryTableRow As Integer
    summaryTableRow = 2
    
    'Keep track of location for opening price
    Dim openPriceRow As Long
    openPriceRow = 2
    
    'Label Header Colums
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Volume"
    
    'Loop through all financial data
    For i = 2 To lastRow
    
        'Check if same stock ticker
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Capture stock ticker name
            ticker = Cells(i, 1).Value
            
            'Add to volume
            volume = volume + (Cells(i, 7).Value)
            
            'Set stock ticker name
            Range("I" & summaryTableRow).Value = ticker
            
            'Set stock price change
            'Range("J" & summaryTableRow).Value = priceChange
            
            'Set stock price percent change
            'Range("K" & summaryTableRow).Value = percentChange
            
            'Set stock volume
            Range("L" & summaryTableRow).Value = volume
            
            'Reset volume
            volume = 0
            
            'Set openning price
            openPrice = Range("C" & openPriceRow).Value
            
            'Get closing price
            closingPrice = Cells(i, 6).Value
            
            'Calculate priceChange
            priceChange = closingPrice - openPrice
            
            'Calculate percentChange
            percentChange = priceChange / closingPrice
            
            'Set Yearly change
            Range("J" & summaryTableRow).Value = priceChange
            
            'Set Percent change
            Range("K" & summaryTableRow).Value = percentChange
            
            'Add one to the summary table row
            summaryTableRow = summaryTableRow + 1
            
            'Move down openPriceRow
            openPriceRow = i + 1
        
        'If the cell immediately following a row is the same ticker
        Else
            
            'Add to volume
            volume = volume + Cells(i, 7).Value
            
        End If
    
    Next i
    
    
    'Conditional Formatting
    Dim rng As Range
    Dim condition1 As FormatCondition, condition2 As FormatCondition
    
    Set rng = Range("J2:J" & lastRow)
    
        rng.FormatConditions.Delete
    
        Set condition1 = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        Set condition2 = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
    
        With condition1
           .Interior.ColorIndex = 4
        End With
    
        With condition2
            .Interior.ColorIndex = 3
        End With
        
    'Row K to % format
    Range("K2:K" & lastRow).NumberFormat = "0.00%"
    Range("Q2:Q3").NumberFormat = "0.00%"
    
    'CHALLENGE #1
    
    'Greatest % Increase
    Dim greatestIncrease As Double
    greatestIncrease = 0

    'Greatest % Decrease
    Dim greatestDecrease As Double
    greatestDecrease = 0

    'Greatest Total Volume
    Dim greatVolume As Double
    greatVolume = 0
    
    For i = 2 To lastRow
        If Cells(i, 11).Value > greatestIncrease Then
            greatestIncrease = Cells(i, 11).Value
            Cells(2, 16) = Cells(i, 9).Value
            Cells(2, 17) = greatestIncrease
        End If
        If Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = Cells(i, 11).Value
            Cells(3, 16) = Cells(i, 9).Value
            Cells(3, 17) = greatestDecrease
        End If
        If Cells(i, 12).Value > greatVolume Then
            greatVolume = Cells(i, 12).Value
            Cells(4, 16) = Cells(i, 9).Value
            Cells(4, 17) = greatVolume
        End If
    Next i
    
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"

End Sub