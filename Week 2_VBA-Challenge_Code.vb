Sub AnnualSummary()

Dim k As Integer 'sheet counter
Dim i As Integer 'row counter
Dim j As Integer 'secondary row counter

For k = 1 To Sheets.Count 'for each sheet
    
    'make our column headers
    Sheets(k).Cells(1, 9) = "Ticker"
    Sheets(k).Cells(1, 10) = "Yearly Change"
    Sheets(k).Cells(1, 11) = "Percent Change"
    Sheets(k).Cells(1, 12) = "Total Stock Volume"
    
    'make our extrema table headers
    Sheets(k).Cells(2, 17) = "Greatest % Increase"
    Sheets(k).Cells(3, 17) = "Greatest % Decrease"
    Sheets(k).Cells(4, 17) = "Greatest Total Volume"
    Sheets(k).Cells(1, 18) = "Ticker"
    Sheets(k).Cells(1, 19) = "Value"
    
    'figure out how many rows we have on this sheet
    TotalRows = Sheets(k).Cells(Rows.Count, 1).End(xlUp).Row
    
    'we'll use this to keep track of how many unique tickers we have, telling us where to put then in the unique list
    Dim UniqueTicker As Integer
        UniqueTicker = 0 'initial unique ticker count. This is running through the whole sheet, which is why it's inside the k-loop, but outside the i-loop
    
    'we'll use this to have a running total of the total sales volume for that ticker
    Dim VolumeCounter As Double 'running total volume of sales for a ticker
            VolumeCounter = 0 'initial volume...not super necessary since the first situation gives it a value
    
    'declare variables to use indetermining yearly change and percent change
    Dim valueAtYearOpen As Double
    Dim valueAtYearClose As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    'this loop is going to run through each row of data, and get info
    For i = 1 To 20000 'TotalRows 'for every row
        
        'Make the unique ticker list
        If Sheets(k).Cells(i + 1, 1) <> Sheets(k).Cells(i, 1) Then 'if the ticker we're looking at isn't the same as the one ABOVE it, this means it's the first of it's kind
            UniqueTicker = UniqueTicker + 1 'add one to the number of unique tickers we've seen
            Sheets(k).Cells(UniqueTicker + 1, 9) = Sheets(k).Cells(i + 1, 1) 'put the ticker we're looking at in the new unique list
            VolumeCounter = Sheets(k).Cells(i + 1, 7) 'this is the first addition to our volume counter...the first time it shows up.
            valueAtYearOpen = Sheets(k).Cells(i + 1, 3) 'capture start value at open while we're here
        Else
            VolumeCounter = VolumeCounter + Sheets(k).Cells(i + 1, 7) 'if it is the same, we want to add its volume to it's running volume counter
        End If
        
        If Sheets(k).Cells(i + 1, 1) <> Sheets(k).Cells(i + 2, 1) Then 'if the ticker we're looking at isn't the same as the one BELOW it, this means it's the last of it's type
            valueAtYearClose = Sheets(k).Cells(i + 1, 6) 'capture close value while we're here
        End If
        
        'year change as difference of close and open
        yearlyChange = valueAtYearClose - valueAtYearOpen 'subtract to find the change
        Sheets(k).Cells(UniqueTicker + 1, 10) = yearlyChange 'put that change in unique list
        
        'conditional format yearly change
        If yearlyChange < 0 Then 'if yearly change is negative
            Sheets(k).Cells(UniqueTicker + 1, 10).Interior.ColorIndex = 3 'make the cell red
            Sheets(k).Cells(UniqueTicker + 1, 11).Interior.ColorIndex = 3 'same as above. NOTE: the instructions ask for this coloring, but the exemplar does not show it.
        ElseIf yearlyChange > 0 Then 'if it's positive
            Sheets(k).Cells(UniqueTicker + 1, 10).Interior.ColorIndex = 4 'make it green
            Sheets(k).Cells(UniqueTicker + 1, 11).Interior.ColorIndex = 4 'same as above. NOTE: the instructions ask for this coloring, but the exemplar does not show it.
        End If 'if it's =0 then leave it blank, not that I super expect that to happen
        
        'percent change is (open/close)-1
        If valueAtYearOpen = 0 Then 'this is my way to avoid a div/0 error that I'm not sure I understand why it even came up...because none of these had 0 value at open? But maybe this is it accidentally reading empy cells, and it doesn't matter if that happens anyways, so I'll make this silly little conditional to handle if we accidentally divide by zero
        Else
            Sheets(k).Cells(UniqueTicker + 1, 11) = FormatPercent(0) 'if we're dividing by 0....just put 0...again, not gonna happen
            Sheets(k).Cells(UniqueTicker + 1, 11) = FormatPercent((valueAtYearClose / valueAtYearOpen) - 1)
        End If
        
        'put in total volume for that ticker in the unique list
        Sheets(k).Cells(UniqueTicker + 1, 12) = VolumeCounter
        
    Next i 'do this allllllll the way down the sheet
    
    'now we want to make the extrema table
        'identify holding variable for extreme values. Starting at whatever the first candidate value is, these will hold the extreme value until a more extreme value comes along and replaces it. If that doesn't happen...congrats, I guess the first ticker had the extreme value!
        Dim greatestPIncrease As Double
            greatestPIncrease = Sheets(k).Cells(2, 11)
        Dim greatestPDecrease As Double
            greatestPDecrease = Sheets(k).Cells(2, 11)
        Dim greatestVolume As Double
            greatestVolume = Sheets(k).Cells(2, 13)
        
        'identify variables fot the tickers that go with those values
        Dim greatestPIncreaseTick As String
            greatestPIncreaseTick = Sheets(k).Cells(2, 9)
        Dim greatestPDecreaseTick As String
            greatestPDecreaseTick = Sheets(k).Cells(2, 9)
        Dim greatestVolumeTick As String
            greatestVolumeTick = Sheets(k).Cells(2, 9)
    
    For j = 1 To UniqueTicker 'for every unique ticker in the ticker list
        
        'if you find a new greater increase, that's my value, if not, move on
        If Sheets(k).Cells(j + 1, 11) > greatestPIncrease Then
            greatestPIncrease = Sheets(k).Cells(j + 1, 11) 'set it to this new greater value
            greatestPIncreaseTick = Sheets(k).Cells(j + 1, 9) 'capture the ticker that goes with it
        End If
        
        'if you find a new greater decrease, that's my value, if not, move on
        If Sheets(k).Cells(j + 1, 11) < greatestPDecrease Then
            greatestPDecrease = Sheets(k).Cells(j + 1, 11) 'set it to this new lesser value
            greatestPDecreaseTick = Sheets(k).Cells(j + 1, 9) ' capture the ticker that goes with it
        End If
        
        'if you find a new greater volume, that's my value, if not, move on
        If Sheets(k).Cells(j + 1, 12) > greatestVolume Then
            greatestVolume = Sheets(k).Cells(j + 1, 12) 'set it to this new greater value
            greatestVolumeTick = Sheets(k).Cells(j + 1, 9) 'capture the ticker that goes with it
        End If
        
        'put these values all in a table
        Sheets(k).Cells(2, 19) = FormatPercent(greatestPIncrease)
        Sheets(k).Cells(3, 19) = FormatPercent(greatestPDecrease)
        Sheets(k).Cells(4, 19) = greatestVolume
        
        'put the tickers with them
        Sheets(k).Cells(2, 18) = greatestPIncreaseTick
        Sheets(k).Cells(3, 18) = greatestPDecreaseTick
        Sheets(k).Cells(4, 18) = greatestVolumeTick

    Next j
                
Sheets(k).Columns("A:S").AutoFit 'give columns nice widths
Next k

End Sub

Sub reset()

'to make it easy to reset and recheck

For k = 1 To Sheets.Count
    Sheets(k).Range("I:T") = ""
    Sheets(k).Range("I:T").Interior.ColorIndex = 2
    Sheets(k).Range("I:T").Borders.ColorIndex = 15
Next k

End Sub
