Sub stockMarket()

    Dim sheet2014 as worksheet, sheet2015 as worksheet, sheet2016 as Worksheet
    Dim tempvol as Double
    Dim tickerPlace as Double
    Dim yearBeginOpen as Double
    Dim yearEndClose as Double
    tickerPlace = 2
    tempvol = 0

    Set sheet2014 = Worksheets("2014")
    Set sheet2015 = Worksheets("2015")
    Set sheet2016 = Worksheets("2016")

    yearBeginOpen = sheet2014.Cells(2,3).value
    for i = 2 to 705714
        ' If we are at the end of a dataset for the stock, then store values
        if sheet2014.Cells(i,1).value <> sheet2014.Cells(i+1,1).value Then
            tempvol = tempvol + sheet2014.Cells(i,7).value
            yearEndClose = sheet2014.Cells(i,6).value
            sheet2014.Range("I" & tickerPlace).value = Cells(i,1).value
            sheet2014.Range("L" & tickerPlace).value = tempvol
            sheet2014.Range("J" & tickerPlace).value = yearEndClose - yearBeginOpen
            if sheet2014.Range("J" & tickerPlace).value > 0 Then  
                sheet2014.Range("J" & tickerPlace).Interior.ColorIndex = 4 ' positive values filled in green
            Else 
                sheet2014.Range("J" & tickerPlace).Interior.ColorIndex = 3 ' negative values filled in red
            End if
            if yearBeginOpen <> 0 Then
                sheet2014.Range("K" & tickerPlace).value = sheet2014.Range("J" & tickerPlace).value / yearBeginOpen ' cannot divide a value of 0 by anything, account for this with an if statement later
            End if
            tempvol = 0
            tickerPlace = tickerPlace + 1
            yearBeginOpen = sheet2014.Cells(i+1,3).value
        Else
            tempvol = tempvol + sheet2014.Cells(i,7).value 
        End If
    
    Next i
    
    ' Yearly change in percent format
    sheet2014.Columns("K:K").Style = "Percent"
    

    tickerPlace = 2
    tempvol = 0
    yearBeginOpen = sheet2015.Cells(2,3).value
    for i = 2 to 760192
        if sheet2015.Cells(i,1).value <> sheet2015.Cells(i+1,1).value Then
            tempvol = tempvol + sheet2015.Cells(i,7).value
            yearEndClose = sheet2015.Cells(i,6).value
            sheet2015.Range("I" & tickerPlace).value = Cells(i,1).value
            sheet2015.Range("L" & tickerPlace).value = tempvol
            sheet2015.Range("J" & tickerPlace).value = yearEndClose - yearBeginOpen
            if sheet2015.Range("J" & tickerPlace).value > 0 Then  
                sheet2015.Range("J" & tickerPlace).Interior.ColorIndex = 4 ' positive values
            Else 
                sheet2015.Range("J" & tickerPlace).Interior.ColorIndex = 3 ' negative values
            End if
            if yearBeginOpen <> 0 Then
                sheet2015.Range("K" & tickerPlace).value = sheet2015.Range("J" & tickerPlace).value / yearBeginOpen ' cannot divide a value of 0 by anything, account for this with an if statement later
            End if
            tempvol = 0
            tickerPlace = tickerPlace + 1
            yearBeginOpen = sheet2015.Cells(i+1,3).value
        Else
            tempvol = tempvol + sheet2015.Cells(i,7).value 
        End If
    
    Next i
    
    ' Yearly change in percent format
    sheet2015.Columns("K:K").Style = "Percent"
    

    tickerPlace = 2
    tempvol = 0
    yearBeginOpen = sheet2016.Cells(2,3).value
    for i = 2 to 797711
        if sheet2016.Cells(i,1).value <> sheet2016.Cells(i+1,1).value Then
            tempvol = tempvol + sheet2016.Cells(i,7).value
            yearEndClose = sheet2016.Cells(i,6).value
            sheet2016.Range("I" & tickerPlace).value = Cells(i,1).value
            sheet2016.Range("L" & tickerPlace).value = tempvol
            sheet2016.Range("J" & tickerPlace).value = yearEndClose - yearBeginOpen
            if sheet2016.Range("J" & tickerPlace).value > 0 Then  
                sheet2016.Range("J" & tickerPlace).Interior.ColorIndex = 4 ' positive values
            Else 
                sheet2016.Range("J" & tickerPlace).Interior.ColorIndex = 3 ' negative values
            End if
            if yearBeginOpen <> 0 Then
                sheet2016.Range("K" & tickerPlace).value = sheet2016.Range("J" & tickerPlace).value / yearBeginOpen ' cannot divide a value of 0 by anything, account for this with an if statement later
            End if
            tempvol = 0
            tickerPlace = tickerPlace + 1
            yearBeginOpen = sheet2016.Cells(i+1,3).value
        Else
            tempvol = tempvol + sheet2016.Cells(i,7).value 
        End If
    
    Next i
   
    ' Yearly change in percent format
    sheet2016.Columns("K:K").Style = "Percent"
   


End Sub