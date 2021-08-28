Sub StockSummary():

    ' Declare worksheet object for iteration
    Dim ws As Worksheet
    ' Declare row counter for value pulling/ writing
    Dim i As Long
    ' Declare last row storage variable
    Dim RowNum As Long
    ' Declare variable to hold running volume total
    Dim RunTot As Double
    ' Declare variable to offset value writing
    Dim PasteOffset As Integer
    ' Set variable to row below headers
    PasteOffset = 2
    ' Declare variable to store first iteration open price
    Dim OpenPrice As Double
    ' Declare variable to store last iteration close price
    Dim ClosePrice As Double
    ' Declare variable to store percent change calc result
    Dim PctChg As Double
    ' Declare variable to hold first iteration state
    Dim FirstTime As Integer
    FirstTime = 0
    ' Declare variable to hold yearly change (calculated from openprice and closeprice)
    Dim YrChg As Double

    ' Begin looping over worksheets
    For Each ws In Worksheets
        ' Write headers in print area
        ws.Cells(1, 9).Value = "<ticker>"
        ws.Cells(1, 10).Value = "<Yearly Change>"
        ws.Cells(1, 11).Value = "<Percent Change>"
        ws.Cells(1, 12).Value = "<total volume>"
        ' Determine number of occupied rows for the sheet
        RowNum = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Begin iterating over rows in sheet
        For i = 2 To RowNum
            ' Test if ticker of current row is same as next row (all but last condition)
            If ws.Cells(i, 1) = ws.Cells(i, 1).Offset(1, 0) Then
                ' If ticker is same as the one below, increment firsttime
                ' The first time through the loop, firsttime will be 1 for the following test
                ' This is the only iteration where firsttime is 1, until the last row,
                ' where firsttime will be reset. This is a way to execute an if statement
                ' only if it is the first time going through the loop.
                FirstTime = FirstTime + 1
                ' Add volume cell for current row to running total
                RunTot = RunTot + ws.Cells(i, 7)
                ' Test if this is first time through loop for ticker in question
                If FirstTime = 1 Then
                    ' If it is, then grab open price for this first row of ticker,
                    ' and store for subsequent percent change calculation
                    OpenPrice = ws.Cells(i, 3)
                ' If it isn't the first time through the loop, do nothing
                Else
                End If
            Else
                ' This else block is only reached when the loop hits the very last row
                ' of the current ticker
                ' In this case, the volume is added to running total as normal
                RunTot = RunTot + ws.Cells(i, 7)
                ' Write the ticker and final volume total to the print area
                ws.Cells(PasteOffset, 9) = ws.Cells(i, 1)
                ws.Cells(PasteOffset, 12) = RunTot
                ' Grab the value of closing price in this last row and store for pct chg calc
                ClosePrice = ws.Cells(i, 6)
                ' Calculate percent change and yearly change with the stored variables
                ' If statement prevents dividing by zero if stock values are 0
                If OpenPrice <> 0 Then
                    PctChg = ((ClosePrice - OpenPrice) / OpenPrice)
                    YrChg = ClosePrice - OpenPrice
                Else
                    PctChg = 0
                    YrChg = 0
                End If
                ' Write percent change to print area using the paste offset to determine where
                ws.Cells(PasteOffset, 11) = PctChg
                ws.Cells(PasteOffset, 11).NumberFormat = "0.00%"
                ws.Cells(PasteOffset, 10) = YrChg
                   ' Evaluate whether the value is above or below 0, and color interior of cell
                    If ws.Cells(PasteOffset, 10).Value > 0 Then
                        ws.Cells(PasteOffset, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(PasteOffset, 10).Interior.ColorIndex = 3
                    End If
                ' Reset running total for next ticker group
                RunTot = 0
                ' Increment the paste offset for the printing so next group doesn't overwrite
                PasteOffset = PasteOffset + 1
                ' Reset first time through loop counter for next ticker group
                FirstTime = 0
            End If
        Next i
    ' Reset paste offset for next worksheet so printing starts again at row 2
    PasteOffset = 2
    ' Repeat proccess for next worksheet
    Next ws
End Sub

Sub ExtremeCheck():
    ' Declare for loop row counter
    Dim i As Long
    ' Declare for loop column counter
    Dim j As Integer
    ' Declare variable to store highest number
    Dim TopDog As Double
    ' Declare variable to store lowest number
    Dim UnderDog As Double
    ' Declare variable to store last row
    Dim RowNum As Long
    ' Declare worksheet object
    Dim ws As Worksheet
    ' Declare variables to hold ticker names
    Dim HiTicker As String
    Dim LwTicker As String
    ' Declare variable to hold paste offset increment
    Dim PstOS As Integer

    ' Loop over worksheets
    For Each ws In Worksheets
        ' Title headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(5, 14).Value = "Least Total Volume"

        ' Set paste offset, highest number, lowest number variables to zero
        PstOS = 0
        TopDog = 0
        UnderDog = 0
        ' Determine number of rows in sheet
        RowNum = ws.Cells(Rows.Count, 9).End(xlUp).Row

        ' Loop over columns
        For j = 11 To 12
            ' Loop over rows
            For i = 2 To RowNum
                'Test to see if current row value is higher than stored
                If ws.Cells(i, j).Value > TopDog Then
                    ' If row considered is higher, make it the new topdog, and grab its name
                    TopDog = ws.Cells(i, j).Value
                    HiTicker = ws.Cells(i, 9).Value
                Else
                End If
                ' Now check to see if the row considered is lower than lowest value
                ' If it is, make it the new underdog and grab its name
                If ws.Cells(i, j).Value < UnderDog Then
                    UnderDog = ws.Cells(i, j).Value
                    LwTicker = ws.Cells(i, 9).Value
                Else
                End If

            Next i
        ' After checking all rows, print the topdog number and its ticker name
        ws.Cells(2 + PstOS, 15).Value = HiTicker
        ws.Cells(2 + PstOS, 16).Value = TopDog
        ' Same with underdog
        ws.Cells(3 + PstOS, 15).Value = LwTicker
        ws.Cells(3 + PstOS, 16).Value = UnderDog
        ' Reset all variables
        HiTicker = "None"
        LwTicker = "None"
        TopDog = 0
        UnderDog = 0
        ' Increment paste offset so next paste happens below
        PstOS = PstOS + 2
        Next j
    ' Erase last row of readout, not relevant
    ws.Range("N5:P5").Value = ""
    ' Change percent change values to proper format
    ws.Range("P2:P3").NumberFormat = "0.00%"
    Next ws
End Sub
'Code credit: Max Parry on Github
