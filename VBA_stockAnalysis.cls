VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub AllStocksAnalysisrefactored()

Dim startTime As Single
Dim endTime As Single


yearValue = InputBox("What year would you like to run analysis on?")

    startTime = Timer
    
    Sheets("All Stock Analysis").Activate

    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Initialize an array of all tickers.

    Dim tickers(12) As String
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"

    
    'Activate the data worksheet.
    Sheets(yearValue).Activate
    
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    

    '1a) Create a ticker Index
    Dim tickerIndex As Single
    tickerIndex = 0
    
    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
        tickerVolumes(i) = 0
        
    Next i
    
    Sheets(yearValue).Activate
    
    '2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
      tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
                
        '3b) Check current row that it is the first row with the selected tickerIndex.
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        '3c) Check current row that it is the last row with the selected ticker
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        '3d) If the next row’s ticker doesn’t match, increase the tickerIndex.
                tickerIndex = tickerIndex + 1
            End If

    Next i
     
    Worksheets("All Stock Analysis").Activate
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
                   
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
        
    
    Worksheets("All Stock Analysis").Activate
    
    'formatting the table
    Range("A1").Font.Bold = True
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").AutoFit
    
    dataRowStart = 4
    dataRowEnd = 15
    
    For i = dataRowStart To dataRowEnd
        'Color cell green for Positive Return
        If Cells(i, 3) > 0 Then
            Cells(i, 3).Interior.Color = vbGreen
        'Color cell red for Negative Return
        ElseIf Cells(i, 3) < 0 Then
            Cells(i, 3).Interior.Color = vbRed
        Else
            Cells(i, 3).Interior.Color = xlNone
        End If
    
    Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
    
        
End Sub