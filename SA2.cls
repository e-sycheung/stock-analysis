VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    
    Worksheets("2018").Activate
    
    rowStart = 2
    
    'Delete: rowEnd = 3013 & used formula from stack overflow
    'https://stackoverflow.com/questions/18088729/row-count
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
   
    
        For i = rowStart To rowEnd
            'increase totalVolume if Ticker is "DQ"
            If Cells(i, 1).Value = "DQ" Then
                totalVolume = totalVolume + Cells(i, 8).Value
            
            End If
            
            If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
            
                startingPrice = Cells(i, 6).Value
        
            End If
            
            If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
            
                endingPrice = Cells(i, 6).Value
                
            End If
        
        Next i
        
    
    
    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
    
End Sub

Sub AllStocksAnalysis()

'Format the output sheet on the "All Stocks Analysis" worksheet.
    
    Worksheets("All Stock Analysis").Activate

    Range("A1").Value = "All Stocks (2018)"
    
    'Create header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
'Initialize an array of all tickers.

    Dim tickers(11) As String
    
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

    'Initialize variables for the starting price and ending price.
    Dim startingPrice As Single
    Dim endingPrice As Single

    'Activate the data worksheet.
    Worksheets("2018").Activate
    
    'Find the number of rows to loop over.
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'MsgBox (RowCount)
    
    For i = 0 To 11
    
        ticker = tickers(i)
        totalVolume = 0
        
        Worksheets("2018").Activate
        
        For j = 2 To RowCount
        
            If Cells(j, 1).Value = ticker Then
                totalVolume = totalVolume + Cells(j, 8).Value
            End If
            
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
                startingPrice = Cells(j, 6).Value
            End If
            
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
            End If
                
        Next j
        
        Worksheets("All Stock Analysis").Activate
        
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
        
    Next i
    
End Sub
    
    