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
    
    rowStart = 2
    'Delete: rowEnd = 3013
    rowEnd = Cells(Row.Count, "A").End(xlUp).Row
    totalVolume = 0
    
    Worksheets("2018").Activate
    
        For i = rowStart To rowEnd
            'increase totalVolume if Ticker is "DQ"
            If Cells(i, 1).Value = "DQ" Then
                totalVolume = totalVolume + Cells(i, 8).Value
            
            End If
        
        Next i
    
    'MsgBox (totalVolume)
    
    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    
    
End Sub
