VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tickerLoop():
    'storing ticker letter to iterate through each
    Dim currTicker As String
    currTicker = ""
    'storing year start value
    Dim yStart As Double
    yStart = 0
    'storing year end value
    Dim yEnd As Double
    yEnd = 0
    'storing total volume
    Dim volume As Double
    volume = 0
    'store current row number for analysis
    Dim aRow As Integer
    aRow = 2

    
    'loop through each sheet
    For Each ws In Worksheets
        
        'change to next sheet
        Sheets(ws.Name).Select
        
        'labels
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"
        
        Range("N2").Value = "Greatest % Increase"
        Range("N3").Value = "Greatest % Decrease"
        Range("N4").Value = "Greatest Total Volume"
        Range("O1").Value = "Ticker"
        Range("P1").Value = "Value"
        
        'reset aRow value
        aRow = 2
        'store current row number for data
        Dim row As Long
        row = 2
        
        'store number of rows
        Dim rowMax As Long
        rowMax = Cells(Rows.Count, 2).End(xlUp).row
        
        'store greatest values
        Dim iTicker As String
        iTicker = ""
        Dim dTicker As String
        dTicker = ""
        Dim vTicker As String
        vTicker = ""
        Dim gIncrease As Double
        gIncrease = 0
        Dim gDecrease As Double
        gDecrease = 0
        Dim gVolume As Double
        gVolume = 0
        
        'go through all of the rows
        While row <= rowMax
        'update values
        currTicker = Cells(row, 1).Value
        yStart = Cells(row, 3).Value
        volume = Cells(row, 7).Value
        'go through each row with same ticker
            While Cells(row, 1).Value = currTicker
                'increase volume
                volume = volume + Cells(row, 7).Value
                row = row + 1
            Wend
                row = row - 1
                'get yEnd value
                yEnd = Cells(row, 6)
                
                'store the values in row
                Cells(aRow, 9).Value = currTicker
                Cells(aRow, 10).Value = yEnd - yStart
                
                'check to make sure doesn't divide by 0
                If yStart <> 0 Then
                    Cells(aRow, 11).Value = Format((yEnd - yStart) / yStart, "Percent")
                Else: Cells(aRow, 11).Value = 0
                End If
                'conditional formatting
                If Cells(aRow, 11).Value < 0 Then
                    Cells(aRow, 11).Interior.ColorIndex = 3
                    'update decrease
                    If Cells(aRow, 11).Value < gDecrease Then
                        gDecrease = Cells(aRow, 11).Value
                        dTicker = currTicker
                    End If
                Else: Cells(aRow, 11).Interior.ColorIndex = 4
                    'update increase
                    If Cells(aRow, 11).Value > gIncrease Then
                        gIncrease = Cells(aRow, 11).Value
                        iTicker = currTicker
                    End If
                End If
                Cells(aRow, 12).Value = volume
                'update volume
                If volume > gVolume Then
                    gVolume = volume
                    vTicker = currTicker
                End If
    
                row = row + 1
                aRow = aRow + 1
        Wend

    'greatest tickers
    Range("O2").Value = iTicker
    Range("O3").Value = dTicker
    Range("O4").Value = vTicker
    Range("P2").Value = Format(gIncrease, "Percent")
    Range("P3").Value = Format(gDecrease, "Percent")
    Range("P4").Value = gVolume
    
    Next ws
    
End Sub


