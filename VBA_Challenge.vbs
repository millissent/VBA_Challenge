VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alpha_testing()

       ' variable to hold the worksheet name
        Dim smWS As Worksheet
        
    ' to loop through all of the worksheets in the Workbook
    For Each smWS In Worksheets
        
        ' find the last row
        Dim lastRow As Long
        lastRow = smWS.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' find the last column
        Dim lastCol As Integer
        lastCol = smWS.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' add new headers to each worksheet
        smWS.Range("I1").Value = "Ticker"
        smWS.Range("J1").Value = "Yearly Change"
        smWS.Range("K1").Value = "Percent Change"
        smWS.Range("L1").Value = "Total Stock Volume"
        
        ' set variables needed for calculations
        Dim tickerName As String
        tickerName = " "
        Dim totalVolume As Double
        totalVolume = 0
        Dim openPrice As Double
        openPrice = 0
        Dim closePrice As Double
        closePrice = 0
        Dim yearlyPriceChange As Double
        yearlyPriceChange = 0
        Dim yearlyPriceChangePercent As Double
        yearlyPriceChangePercent = 0
                      
        ' Set location for variables
        Dim summaryTableRow As Long
        summaryTableRow = 2
        
  ' Set initial value of beginning stock value for the worksheets
  openPrice = smWS.Cells(2, 3).Value
  
  ' loop from the beginning of the first worksheet to the end of the last
  For i = 2 To lastRow
  
  ' Check to see if we are still are on same ticker name
  If smWS.Cells(i + 1, 1).Value <> smWS.Cells(i, 1).Value Then
  
        ' Set ticker name
        tickerName = smWS.Cells(i, 1).Value
        
       ' Calculate closing price and yearly price change variables
       closePrice = smWS.Cells(i, 6).Value
       yearlyPriceChange = closePrice - openPrice
       
        ' Ensure that openPrice is not zero so that we don't receive an error message
        If openPrice <> 0 Then
        yearlyPriceChangePercent = (yearlyPriceChange / openPrice) * 100
       
       End If
        
        ' Add to ticker total volume
        totalVolume = totalVolume + smWS.Cells(i, 7).Value
        
        ' Print ticker name in Column I
        smWS.Range("I" & summaryTableRow).Value = tickerName
        
        ' Print yearly price change in column J
        smWS.Range("J" & summaryTableRow).Value = yearlyPriceChange
        
         ' Fill in Yearly Change with color to denote positive or negative change
        If yearlyPriceChange > 0 Then
            smWS.Range("J" & summaryTableRow).Interior.ColorIndex = 4

        ElseIf yearlyPriceChange <= 0 Then
            smWS.Range("J" & summaryTableRow).Interior.ColorIndex = 3
        
        End If
        
        ' Print yearly price change as a percent in column K
        smWS.Range("K" & summaryTableRow).Value = (CStr(yearlyPriceChangePercent) & "%")
        
        ' Print total stock volume in column L
        smWS.Range("L" & summaryTableRow).Value = totalVolume
        
        ' Add 1 to summary table row count
        summaryTableRow = summaryTableRow + 1
        
        ' Calculate next open price
        openPrice = smWS.Cells(i + 1, 3).Value
          
       End If

    Next i

    Next smWS

End Sub

