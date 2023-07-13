# stock-analysis 1st part
Sub StockAnalysis()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRow As Integer
    
    
    Set ws = ThisWorkbook.Worksheets("2018")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
 
    summaryTableRow = 2
   
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ticker = ws.Cells(i, 1).Value
            
            openingPrice = ws.Cells(i, 3).Value
        End If
        
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            closingPrice = ws.Cells(i, 6).Value
    
            yearlyChange = closingPrice - openingPrice
            percentChange = (yearlyChange / openingPrice) * 100
         
            ws.Cells(summaryTableRow, 9).Value = ticker
            ws.Cells(summaryTableRow, 10).Value = yearlyChange
            ws.Cells(summaryTableRow, 11).Value = percentChange
            ws.Cells(summaryTableRow, 12).Value = totalVolume
            
            If yearlyChange >= 0 Then
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(summaryTableRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
           
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                greatestIncreaseTicker = ticker
            End If
            
            If percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                greatestDecreaseTicker = ticker
            End If
            
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                greatestVolumeTicker = ticker
            End If
    
            summaryTableRow = summaryTableRow + 1
            
            ticker = ""
            openingPrice = 0
            closingPrice = 0
            yearlyChange = 0
            percentChange = 0
            totalVolume = 0
        End If
    Next i
    
   
End Sub

![image](https://github.com/Cindymp/stock-analysis/assets/135760131/aa14ba86-f429-438c-af45-fc1b229810fe)
