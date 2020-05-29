
' VBA HOMEWORK - THE VBA OF WALL STREET
' Week 2
Sub AnalyzeStocks()


' This sub-routine performs the following functions:
' 1) Loops through all the stocks for a given year in each worksheet
' 2) Computes and displays the following summary information for each stock, within each worksheet:
'    * Ticker Symbol
'    * Yearly Price Change
'    * Yearly Percent Change
'    * Total Stock Volume


    Dim curr_worksheet As Worksheet
    Dim row As Long
    Dim lastRow As Long
    Dim openingPriceRowIndex, summaryRowIndex As Long
    Dim openingPrice, closingPrice, yearlyChange, percentChange As Variant
    Dim currentStockVolume, totalStockVolume As Variant

    Dim currentTicker, nextTicker As String

    For Each curr_worksheet In Worksheets

        With curr_worksheet
        
            lastRow = .Cells(Rows.Count, 1).End(xlUp).row

            openingPriceRowIndex = 2

            totalStockVolume = 0

            .Range("I1").Value = "Ticker"
            .Range("J1").Value = "Yearly Change"
            .Range("K1").Value = "Percent Change"
            .Range("L1").Value = "Total Stock Volume"
            summaryRowIndex = 2


            For row = 2 To lastRow

                currentTicker = .Cells(row, 1).Value
                nextTicker = .Cells(row + 1, 1).Value

                currentStockVolume = .Cells(row, 7).Value
      
                totalStockVolume = totalStockVolume + currentStockVolume

                If currentTicker <> nextTicker Then
        
                    openingPrice = .Cells(openingPriceRowIndex, 3).Value
                    closingPrice = .Cells(row, 6).Value

                    yearlyChange = closingPrice - openingPrice
                    If (openingPrice <> 0) And (yearlyChange <> 0) Then
                        
                        percentChange = (closingPrice - openingPrice) / openingPrice
                    
                    Else
                        
                        percentChange = 0
                        
                    End If

                    .Range("I" & summaryRowIndex) = currentTicker
                    
                    .Range("J" & summaryRowIndex) = yearlyChange
                    .Range("J" & summaryRowIndex).NumberFormat = "0.00"
                    
                    .Range("K" & summaryRowIndex) = percentChange
                    .Range("K" & summaryRowIndex).NumberFormat = "0.00%"
                    
                    .Range("L" & summaryRowIndex) = totalStockVolume
                    .Range("L" & summaryRowIndex).NumberFormat = "#,##0"

                    totalStockVolume = 0
                    summaryRowIndex = summaryRowIndex + 1
                    openingPriceRowIndex = row + 1

                End If

            Next row
         

            Call ApplyConditionalFormatting(curr_worksheet)

            Call DisplayMaxMinStockData(curr_worksheet)
           


        End With

    Next curr_worksheet
    


End Sub

Sub ApplyConditionalFormatting(ws As Worksheet)


' This sub-routine applies conditional formatting to the Yearly Change column in all worksheets

    Dim yearlyChangeRange As Range
    

    Dim positiveChangeFormat, negativeChangeFormat As FormatCondition
    

    Set yearlyChangeRange = ws.Range("J2", ws.Range("J2").End(xlDown))
    
    With yearlyChangeRange
        
   
        .FormatConditions.Delete


        Set positiveChangeFormat = .FormatConditions.Add(xlCellValue, xlGreaterEqual, "=0")
        Set negativeChangeFormat = .FormatConditions.Add(xlCellValue, xlLess, "=0")

        positiveChangeFormat.Interior.Color = vbGreen
        negativeChangeFormat.Interior.Color = vbRed
       
    End With

End Sub

Sub DisplayMaxMinStockData(ws As Worksheet)


' This sub-routine displays stocks with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume" for a given year

   
        .Range("P1") = "Ticker"
        .Range("Q1") = "Value"
        .Range("O2") = "Greatest % Increase"
        .Range("O3") = "Greatest % Decrease"
        .Range("O4") = "Greatest Total Volume"
        
        
        Dim maxPercentIncrease, maxPercentDecrease, maxTotalVolume
        Dim whatToFind, tickerValue
        Dim findCell As Range
        maxPercentIncrease = Application.WorksheetFunction.Max(.Range("K2:K" & Rows.Count))
        .Range("Q2").NumberFormat = "0.00%"
        .Range("Q2") = maxPercentIncrease
        
        maxPercentDecrease = Application.WorksheetFunction.Min(.Range("K2:K" & Rows.Count))
        .Range("Q3").NumberFormat = "0.00%"
        .Range("Q3") = maxPercentDecrease
        

        maxTotalVolume = Application.WorksheetFunction.Max(.Range("L2:L" & Rows.Count))
        .Range("Q4") = maxTotalVolume
        

        whatToFind = FormatPercent(maxPercentIncrease, 2, vbUseDefault, vbUseDefault, vbFalse)
        

        Set findCell = .Range("K2:K" & Rows.Count).Find(What:=whatToFind, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not findCell Is Nothing Then
            

            tickerValue = .Cells(findCell.row, 9)
            .Range("P2").Value = tickerValue
        
        Else
        
 
            .Range("P2").Value = "Not Found"
        
        End If

        whatToFind = FormatPercent(maxPercentDecrease, 2)
        
  
        Set findCell = .Range("K2:K" & Rows.Count).Find(What:=whatToFind, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not findCell Is Nothing Then
            
          
            tickerValue = .Cells(findCell.row, 9)
            .Range("P3").Value = tickerValue
        
        Else
        

            .Range("P3").Value = "Not Found"
        
        End If
  
        whatToFind = Format(maxTotalVolume, "#,##0")
         
 
        Set findCell = .Range("L2:L" & Rows.Count).Find(What:=whatToFind, LookIn:=xlValues, LookAt:=xlWhole)
        
        If Not findCell Is Nothing Then
        

            tickerValue = .Cells(findCell.row, 9)
            .Range("P4").Value = tickerValue
            
        Else
        

            .Range("P4").Value = "Not Found"
        
        End If

    End With


End Sub