Attribute VB_Name = "Module1"
Sub stockAnalysis()

    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim summaryTableRowIndex As Integer
    Dim lastRow As Long
    Dim i As Long
    Dim ws As Worksheet
    

    For Each ws In ThisWorkbook.Worksheets
        
    summaryTableRowIndex = 2
    openingPrice = ws.Range("C2").Value
    greatestIncreaseValue = 0
    greatestDecreaseValue = 0
    greatestVolumeValue = 0

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = ticker Then
            totalVolume = totalVolume + ws.Cells(i, 7).Value

        Else
                
        If Not ticker = "" Then

        yearlyChange = closingPrice - openingPrice
        percentChange = yearlyChange / openingPrice

        ws.Range("I" & summaryTableRowIndex).Value = ticker
        ws.Range("J" & summaryTableRowIndex).Value = yearlyChange
        ws.Range("K" & summaryTableRowIndex).Value = percentChange
        ws.Range("L" & summaryTableRowIndex).Value = totalVolume
                    
        summaryTableRowIndex = summaryTableRowIndex + 1
        End If
            
            ticker = ws.Cells(i, 1).Value
            openingPrice = ws.Cells(i, 3).Value
            closingPrice = ws.Cells(i, 6).Value
            totalVolume = ws.Cells(i, 7).Value
                
        End If
            closingPrice = ws.Cells(i, 6).Value
            
    Next i
        
        yearlyChange = closingPrice - openingPrice
        percentChange = yearlyChange / openingPrice
        
        ws.Range("I" & summaryTableRowIndex).Value = ticker
        ws.Range("J" & summaryTableRowIndex).Value = yearlyChange
        ws.Range("K" & summaryTableRowIndex).Value = percentChange
        ws.Range("L" & summaryTableRowIndex).Value = totalVolume
        
        
    Next ws
    
End Sub
