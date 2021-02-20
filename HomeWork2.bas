Attribute VB_Name = "Module1"
   Sub credit_card()


    Dim Ticker As String
    Dim i, OpeningValue, ClosingValue, yearly_change, Percent_Change As Double
    Dim ws As Worksheet
    
    Dim Vol_Total As Double
    Vol_Total = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    ws.Activate
    Summary_Table_Row = 2
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
        For i = 2 To LastRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                Ticker = ws.Cells(i, 1).Value
                
                ws.Cells(2, "N").Value = Cells(2, 3).Value
                
                
                ClosingValue = ws.Cells(i, 6)
                
                
                OpeningValue = ws.Cells(i + 2, 3)
                
                
                yearly_change = Cells(Summary_Table_Row, "N").Value - Cells(Summary_Table_Row, "O").Value
                
                
                Cells(i, "K").Value = "Yes"
 
                Vol_Total = Vol_Total + Cells(i, 7).Value


                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("N" & Summary_Table_Row).Value = OpeningValue
                
                ws.Range("O" & Summary_Table_Row).Value = ClosingValue
                
                ws.Range("L" & Summary_Table_Row).Value = Vol_Total
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
 '               ws.Range("K" & Summary_Table_Row).Value = Percent_Change

                Summary_Table_Row = Summary_Table_Row + 1
      

                Vol_Total = 0
            Else
                Vol_Total = Vol_Total + Cells(i, 7).Value

        End If
     Next i
     
    Next ws
End Sub
