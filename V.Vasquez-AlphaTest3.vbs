Sub WorksheetLoop()
Dim ws As Worksheet
Dim Ticker As String
Dim Yearlychange As Double
Dim Percentchange As Double
Dim Openchange As Double
Openchange = Cells(2, 3).Value
Dim TotalSV As Double
TotalSV = 0
Dim i As Double
Dim SumTableRow As Integer
SumTableRow = 2
Dim LRow As Long
LRow = Cells(Rows.Count, 1).End(xlUp).Row

For Each ws In Worksheets
        Worksheets(ws.Name).Activate
        SumTableRow = 2
        For i = 2 To LRow

            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                Ticker = Cells(i, 1).Value
                Dim Closechange As Double
                Closechange = Cells(i, 6).Value
                Yearlychange = Closechange - Openchange
                
                    If Openchange = 0 Then
                        Cells(SumTableRow, 11).Value = "N/A"
                    Else
                        Percentchange = (Closechange - Openchange) / Openchange
                    End If

                Openchange = Cells(i + 1, 3).Value
                TotalSV = TotalSV + Cells(i, 7).Value
                Range("I" & SumTableRow).Value = Ticker
                Range("J" & SumTableRow).Value = Yearlychange
                    If Cells(SumTableRow, 10).Value < 0 Then
                    Cells(SumTableRow, 10).Interior.ColorIndex = 4
                        
                    Else
                    Cells(SumTableRow, 10).Interior.ColorIndex = 3

                    End If
                Range("K" & SumTableRow).Value = Format(Percentchange, "0.00%")
                Range("L" & SumTableRow).Value = TotalSV
                SumTableRow = SumTableRow + 1
                TotalSV = 0
                    
            Else

                TotalSV = TotalSV + Cells(i, 7).Value
                    
            End If
            
        Next i
   
Next ws

End Sub


