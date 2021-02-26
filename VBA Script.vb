sub stock()

    Dim ws as Worksheet

    For Each ws in Worksheets

        Dim Ticker As String

        Dim Total_Stock_Volume as Double
        Total_Stock_Volume = 0

        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        Dim Open_Amount As Double
        Open_Amount = Cells(2,3).Value

        Dim Close_Amount As Double

        Dim Yearly_Change As Double

        Dim Percent_Change As Long

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        Range("I1,P1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percent Change"
        Range("L1").Value = "Total Stock Volume"

        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"

        Range("Q1").Value = "Value"

        For i = 2 to LastRow

            If cells(i+1,1).value <> cells(i,1).Value Then

                Ticker = Cells(i,1).Value

                Close_Amount = Cells(i,6).Value
                Yearly_Change = Close_Amount - Open_Amount
                Percent_Change = Yearly_Change/Open_Amount

                Total_Stock_Volume = Total_Stock_Volume + Cells(i,7).Value

                Range("I" & Summary_Table_Row).Value = Ticker
                Range("J" & Summary_Table_Row).Value = Yearly_Change
                Range("K" & Summary_Table_Row).Value = Percent_Change
                Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                If Yearly_Change < 0 Then
                Range("J"&Summary_Table_Row).Interior.ColorIndex = 3
                Elseif Yearly_Change > 0 Then
                Range("J"&Summary_Table_Row).Interior.ColorIndex = 4
                Else
                Range("J"&Summary_Table_Row).Interior.ColorIndex = 6
                End if
                
                Summary_Table_Row = Summary_Table_Row + 1

                Total_Stock_Volume = 0
                Open_Amount = Cells(i+1,3).Value
            
            Else

                Total_Stock_Volume = Total_Stock_Volume + Cells(i,7).Value

            End If

        Next i

        Range("Q2").Value = Application.WorksheetFunction.Max(Range("J2:J" & Summary_Table_Row))
        Range("Q3").Value = Application.WorksheetFunction.Min(Range("J2:J" & Summary_Table_Row))
        Range("Q4").Value = Application.WorksheetFunction.Max(Range("L2:L" & Summary_Table_Row))
        
    Next ws

End sub