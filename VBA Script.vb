sub stock()

    Dim ws as Worksheet

    'loop through each workshet within the workbook
    For Each ws in Worksheets

        Dim Ticker As String

        Dim Total_Stock_Volume as Double
        Total_Stock_Volume = 0

        'set summary table row as integer and initial value to 2 where values begin
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2

        'set initial value to open amount from beginning of year for A
        Dim Open_Amount As Double
        Open_Amount = ws.Cells(2,3).Value

        Dim Close_Amount As Double
        Dim Yearly_Change As Double
        Dim Percent_Change As Double

        'grab the last row within the worksheet
        'citation: https://www.excelcampus.com/vba/find-last-row-column-cell/ 
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        'label headers for the new columns
        ws.Range("I1,P1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"

        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"

        ws.Range("Q1").Value = "Value"

        'loop through all rows in the worksheet
        For i = 2 to LastRow

            'if the ticker symbol doesn't match the next symbol then execute
            If ws.Cells(i+1,1).value <> ws.Cells(i,1).Value Then

                'set ticker symbol
                Ticker = ws.Cells(i,1).Value
                
                'set close amount to the close amount from end of year for this ticker
                Close_Amount = ws.Cells(i,6).Value
                
                'set yearly change as difference between closing and opening amount
                Yearly_Change = Close_Amount - Open_Amount
                
                'do a if-then statement to catch any errors on divisor by 0
                If Open_Amount = 0 Then 
                    Percent_Change = 0
                    Else
                    'calculation for percent change
                    Percent_Change = Yearly_Change/Open_Amount
                End if

                'total stock volume should be running total from this stock ticker
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i,7).Value

                'input respective values into the summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'conditional formatting based on yearly change value
                'citation: http://dmcritchie.mvps.org/excel/colors.htm
                If Yearly_Change < 0 Then
                ws.Range("J"&Summary_Table_Row).Interior.ColorIndex = 3
                Elseif Yearly_Change > 0 Then
                ws.Range("J"&Summary_Table_Row).Interior.ColorIndex = 4
                Else
                ws.Range("J"&Summary_Table_Row).Interior.ColorIndex = 6
                End if
                
                'summary table row counter each time ticker symbol changes
                Summary_Table_Row = Summary_Table_Row + 1

                'reset total stock volume to 0 and lock open amount to the next ticker's open amount from beginning of year
                Total_Stock_Volume = 0
                Open_Amount = ws.Cells(i+1,3).Value
            
            Else

                'if it's on the same ticker symbol, keep running total on total stock volume
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i,7).Value

            End If

        Next i

        'grab the max or min value from the summary table and format accordingly
        'citation: https://stackoverflow.com/questions/31906571/excel-vba-find-maximum-value-in-ws.Range-on-specific-sheet/31906916
        ws.Range("Q2").Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row))
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row))
        ws.Range("Q3").NumberFormat = "0.00%"
        ws.Range("Q4").Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row))
        
        'lookup corresponding stock ticker based on max values
        'citation: https://www.educba.com/vba-lookup/
        ws.Range("P2").Value = WorksheetFunction.XLookup(ws.Range("Q2"),ws.Range("K2:K"&Summary_Table_Row),ws.Range("I2:I"&Summary_Table_Row))
        ws.Range("P3").Value = WorksheetFunction.XLookup(ws.Range("Q3"),ws.Range("K2:K"&Summary_Table_Row),ws.Range("I2:I"&Summary_Table_Row))
        ws.Range("P4").Value = WorksheetFunction.XLookup(ws.Range("Q4"),ws.Range("L2:L"&Summary_Table_Row),ws.Range("I2:I"&Summary_Table_Row))

        'Autofit new columns
        'citation: https://docs.microsoft.com/en-us/office/vba/api/excel.ws.Range.autofit
        ws.Range("I:Q").Columns.AutoFit

    Next ws

End sub