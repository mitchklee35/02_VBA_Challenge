sub Excelchallenge():

    For each ws in worksheets
        Dim Ticker as string
        Dim Yearly_Change as Double
        Dim Percent_Change as Double
        Dim Total_Stock_Volume as Double
        Dim Summary_Table_Row As Integer
        
        Dim Ticker_name as string
        Dim Yearly_open as Double
        Dim Yearly_close as Double
        Dim Yearly_Change_name as string
        Dim Percent_Change_name as string
        Dim Total_Stock_Volume_name as string
        Dim start as Double


        Ticker_name = "Ticker"
        Yearly_Change_name = "Yearly Change"
        Percent_Change_name = "Percent Change"
        Total_Stock_Volume_name = "Total Stock Volume"
        Greatest_Increase = "Greatest % Increase"
        Greatest_Decrease = "Greates % Decrease"
        Greatest_Total_Volume = "Greatest Total Volume"
        Value = "Value"

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        Total_Stock_Volume = 0
        Summary_Table_Row = 2
        Yearly_Change = 0
        Yearly_open = 0
        Yearly_close = 0
        Percent_Change = 0
        Start = 2

        ws.Cells(1, 9).Value = Ticker_name
        ws.Cells(1, 10).Value = Yearly_Change_name
        ws.Cells(1, 11).Value = Percent_Change_name
        ws.Cells(1, 12).Value = Total_Stock_Volume_name
        ws.Cells(2, 14).Value = Greatest_Increase
        ws.Cells(3, 14).Value = Greatest_Decrease
        ws.Cells(4, 14).Value = Greatest_Total_Volume
        ws.Cells(1, 15).Value = Ticker_name
        ws.Cells(1, 16).Value = Value

        ws.Range("K:K").Style = "Percent"
        ws.Range("L:L").Style = "Currency"
        ws.Range("J:J").Style = "Currency"

        For i = 2 To LastRow

            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker = ws.Cells(i, 1).Value
                Yearly_open = ws.Cells(Start, 3).Value
                    
                    If Yearly_open = 0 Then
                    Yearly_open = .001
                    End If

                Yearly_close = ws.Cells(i, 6).Value
                Yearly_Change = Yearly_close - Yearly_open
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                Percent_Change = (Yearly_close - Yearly_open) / Yearly_open


                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
                ws.Range("K" & Summary_Table_Row).Value = Percent_Change   
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume


                Summary_Table_Row = Summary_Table_Row + 1
                Total_Stock_Volume = 0
                Yearly_Change = 0
                Yearly_open = 0
                Yearly_close = 0
                Start = i + 1 
                
            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            End If



        Next i

        LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row

        For j = 2 to LastRow
            if ws.Cells(j ,10).Value >= 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End if
        Next J

        Dim PercentRng as Range
        Dim VolumeRng as Range
        Dim DblMax as Double
        Dim DblMin as Double
        Dim VolMax as Double

        VolMax = 0
        DblMin = 0
        DblMax = 0

        ws.Cells(2, 16).Style = "Percent"
        ws.Cells(3, 16).Style = "Percent"
        ws.Cells(4, 16).Style = "Currency"

        For k = 2 To lastRow
            If ws.Cells(k, 11).Value > DblMax Then
                DblMax = ws.Cells(k, 11).Value
                ws.Cells(2, 15) = ws.Cells(k, 9).Value
                ws.Cells(2, 16) = DblMax
            End If
            If ws.Cells(k, 11).Value < DblMin Then
                DblMin = ws.Cells(k, 11).Value
                ws.Cells(3, 15) = ws.Cells(k, 9).Value
                ws.Cells(3, 16) = DblMin
            End If
            If ws.Cells(k, 12).Value > VolMax Then
                VolMax = ws.Cells(k, 12).Value
                ws.Cells(4, 15) = ws.Cells(k, 9).Value
                ws.Cells(4, 16) = VolMax
            End If
        Next k

    Next ws

end sub