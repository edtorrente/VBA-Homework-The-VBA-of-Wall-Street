

Sub stock_analysis_using_VBA():

    'Activate worksheetS

    For Each ws In Worksheets

        'all columns in all worksheets are adjusted to fit column header
                        
'        **NOTE**: Testing Data
'        Worksheets("A").Cells.EntireColumn.AutoFit
'        Worksheets("B").Cells.EntireColumn.AutoFit
'        Worksheets("C").Cells.EntireColumn.AutoFit
'        Worksheets("D").Cells.EntireColumn.AutoFit
'        Worksheets("E").Cells.EntireColumn.AutoFit
'        Worksheets("F").Cells.EntireColumn.AutoFit
'        Worksheets("P").Cells.EntireColumn.AutoFit
    
        Worksheets("2014").Cells.EntireColumn.AutoFit
        Worksheets("2015").Cells.EntireColumn.AutoFit
        Worksheets("2016").Cells.EntireColumn.AutoFit

        'establish column header for each worksheet
        ws.Range("I1").Value = "ticker"
        ws.Range("J1").Value = "yearly change"
        ws.Range("K1").Value = "percent change"
        ws.Range("L1").Value = "total stock volume"
        ws.Range("O2").Value = "greates % increase"
        ws.Range("O3").Value = "greatest % decrease"
        ws.Range("O4").Value = "greatest total volume"
        ws.Range("P1").Value = "ticker"
        ws.Range("Q1").Value = "value"

        'declare variables
        Dim ticker_name As String
        Dim sum_ticker_vol As Double
        Dim year_open As Double
        Dim year_close As Double
        Dim year_change As Double
        Dim last_value As Long
        Dim tally As Long
        Dim percent_change As Double
        Dim max_increase As Double
        Dim max_decrease As Double
        Dim LastRowValue As Long
        Dim max_total_volume As Double

    'initiate the counter
        sum_ticker_vol = 0
        tally = 2
        last_value = 2
        max_increase = 0
        max_decrease = 0
        max_total_volume = 0

      
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row

            ' add to sum_ticker_vol
            sum_ticker_vol = sum_ticker_vol + ws.Cells(i, 7).Value
            'check if still the same ticker name.
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                'set ticker_name
                ticker_name = ws.Cells(i, 1).Value
                'add ticker_name to tally
                ws.Range("I" & tally).Value = ticker_name
                'print ticker_name to tally
                ws.Range("L" & tally).Value = sum_ticker_vol
                'reset
                sum_ticker_vol = 0


                'iniate year_open, year_close, year_change
                year_open = ws.Range("C" & last_value)
                year_close = ws.Range("F" & i)
                year_change = year_close - year_open
                ws.Range("J" & tally).Value = year_change

                'find % change
                If year_open = 0 Then
                    percent_change = 0
                Else
                    year_open = ws.Range("C" & last_value)
                    percent_change = year_change / year_open
                End If
                'format to include percent change
                ws.Range("K" & tally).NumberFormat = "0.000%"
                ws.Range("K" & tally).Value = percent_change

                'conditional formatting in column year_change. Green > 0 and Red < 0.
                If ws.Range("J" & tally).Value >= 0 Then
                    ws.Range("J" & tally).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & tally).Interior.ColorIndex = 3
                End If
            
                'add one to tally
                tally = tally + 1
                last_value = i + 1
                End If
            Next i


            'bonus: finding greatest % increase, greatest % decrease and greatest total volume
            LastRow = ws.Cells(Rows.Count, 11).End(xlUp).Row
        
            'counter
            For i = 2 To ws.Cells(Rows.Count, 11).End(xlUp).Row
                'greatest % increase
                If ws.Range("K" & i).Value > ws.Range("Q2").Value Then
                    ws.Range("Q2").Value = ws.Range("K" & i).Value
                    ws.Range("P2").Value = ws.Range("I" & i).Value
                End If
                'greatest % decrease
                If ws.Range("K" & i).Value < ws.Range("Q3").Value Then
                    ws.Range("Q3").Value = ws.Range("K" & i).Value
                    ws.Range("P3").Value = ws.Range("I" & i).Value
                End If
                'greatest total volume
                If ws.Range("L" & i).Value > ws.Range("Q4").Value Then
                    ws.Range("Q4").Value = ws.Range("L" & i).Value
                    ws.Range("P4").Value = ws.Range("I" & i).Value
                End If

            Next i
           'format to include percent change
            ws.Range("Q2").NumberFormat = "0.000%"
            ws.Range("Q3").NumberFormat = "0.000%"
            
        ' Format Table Columns To Auto Fit
        ws.Columns("I:Q").AutoFit

    Next ws

End Sub


Sub Clear_Contents()

For Each ws In Worksheets
    
        ws.Columns("I:Q").ClearContents
        ws.Columns("I:Q").ClearFormats
    
    Next ws

End Sub

    


