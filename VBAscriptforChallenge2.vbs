Sub Tickers():

'convert text to number format in col B (date)
'Range("b:b") = Range("b:b").Value
'Dim DateRange As Range
'Set DateRange = Range("b:b")

'Dim earliestDate As Integer
'earliestDate = Application.Match(Application.Min(DateRange), DateRange)

'Dim newestDate As Integer
'newestDate = Application.Match(Application.Max(DateRange), DateRange)
'Yearly_Change = newestDate - earliestDate
'            Range("j" & Summary_Table_Row).Value = Yearly_Change

'run on every worksheet at once
For Each ws In Worksheets

'new column headers
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"
ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"

'format columns and number formats
ws.Range("I:L").EntireColumn.AutoFit
ws.Range("o:o").EntireColumn.AutoFit
ws.Range("J:J").NumberFormat = "0.00"
ws.Range("K:K").NumberFormat = "0.00%"
ws.Range("Q2:Q3").NumberFormat = "0.00%"


Dim Yearly_Change As Single
Dim Percent_Changeange As Integer
Dim Volume_Total As Double
'Dim Greatest_Increase As Integer
'Dim Greatest_Decrease As Integer
'Dim Greatest_Volume As Double


Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'lastrow = Cells(Rows.Count, 1).End(xlUp).Row
lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
Volume_Total = 0
openPrice = 0

Greatest_Increase_Value = 0
Greatest_Decrease_Value = 0
Greatest_Volume_Value = 0
Greatest_Increase_Ticker = ""
Greatest_Decrease_Ticker = ""
Greatest_Volume_Ticker = ""


    For i = 2 To lastrow
    
        If openPrice = 0 Then
            openPrice = ws.Cells(i, "C")
        End If
        
    
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'Ticker
            ws.Cells(Summary_Table_Row, "I") = ws.Cells(i, 1).Value
            
            'Yearly Change
            Yearly_Change = ws.Cells(i, 6) - openPrice
            ws.Cells(Summary_Table_Row, "J") = Yearly_Change
            'ws.Cells(Summary_Table_Row, "J").NumberFormat = "0.00"
            If Yearly_Change > 0 Then
                ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
            Else
                ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
            End If
            
            'Percent Change
            'Cells(Summary_Table_Row, "K") = Yearly_Change / openPrice * 100
            Percent_Change = Yearly_Change / openPrice
            ws.Cells(Summary_Table_Row, "K") = Percent_Change
            'ws.Cells(Summary_Table_Row, "K").NumberFormat = "0.00%"
            
            'Total Stock Volume
            Volume_Total = Volume_Total + ws.Cells(i, 7).Value
            ws.Range("L" & Summary_Table_Row).Value = Volume_Total
            
            'Greatest % Increase
            If Percent_Change > Greatest_Increase_Value Then
                Greatest_Increase_Value = Percent_Change
                Greatest_Increase_Ticker = ws.Cells(i, "A")
            End If
            
            'Greatest % Decrease
            If Percent_Change < Greatest_Decrease_Value Then
                Greatest_Decrease_Value = Percent_Change
                Greatest_Decrease_Ticker = ws.Cells(i, "A")
            End If
            
                'Greatest Total Volume
            If Volume_Total > Greatest_Volume_Value Then
                Greatest_Volume_Value = Volume_Total
                Greatest_Volume_Ticker = ws.Cells(i, "A")
            End If

            
            'Reset
            Summary_Table_Row = Summary_Table_Row + 1
            Volume_Total = 0
            openPrice = 0
            
        End If
    Next i
    
ws.Range("P2") = Greatest_Increase_Ticker
ws.Range("Q2") = Greatest_Increase_Value
ws.Range("P3") = Greatest_Decrease_Ticker
ws.Range("Q3") = Greatest_Decrease_Value
ws.Range("P4") = Greatest_Volume_Ticker
ws.Range("Q4") = Greatest_Volume_Value

    
'Greatest Increase/Decrease/Volume
'Range("Q2") = WorksheetFunction.Max(Range("K:k"))     '''''''add ws
'Range("Q3") = WorksheetFunction.Min(Range("K:k"))     '''''''add ws
'Range("Q2:q3").NumberFormat = "0.00%"                   '''''''add ws
'Range("Q4") = WorksheetFunction.Max(Range("L:L"))     '''''''add ws

  
Next ws

End Sub

