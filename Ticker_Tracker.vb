Sub AlphaStock():

'Defining variables
Dim ticker As String
Dim line As Integer
Dim i As Long
Dim LastRow As Long
Dim Record_Start As Double
Dim Record_End As Double
Dim change As Double
Dim sum As Double
Dim percent_change As Double
Dim vol As Double

'Create for loop for all worksheets
For Each ws In Worksheets
    'Initialize result grid
    line = 1
    'Write out result header
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Volume of Stock"

    'Find last row of worksheet
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    'Initialize total volume
    vol = 0
    great_decrease = 0
    great_increase = 0
    'Initialize opening value
    Record_Start = ws.Cells(2, 3).Value
    'Into the data
    For i = 2 To LastRow

    'Add to total volume
    vol = vol + ws.Cells(i, 7).Value
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'move to new line for results
            line = line + 1
    
    
            'input ticker name
            ticker = ws.Cells(i, 1).Value
            'output ticker name to result grid
            ws.Cells(line, 9).Value = ticker
    
            'input closing value
            Record_End = ws.Cells(i, 6).Value
            'calculate change
            change = Record_End - Record_Start
            'output change
            ws.Cells(line, 10).Value = change
            'calcuate sum for percent change
            sum = Record_End + Record_Start
        
            'fix for if final and initial values = 0 & if conditional exporting percent change
            If sum = 0 Then
                'if both = 0 exports percent change as 0
                percent_change = 0
                'formatting
                ws.Cells(line, 11).Value = percent_change
                ws.Cells(line, 11).Style = "Percent"
                'inputs new opening value
                Record_Start = ws.Cells(i + 1, 3).Value
            
            Else
            
                If Record_Start <> 0 Then
                    'calculates percent change and unifies it so we get if the chage was positive or negative
                    percent_change = (Record_End / Record_Start) - 1
                Else
                    'corrects for dividing by zero = inf
                    percent_change = "1000000000"
                End If
            
                'finding greatest and least % change
                If percent_change > great_increase Then
                    great_increase = percent_change
                    great_ticker = ticker
                ElseIf percent_change < great_decrease Then
                    great_decrease = percent_change
                    least_ticker = ticker
                End If

                'formatting including if percent change is infinite
                if percent_change > 100000000 Then
                    ws.Cells(line,11).Value = "Infinite"
                Else
                    ws.Cells(line, 11).Value = percent_change
                    ws.Cells(line, 11).Style = "Percent"
                End if

                'inputs new opening value
                Record_Start = ws.Cells(i + 1, 3).Value
            End If
        
            'formatting for color of cell
            If ws.Cells(line, 10).Value < 0 Then
                ws.Cells(line, 10).Interior.ColorIndex = 3
            ElseIf ws.Cells(line, 10).Value > 0 Then
                ws.Cells(line, 10).Interior.ColorIndex = 4
            End If

        
        
            'output total volume of trades
            ws.Cells(line, 12).Value = vol
            'finding greatest volume of trades
            If vol > greatest_vol Then
                greatest_vol = vol
                vol_ticker = ticker
            End If
            'resets for next ticker
            vol = 0
        
        End If
    
    Next i

    'CHALLENGE ACCEPTED-----------------------------------
    'Outputting maximum values for each worksheet
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"


    ws.Cells(2, 16).Value = great_ticker
    ws.Cells(3, 16).Value = least_ticker
    ws.Cells(4, 16).Value = vol_ticker
    'Correcting for change with dividing by zero
    If great_increase > 20000 Then
        ws.Cells(2, 17).Value = "Infinity"
    Else
        ws.Cells(2, 17).Value = great_increase
        ws.Cells(2, 17).Style = "Percent"
    End If

    ws.Cells(3, 17).Value = great_decrease
    ws.Cells(3, 17).Style = "Percent"
    ws.Cells(4, 17).Value = greatest_vol
    '-----------------------------------------------------
'goes to next worksheet
Next ws



End Sub