Sub Stock_Instructions1()

  'we need to declare variable names firstly
    Dim total As Double
    Dim ticker As String
    Dim ticker_summary As Double
    Dim ticker_change As Double
    Dim open_price As Double
    Dim close_price As Double
    
    For Each ws In Worksheets
        total = 0
        ticker_summary = 2
        ticker_change = 2
        
    ' Giving specified  value to the Cells
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
    ' Creating Loops
        For i = 2 To ws.Cells(Rows.Count, 1).End(xlUp).Row
            total = total + ws.Cells(i, 7).Value
            ticker = ws.Cells(i, 1).Value
            open_price = ws.Cells(ticker_change, 3)
            
        ' With If statement sammarize the different ticker value
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                close_price = ws.Cells(i, 6)
                ws.Cells(ticker_summary, 9).Value = ticker
                ws.Cells(ticker_summary, 10).Value = close_price - open_price
             
          ' Set cell to null for avoid dividing by 0
                If open_price = 0 Then
                    ws.Cells(ticker_summary, 11).Value = Null
                Else
                    ws.Cells(ticker_summary, 11).Value = (close_price - open_price) / open_price
                End If
                ws.Cells(ticker_summary, 12).Value = total
                
                ' Create clour loop for green and red
                If ws.Cells(ticker_summary, 10).Value > 0 Then
                    ws.Cells(ticker_summary, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(ticker_summary, 10).Interior.ColorIndex = 3
                End If
                
                ws.Cells(ticker_summary, 11).NumberFormat = "0.00%"
                
                
                total = 0
                ticker_summary = ticker_summary + 1
                ticker_change = i + 1
            End If
            
        Next i

        ws.Columns("J").AutoFit
        ws.Columns("K").AutoFit
        ws.Columns("L").AutoFit

    Next ws
End Sub

Sub Challenge_Instructions()

    Call Stock_Instructions1
'In this some part of code writing, my second referance is: C-L-Nguyen
'https://github.com/c-l-nguyen/the-VBA-of-Wall-Street/tree/master/code/solution_levels


    For Each ws In Worksheets
        ws.Range("O2") = "Greatest % Increase"
        ws.Range("O3") = "Greatest % Decrease"
        ws.Range("O4") = "Greatest Total Volume"
        ws.Range("P1") = "Ticker"
        ws.Range("Q1") = "Value"
        
        Dim max, min As Double
        Dim min_row_index, max_row_index, max_total_index As Integer
        Dim max_total As Double
        
        max = 0
        min = 0
        max_total = 0
        
        For i = 2 To ws.Cells(Rows.Count, 9).End(xlUp).Row
          ' Create if statement for min/max percentage change
            If ws.Cells(i, 11) > max Then
                max = ws.Cells(i, 11)
                max_row_index = i
            End If
            
            If ws.Cells(i, 11) < min Then
                min = ws.Cells(i, 11)
                min_row_index = i
            End If
            
            ' replace the max total volume value with if.
            If ws.Cells(i, 12) > max_total Then
                max_total = ws.Cells(i, 12)
                max_total_index = i
            End If
        Next i
        
        ' Put  the values to specified cells
        ws.Range("P2") = ws.Cells(max_row_index, 9).Value
        ws.Range("P3") = ws.Cells(min_row_index, 9).Value
        ws.Range("P4") = ws.Cells(max_total_index, 9).Value
        
        ws.Range("Q2") = max
        ws.Range("Q3") = min
        ws.Range("Q4") = max_total
        
        ws.Range("Q2").NumberFormat = "0.00%"
        ws.Range("Q3").NumberFormat = "0.00%"

        ws.Columns("O").AutoFit
        ws.Columns("P").AutoFit
        ws.Columns("Q").AutoFit
    
    Next ws
End Sub
