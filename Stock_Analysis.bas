Attribute VB_Name = "Module1"
Sub stock_market():


'initializing variables
Dim total As Double
Dim yearly_change As Double
Dim percent_change As Double
Dim index As Long
Dim target As Integer
Dim initial_value As Long
Dim row_count As Long
Dim bonus_row_count As Long
Dim i As Long
Dim j As Long
Dim k As Long
Dim greatest_value As Double
Dim lowest_value As Double
Dim highest_total As Double
Dim ticker1 As String
Dim ticker2 As String
Dim ticker3 As String
Dim ws As Worksheet

For Each ws In Worksheets
    

    'Adding the respective title
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("k1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    'set initial values for ceratin variable
    target = 2
    initial_value = 2
    yearly_change = 0
    total = 0
    percent_change = 0
    greatest_value = 0
    lowest_value = 0
    highest_total = 0
    row_count = ws.Cells(Rows.Count, "A").End(xlUp).Row
    

    
    For index = 2 To row_count
        If (ws.Cells(index + 1, 1).Value = ws.Cells(index, 1).Value) Then
            total = total + ws.Cells(index, 7).Value
        Else
            total = total + ws.Cells(index, 7).Value
            If (total = 0) Then
            ws.Range("I" & target).Value = ws.Cells(index, 1).Value
            ws.Range("J" & target).Value = 0
            ws.Range("K" & target).Value = "%" & 0
            ws.Range("L" & target).Value = total
            Else
                If (ws.Cells(initial_value, 3).Value = 0) Then
                    For new_value = initial_value To index:
                        If (ws.Cells(new_value, 3).Value <> 0) Then
                            initial_value = new_value
                            Exit For
                        End If
                    Next new_value
                End If
                
                'yearly change and percentage Calculation
                yearly_change = (ws.Cells(index, 6).Value - ws.Cells(initial_value, 3))
                percent_change = Round((yearly_change / ws.Cells(initial_value, 3) * 100), 2)
                
                'print the results
                ws.Range("I" & target).Value = ws.Cells(index, 1).Value
                ws.Range("J" & target).Value = yearly_change
                ws.Range("K" & target).Value = "%" & percent_change
                ws.Range("L" & target).Value = total
                
                initial_value = index + 1
                
                'color coding the change
                If yearly_change < 0 Then
                    ws.Range("J" & target).Interior.ColorIndex = 3
                ElseIf yearly_change > 0 Then
                    ws.Range("J" & target).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & target).Interior.ColorIndex = 0
                End If
            End If
             'initialize the next ticker initial value & increment target cell
                
                target = target + 1
                total = 0
                percent_change = 0
                yearly_change = 0
        End If
    Next index
    
    bonus_row_count = ws.Cells(Rows.Count, "I").End(xlUp).Row
   
    
    
    'calculating greatest value
    For i = 2 To bonus_row_count
        If (greatest_value < ws.Cells(i, 11).Value) Then
            greatest_value = ws.Cells(i, 11).Value
            ticker1 = ws.Cells(i, 9)
        End If
    Next i
    'printing the result
        ws.Range("P2").Value = ticker1
        ws.Range("Q2").Value = greatest_value
    
    'calculating lowest value
    For j = 2 To bonus_row_count
        If (lowest_value > ws.Cells(j, 11).Value) Then
            lowest_value = ws.Cells(j, 11).Value
            ticker2 = ws.Cells(j, 9)
        End If
    Next j
    'printing the result
        ws.Range("P3").Value = ticker2
        ws.Range("Q3").Value = lowest_value
        
    'calculating highest total value
    For k = 2 To bonus_row_count
        If (highest_total < ws.Cells(k, 12).Value) Then
            highest_total = ws.Cells(k, 12).Value
            ticker3 = ws.Cells(k, 9)
        End If
    Next k
    'printing the result
        ws.Range("P4").Value = ticker3
        ws.Range("Q4").Value = highest_total
    
    
        
    
Next ws

End Sub
