Sub Stocks()

Dim ws As Worksheet
    
    For Each ws In Worksheets
        'Page Setup
            ws.Activate
            
                'Headers
                    Cells(1, 9).Value = "Ticker"
                    Cells(1, 10).Value = "Yearly Change"
                    Cells(1, 11).Value = "Percent Change"
                    Cells(1, 12).Value = "Total Stock Volume"
                    
                'Column Format
                    Columns(12).Select
                    Selection.Style = "Comma"
                    Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
                    
          
        'Variables
            Dim result As Integer
            result = 2
        
            Dim total As Double
            total = 0
              
            Dim last As Long
            last = Cells(Rows.Count, 1).End(xlUp).Row
              
            Dim start As Double
            start = Cells(2, 3).Value
            
            Dim fin As Double
            Dim change As Double
                           
        For i = 2 To last
    
                'next cell the same as current cell
                If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
                    
                    'add line amount to total
                        total = total + Cells(i, 7).Value
                        
                Else 'next cell is different from current cell
                    
                    
                        'add line amount to total,
                            total = total + Cells(i, 7).Value
                        
                        'write ticker name and total to results chart
                            Cells(result, 9).Value = Cells(i, 1).Value
                            Cells(result, 12).Value = total
                        
                        'calculate year change
                            fin = Cells(i, 6).Value
                            change = fin - start
                        
                        'write to results chart
                            Cells(result, 10).Value = change
                            
                                'in case of a null start value
                                If start = 0 Then
                                    Cells(result, 11).Value = "n/a"
                                Else
                                    Cells(result, 11).Value = change / start
                                End If
                                Cells(result, 11).NumberFormat = "0%"
                            
                            'otherwise format cells based on change amount
                            If change > 0 Then
                                Cells(result, 10).Interior.Color = vbGreen
                            ElseIf change < 0 Then
                                Cells(result, 10).Interior.Color = vbRed
                            Else
                                Cells(result, 10).Interior.Color = vbYellow
                            End If
                                                
                    'reset total to zero, update result row
                        total = 0
                        result = result + 1
                        start = Cells(i + 1, 3).Value
                        
                
                End If
    
        Next i
    
    Next ws

End Sub




