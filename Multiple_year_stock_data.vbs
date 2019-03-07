Sub Easy_Level():

Dim c As Integer
Dim n As Integer
Dim ws As Worksheet


' wont repeat worksheets.

For Each ws In ThisWorkbook.Sheets
    With ws
    
    ' define variables.
    Dim rowcount As Long
    Dim ticker As String
    Dim total As Variant
    total = CDec(total)
    Dim j As Long
    j = 2
  
' getting an error for rowcount not sure why.
rowcount = ws.Cells(Rows.Count, "A").End(xlUp).Row


        
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Total Stock Volume"
        
            For i = 2 To 71226
                ticker = ws.Cells(i, 1).Value
                total = total + ws.Cells(i, 7).Value
                    If ticker <> ws.Cells(i + 1, 1).Value Then
                        ws.Cells(j, 9).Value = ticker
                        ws.Cells(j, 10).Value = total
                        ticker = 0
                        total = 0
                        j = j + 1
                    End If
            Next i
    End With
    
Next ws

End Sub


        
' MODERATE CODE
'' Need: conditional formating
    
    ' ActiveWorksheet.Columns(10).numberFromat = "#,##0"
    
    
' MODERATE LOOP

'add j variable to find 2nd cell!
        ' For i = 2 To 70926
            'yearly_change = Cells(i, 2) - Cells(i + j, 5)
            
        'Next i
'# Next n


