'Set Declarations

Dim Ticker As String
    Dim Volume As Double
    Dim Count As Double
    Dim ws As Worksheet
    Dim str As String
    str = Format(0.99, "Percent")
   
    'ST=SummaryTable
    Dim ST As Integer
    ST = 2
    
    Dim i, j As Long
    Dim LastRow As Long
    
    
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
'Start Counter to 0
   
    Volume = 0
     
    Count = 0
'Set Headers
    Cells(1, 9).Value = "TickerValue"
    Cells(1, 10).Value = "CostChg"
    Cells(1, 11).Value = "PercentChg"
    Cells(1, 12).Value = "Total Annual Vol"
  
'Loop through all rows
    For i = 2 To LastRow
        'Sum values in column 7 same Ticker
        If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
            Volume = Volume + Cells(i, 7).Value
            Count = Count + 1
        
        'Then Hold the volumes and re-set all variable to 0
        Else
            Cells(ST, 9).Value = Cells(i, 1).Value
            Cells(ST, 10).Value = Cells(i, 6).Value - Cells(i - Count, 3).Value
            
            'Color cells on value, pos green, neg red
            If Cells(ST, 10).Value < 0 Then
                Cells(ST, 10).Interior.ColorIndex = 3
            Else
                Cells(ST, 10).Interior.ColorIndex = 4
            End If
            
            'set if conditon to evaluate and divide by 0
            If Cells(i - Count, 3).Value = 0 Then
            Cells(ST, 11).Value = "0"
            
            Else
            
            Cells(ST, 11).Value = (Cells(ST, 10).Value / Cells(i - Count, 3).Value)
            Cells(ST, 11).NumberFormat = "0.00%"
            End If
            'Calcule volume of value in column 7
            Cells(ST, 12).Value = Volume + Cells(i, 7).Value
            
            'reset counter
            Count = 0
            Volume = 0
            ST = ST + 1
        End If
    Next i
End Sub

