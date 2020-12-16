VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_analysis()
    'Dimensions
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim change As Double
    Dim start As Long
    Dim lastRow As Long
    Dim days As Integer
    Dim avgChange As Double
    
    'Column titles
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    'Initially
    j = 0
    total = 0
    change = 0
    start = 2
    
    'Get last row #
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        'Print if ticker has changed
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            total = total + Cells(i, 7).Value
            If total = 0 Then
                'Put results in column
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = "%" & 0
                Range("L" & 2 + j).Value = 0
            Else
                'Go to first non-zero
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
                'Calculate change in price
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = Round((change / Cells(start, 3) * 100), 2)
                
                'For the next ticker
                start = i + 1
                'Display in column
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = Round(change, 2)
                Range("K" & 2 + j).Value = "%" & percentChange
                Range("L" & 2 + j).Value = total
                
                'Conditional format: green for positive change, red for negative change
                If change > 0 Then
                    Range("J" & 2 + j).Interior.ColorIndex = 4
                ElseIf change < 0 Then
                    Range("J" & 2 + j).Interior.ColorIndex = 3
                Else
                    Range("J" & 2 + j).Interior.ColorIndex = 0
                End If
                
            End If
            
            'Reset values; next ticker
            total = 0
            change = 0
            j = j + 1
            days = 0
        'Get total for each ticker
        Else
            total = total + Cells(i, 7).Value
        End If
        
    Next i
    
End Sub

