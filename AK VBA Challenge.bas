Attribute VB_Name = "Module1"
Sub ticker()

'Define Variables

Dim stock_name As String

Dim yearly_change As Double
yearly_change = 0

Dim percent_change As Double
percent_change = 0

Dim total_stock_volume As Double
total_stock_volumne = 0

Dim start_value As Double

'Loop through Worksheets
For Each ws In Worksheets
    'From TA Mark
    ws.Activate
    total_stock_volume = 0
    
'From TA Mark
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2
'From TA Mark- Define Last Row
lastRow = Cells(Rows.Count, "A").End(xlUp).Row

'Create Headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
    
    'Go through rows
    For i = 2 To lastRow
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            start_value = Cells(i, 3).Value
        End If
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'Formulas to create number/name of each row
            stock_name = Cells(i, 1).Value
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
            'Yearly_change and Percent_Change from TA Mark
            yearly_change = Cells(i, 6).Value - start_value
                If start_value = 0 Then
                    start_value = 1
                End If
            percent_change = yearly_change / start_value
        'print each row
            Range("I" & Summary_Table_Row).Value = stock_name
            Range("J" & Summary_Table_Row).Value = yearly_change
            Range("K" & Summary_Table_Row).Value = percent_change
            Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = total_stock_volume
            
            
        'Color Rows of yearly_change
        If Cells(Summary_Table_Row, 10).Value > 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
        Else
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
        End If
        'Color Rows of percent_change
        If Cells(Summary_Table_Row, 11).Value > 0 Then
            Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
        Else
            Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
        End If
        
        'Move to next row
            Summary_Table_Row = Summary_Table_Row + 1
        'Restart formulas
            total_stock_volume = 0
            yearly_change = 0
            percent_change = 0
            
'If the next row is the same brand
        Else
            total_stock_volume = total_stock_volume + Cells(i, 7).Value
        End If
    Next i
    
    
'Calculatations Here
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest Percent Increase"
    Cells(3, 14).Value = "Greatest Percent Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    'Greatest % Increase
    lastRow_Calc = Cells(Rows.Count, "J").End(xlUp).Row

        For i = 2 To lastRow_Calc:
            Dim max As Double
            max = Application.WorksheetFunction.max(Columns("K"))
        Cells(2, 16).Value = max
        Cells(2, 16).NumberFormat = ("0.00%")
        Next i
    'Greatest % Decrease
        For i = 2 To lastRow_Calc:
            Dim min As Double
            min = Application.WorksheetFunction.min(Columns("K"))
        Cells(3, 16).Value = min
        Cells(3, 16).NumberFormat = ("0.00%")
        Next i
    'Greatest Total Volume
        For i = 2 To lastRow_Calc:
            Dim max_volume As Double
            max_volume = Application.WorksheetFunction.max(Columns("L"))
        Cells(4, 16).Value = max_volume
        Next i
Next ws

End Sub


