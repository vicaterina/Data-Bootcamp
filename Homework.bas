Attribute VB_Name = "Module2"
Sub AdvanceSheets()
    Dim Sheet As Worksheet
    Application.ScreenUpdating = False
    For Each Sheet In Worksheets
        Sheet.Select
        Call moderate
         Cells(1, 15).Value = "Ticker"
         Cells(1, 16).Value = "Value"
         Cells(2, 14).Value = "Greatest % Increase"
         Cells(3, 14).Value = "Greatest % Decrease"
         Cells(4, 14).Value = "Greatest Total Volume"
         
         Cells(2, 16).Value = Application.WorksheetFunction.Max(Columns("K"))
         Cells(3, 16).Value = Application.WorksheetFunction.Min(Columns("K"))
         Cells(4, 16).Value = Application.WorksheetFunction.Max(Columns("L"))
        
   
        
    Next
    Application.ScreenUpdating = True
End Sub






Sub moderate()

Dim volume As Double
Dim output_cell As Integer
Dim open_price As Double
Dim close_price As Double
Dim yearly_change As Double
Dim pct_change As Double


lastrow = Range("A1").End(xlDown).Row

output_cell = 1
volume = 0


Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

For i = 2 To lastrow
volume = volume + Cells(i, 7).Value
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
    counter = counter + 1
    Else
        volume = volume + Cells(i + 1, 7)
        output_cell = output_cell + 1
        
        open_price = Cells(i - counter, 3).Value
        closing_price = Cells(i, 6).Value
        yearly_change = closing_price - open_price
            If open_price <> 0 Then
            pct_change = (closing_price - open_price) / open_price
            Else
            pct_change = 0
            End If
        Cells(output_cell, 9).Value = Cells(i, 1).Value
        Cells(output_cell, 12).Value = volume
        Cells(output_cell, 10).Value = yearly_change
        Cells(output_cell, 11).Value = pct_change
    
        counter = 0
        volume = 0
    
    End If
main_counter = main_counter + 1
Next i

    

 
End Sub


