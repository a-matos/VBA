Attribute VB_Name = "Module1"
Sub Stock_Group()
Dim WS As Worksheet

For Each WS In ThisWorkbook.Worksheets
WS.Activate

Dim Ticker As String
Dim Yearly_Change   As Double
Dim Percent_Change As Double
Dim Total_Stock_Volume  As Double
Dim Open_price As Double
Dim Close_price As Double

Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

Dim last_col As Range
Set last_col = Cells(1, Columns.Count).End(xlToLeft).Offset(0, 3)
last_col.Resize(1, 6).Value = Array("Ticker", "Yearly_Change", "Percent_Change", "Total_Stock_Volume", "Open_price", "Close_price")


Total_Stock_Volume = 0
Row = 2

For I = 2 To Lastrow


    If Cells(I, 1).Value <> Cells(I - 1, 1).Value Then
        Open_price = Cells(I, 3)
        Range("n" & Row).Value = Open_price
        Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
        
        
        
    ElseIf Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
        Ticker = Cells(I, 1).Value
        Close_price = Cells(I, 6).Value
        Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
        Yearly_Change = Close_price - Open_price
        Percent_Change = Yearly_Change / Open_price
    
            Range("j" & Row).Value = Ticker
            Range("k" & Row).Value = Yearly_Change
            Range("l" & Row).Value = FormatPercent(Percent_Change)
            Range("m" & Row).Value = Total_Stock_Volume
            Range("o" & Row).Value = Close_price

        Total_Stock_Volume = 0

        Row = Row + 1
        
       
     Else
        Total_Stock_Volume = Total_Stock_Volume + Cells(I, 7).Value
    
            Range("m" & Row).Value = Total_Stock_Volume
    
   
     End If
     

Next I

  
For j = 11 To 12

New_Lastrow = Cells(Rows.Count, 1).End(xlUp).Row

For I = 2 To New_Lastrow


    If Cells(I, j) < 0 Then
        Cells(I, j).Interior.ColorIndex = 3
        
    
    ElseIf Cells(I, j) >= "0" Then
        Cells(I, j).Interior.ColorIndex = 4
        

        
    End If


Next I
Next j

Next WS

End Sub
