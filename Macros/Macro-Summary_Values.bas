Attribute VB_Name = "Module2"
Sub Summary_Values()

Dim WS As Worksheet

For Each WS In ThisWorkbook.Worksheets
WS.Activate



' TO CREATE HEADERS
Dim Greatest_Increase As String
Dim Greatest_Decrease As String
Dim Greatest_Volume As String
    Cells(2, Columns.Count).End(xlToLeft).Offset(0, 3) = "Greatest_Increase"
    Cells(3, Columns.Count).End(xlToLeft).Offset(0, 3) = "Greatest_Decrease"
    Cells(4, Columns.Count).End(xlToLeft).Offset(0, 3) = "Greatest_Volume"


Dim last_col As Range
    Set last_col = Cells(1, Columns.Count).End(xlToLeft).Offset(0, 4)
        last_col.Resize(1, 2).Value = Array("Ticker", "Value")

'FOR LOOP
Lastrow = Cells(Rows.Count, 10).End(xlUp).Row


Dim Percent_Change As String
Dim Total_Stock_Volume As String
    Percent_Change = Range("L1").Value
    Total_Stock_Volume = Range("M1").Value


For I = 2 To Lastrow

    If Cells(I, 12) = WorksheetFunction.Max(Range("L2:L" & Lastrow)) Then
        Output_Greatest_Increase_Ticker = Cells(I, 10).Value
        Output_Greatest_Increase_Value = Cells(I, 12).Value
        Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1) = Output_Greatest_Increase_Ticker
        Cells(2, Columns.Count).End(xlToLeft).Offset(0, 1) = FormatPercent(Output_Greatest_Increase_Value)
        
    ElseIf Cells(I, 12) = WorksheetFunction.Min(Range("L2:L" & Lastrow)) Then
        Output_Greatest_Decrease_Ticker = Cells(I, 10).Value
        Output_Greatest_Decrease_Value = Cells(I, 12).Value
        Cells(3, Columns.Count).End(xlToLeft).Offset(0, 1) = Output_Greatest_Decrease_Ticker
        Cells(3, Columns.Count).End(xlToLeft).Offset(0, 1) = FormatPercent(Output_Greatest_Decrease_Value)
        
    ElseIf Cells(I, 13) = WorksheetFunction.Max(Range("M2:M" & Lastrow)) Then
        Output_Greatest_Volume_Ticker = Cells(I, 10).Value
        Output_Greatest_Volume_Value = Cells(I, 13).Value
        Cells(4, Columns.Count).End(xlToLeft).Offset(0, 1) = Output_Greatest_Volume_Ticker
        Cells(4, Columns.Count).End(xlToLeft).Offset(0, 1) = Output_Greatest_Volume_Value


        
    End If
    
Next I
    
    
Next WS

End Sub

