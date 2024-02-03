Sub Stock_Data()
    Dim ws As Worksheet
    Dim Ticker_Symbol As String
    Dim LastRow As Long
    Dim Start_Open_Price As Double
    Dim Open_Price_Captured As Boolean
    Dim End_Close_Price As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Total_Stock_Volume As Double
    Dim Summary_Table_Row As Integer
    
    For Each ws In Worksheets
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Summary_Table_Row = 2
        Open_Price_Captured = False
        
        ' Clear existing summary table data
        ws.Range("J:O").Clear
        
        ' Set column names for new values
        ws.Cells(1, 10).Value = "Ticker"
        ws.Cells(1, 11).Value = "Yearly Change"
        ws.Cells(1, 12).Value = "Percent Change"
        ws.Cells(1, 13).Value = "Total Stock Volume"
        ws.Cells(1, 17).Value = "Ticker"
        ws.Cells(1, 18).Value = "Value"
        
        For i = 2 To LastRow
            If Open_Price_Captured = False Then
                Start_Open_Price = ws.Cells(i, 3).Value
                Open_Price_Captured = True
            End If
            
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                Ticker_Symbol = ws.Cells(i, 1).Value
                ws.Cells(Summary_Table_Row, 10).Value = Ticker_Symbol
                
                End_Close_Price = ws.Cells(i, 6).Value
                Yearly_Change = End_Close_Price - Start_Open_Price
                ws.Cells(Summary_Table_Row, 11).Value = Yearly_Change
                
                Percent_Change = Yearly_Change / Start_Open_Price
                ws.Cells(Summary_Table_Row, 12).Value = Percent_Change
                ws.Cells(Summary_Table_Row, 12).NumberFormat = "0.00%"
                
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                ws.Cells(Summary_Table_Row, 13).Value = Total_Stock_Volume
                
                If Yearly_Change < 0 Then
                    ws.Cells(Summary_Table_Row, 11).Interior.Color = RGB(255, 0, 0)
                Else
                    ws.Cells(Summary_Table_Row, 11).Interior.Color = RGB(0, 255, 0)
                End If
                
                Summary_Table_Row = Summary_Table_Row + 1
                
                ' Reset variables for next ticker
                End_Close_Price = 0
                Yearly_Change = 0
                Percent_Change = 0
                Total_Stock_Volume = 0
                Open_Price_Captured = False
            Else
                End_Close_Price = ws.Cells(i, 6).Value
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Find max increase, decrease, and total volume
        Dim Summary_Last_Row As Long
        Dim Max_Increase As Double
        Dim Max_Decrease As Double
        Dim Max_Total As Double
        
        Summary_Last_Row = ws.Cells(ws.Rows.Count, 10).End(xlUp).Row
        Max_Increase = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Last_Row))
        Max_Decrease = WorksheetFunction.Min(ws.Range("L2:L" & Summary_Last_Row))
        Max_Total = WorksheetFunction.Max(ws.Range("M2:M" & Summary_Last_Row))
        
        ' Output max increase, decrease, and total volume
        ws.Cells(2, 16).Value = "Greatest % Increase"
        ws.Cells(3, 16).Value = "Greatest % Decrease"
        ws.Cells(4, 16).Value = "Greatest Total Volume"
        
 
        
        ' Formatting
        ws.Cells(2, 18).NumberFormat = "0.00"
        ws.Cells(3, 18).NumberFormat = "0.00"
        
 For i = 2 To Summary_Last_Row
            For j = 12 To 12
                If ws.Cells(i, j).Value = Max_Increase Then
                    ws.Cells(2, 18).Value = Max_Increase * 100
                    ws.Cells(2, 17).Value = ws.Cells(i, j - 2).Value
                ElseIf ws.Cells(i, j).Value = Max_Decrease Then
                    ws.Cells(3, 18).Value = Max_Decrease * 100
                    ws.Cells(3, 17).Value = ws.Cells(i, j - 2).Value
                End If
            Next j
            For j = 13 To 13
                If ws.Cells(i, j).Value = Max_Total Then
                    ws.Cells(4, 18).Value = Max_Total
                    ws.Cells(4, 17).Value = ws.Cells(i, j - 3).Value
                End If
            Next j
        Next i

 
        
    Next ws
End Sub