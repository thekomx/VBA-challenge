Attribute VB_Name = "VBAStocks"
Sub VBAStocks()
    '===== Give columns a name =====
    Const Col_Ticker As Integer = 1 'The number given according to the data column
    Const Col_Date As Integer = 2
    Const Col_Open As Integer = 3
    Const Col_High As Integer = 4
    Const Col_Low As Integer = 5
    Const Col_Close As Integer = 6
    Const Col_Volume As Integer = 7
    
    Const Col_Chg_Ticker As Integer = 9
    Const Col_Chg_Value As Integer = 10
    Const Col_Chg_Percent As Integer = 11
    Const Col_Chg_Volume As Integer = 12
    
    Const Col_Grt_X As Integer = 14
    Const Col_Grt_Ticker As Integer = 15
    Const Col_Grt_Value As Integer = 16
    '=================================
    
    Const Row_Start As Integer = 2  'First row that contain data (after header row)
    
    Dim Ticker As String
    Dim Ticker_First_Row As Double  'First row of a ticker
    
    Dim WS As Worksheet
    Dim Row_Total As Double         'Sheet's total rows
    Dim RowX As Double
    Dim Row_Update As Double        'The row to put data in (for the "Yearly Change" part)
    
    '== Variables for temporary use ==
    Dim Perc_Change As Double
    Dim Val_Change As Double
    Dim Val_Temp As Double
    Dim Row_Temp As Double
    '=================================
    
    
    
    
    'Iterate throught all sheets
    For Each WS In Worksheets
        With WS
            
            '===== Insert header for "Yearly Change" output =====
            .Cells(Row_Start - 1, Col_Chg_Ticker).Value = "Ticker"
            .Cells(Row_Start - 1, Col_Chg_Value).Value = "Yearly Change"
            .Cells(Row_Start - 1, Col_Chg_Percent).Value = "Percent Change"
            .Cells(Row_Start - 1, Col_Chg_Volume).Value = "Total Stock Volume"
            '====================================================
        
            Ticker = .Cells(Row_Start, Col_Ticker).Value
            Ticker_First_Row = Row_Start
            Row_Update = Row_Start - 1
            
            Row_Total = .Cells(Rows.Count, Col_Ticker).End(xlUp).Row    'Get the number of rows of the sheet
            
            For RowX = Row_Start To Row_Total
            
                '##### Into this condition only if the current row "Ticker" has change OR if it the last row of the sheet #####
                If (RowX = Row_Total) Or (.Cells(RowX, Col_Ticker).Value <> Ticker) Then
                
                    If RowX <> Row_Total Then Row_Temp = RowX - 1 Else Row_Temp = RowX
                    
                    '===== Calculate the price change and the percentage =====
                    Val_Change = .Cells(Row_Temp, Col_Close).Value - .Cells(Ticker_First_Row, Col_Close).Value
                    If Val_Change = 0 Then
                        Perc_Change = 0
                    ElseIf .Cells(Ticker_First_Row, Col_Close).Value = 0 Then 'This to prevent "divided by zero"
                        Perc_Change = Val_Change
                    Else
                        Perc_Change = Val_Change / .Cells(Ticker_First_Row, Col_Close).Value
                    End If
                    '=========================================================
                    
                    '===== Put data in to the "Yearly Change" part =====
                    Row_Update = Row_Update + 1
                    .Cells(Row_Update, Col_Chg_Ticker).Value = Ticker
                    .Cells(Row_Update, Col_Chg_Value).Value = Val_Change
                    .Cells(Row_Update, Col_Chg_Percent).Value = Perc_Change
                    .Cells(Row_Update, Col_Chg_Volume).Value = WorksheetFunction.Sum(Range(.Cells(Ticker_First_Row, Col_Volume), .Cells(Row_Temp, Col_Volume)))
                    
                    'Set cells appearance
                    .Cells(Row_Update, Col_Chg_Percent).NumberFormat = "0.00%"
                    If Val_Change < 0 Then
                        .Cells(Row_Update, Col_Chg_Value).Interior.ColorIndex = 3
                    Else
                        .Cells(Row_Update, Col_Chg_Value).Interior.ColorIndex = 4
                    End If
                    '===================================================
                    
                    Ticker = .Cells(RowX, Col_Ticker).Value
                    Ticker_First_Row = RowX
                    
                End If
                '##############################################################################################################
                
            Next RowX   ' The loop will exit after the last row of the sheet
            
            '======= Update the Greatest data, using "WorksheetFunction" =======
            .Cells(Row_Start - 1, Col_Grt_Ticker).Value = "Ticker"
            .Cells(Row_Start - 1, Col_Grt_Value).Value = "Value"
            
            .Cells(Row_Start, Col_Grt_X).Value = "Greatest % Increase"
            '***** Get the values from "Yearly Change" columns *****
            Val_Temp = WorksheetFunction.Max(Range(.Cells(Row_Start, Col_Chg_Percent), .Cells(Row_Update, Col_Chg_Percent)))
            'Find the row that the value is, then get the ticker from the row
            Row_Temp = WorksheetFunction.Match(Val_Temp, Range(.Cells(Row_Start, Col_Chg_Percent), .Cells(Row_Update, Col_Chg_Percent)), 0) + 1
            '*******************************************************
            .Cells(Row_Start, Col_Grt_Ticker).Value = .Cells(Row_Temp, Col_Chg_Ticker).Value
            .Cells(Row_Start, Col_Grt_Value).Value = Val_Temp
            .Cells(Row_Start, Col_Grt_Value).NumberFormat = "0.00%"
            
            .Cells(Row_Start + 1, Col_Grt_X).Value = "Greatest % Dedrease"
            Val_Temp = WorksheetFunction.Min(Range(.Cells(Row_Start, Col_Chg_Percent), .Cells(Row_Update, Col_Chg_Percent)))
            Row_Temp = WorksheetFunction.Match(Val_Temp, Range(.Cells(Row_Start, Col_Chg_Percent), .Cells(Row_Update, Col_Chg_Percent)), 0) + 1
            .Cells(Row_Start + 1, Col_Grt_Ticker).Value = .Cells(Row_Temp, Col_Chg_Ticker).Value
            .Cells(Row_Start + 1, Col_Grt_Value).Value = Val_Temp
            .Cells(Row_Start + 1, Col_Grt_Value).NumberFormat = "0.00%"
            
            .Cells(Row_Start + 2, Col_Grt_X).Value = "Greatest Total Volume"
            Val_Temp = WorksheetFunction.Max(Range(.Cells(Row_Start, Col_Chg_Volume), .Cells(Row_Update, Col_Chg_Volume)))
            Row_Temp = WorksheetFunction.Match(Val_Temp, Range(.Cells(Row_Start, Col_Chg_Volume), .Cells(Row_Update, Col_Chg_Volume)), 0) + 1
            .Cells(Row_Start + 2, Col_Grt_Ticker).Value = .Cells(Row_Temp, Col_Chg_Ticker).Value
            .Cells(Row_Start + 2, Col_Grt_Value).Value = Val_Temp
            '====================================================================
            
            .Cells.EntireColumn.AutoFit 'Adjust output appearance
            
        End With
    Next WS

End Sub
