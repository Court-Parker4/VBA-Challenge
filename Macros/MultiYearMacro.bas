Attribute VB_Name = "Module1"
Sub VbaChallenge1():
    'Define workbooks
    Dim ws As Worksheet
    Dim Summary_Table As Boolean
   'Dim Last_Row_Count As Long
    

    Summary_Table = False
    
    'Loop through all of the worksheets
    For Each ws In Worksheets
    'ws = ThisWorkbook.Worksheets("A")
        'MsgBox (ws.Name)
        
        'Set initial variables
        Dim Ticker_Symbol As String
        Ticker_Symbol = " "
        Dim Opening_Price As Double
        Opening_Price = 0
        Dim Closing_Price As Double
        Closing_Price = 0
        Dim Yearly_Change As Double
        Yearly_Change = 0
        Dim Percent_Change As Double
        Percent_Change = 0
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        'This is the Summary table
        Dim Track_Ticker_Row As Double
        Track_Ticker_Row = 2
     
        'Dim last_row_count As Double
        Dim i As Long

        'last_row_count = "70926"
        'MsgBox (Cells(Rows.Count, 1).End(xlUp).Row)
        last_row_count = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
       ' .End(x1Up).Row

        If Summary_Table Then
            'Inserting Headers via Cells
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            'Resize Headers
            ws.Range("J1").EntireColumn.ColumnWidth = 15
            ws.Range("K1").EntireColumn.ColumnWidth = 15
            ws.Range("L1").EntireColumn.ColumnWidth = 18
        Else
            'Loop to first worksheet
            Summary_Table = True
        End If
        Opening_Price = ws.Cells(2, 3).Value
    
        'Loop through all tickers
        For i = 2 To last_row_count
        
            'Check if still in the same ticker symbol, IF it is not then add total stock volume
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                'Set the ticker symbol
                Ticker_Symbol = ws.Cells(i, 1).Value
            
                'Set closing Price
                Closing_Price = ws.Cells(i, 6).Value
            
                'Set the Opening Price
                Opening_Price = ws.Cells(i, 3).Value
            
                'Calculate yearly change in stock open and close price
                Yearly_Change = Closing_Price - Opening_Price
            
                If Opening_Price <> 0 Then
                Percent_Change = (Yearly_Change / Opening_Price) * 100
            
                End If
                'Add to the stock volume total
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
                'Print the ticker symbol in the summary table
                ws.Range("I" & Track_Ticker_Row).Value = Ticker_Symbol
            
                'Print the Yearly Change in the stock opening and closing price
                ws.Range("J" & Track_Ticker_Row).Value = Yearly_Change
            
                'Highlight the Positive and Negative Change in Stock Price
                If (Yearly_Change > 0) Then
                    'Postive = green
                    ws.Range("J" & Track_Ticker_Row).Interior.ColorIndex = 4
                ElseIf (Yearly_Change <= 0) Then
                    'Negative = Red
                    ws.Range("J" & Track_Ticker_Row).Interior.ColorIndex = 3
                End If
                'Print the Percent Change in the stock opening and Closing price
                ws.Range("K" & Track_Ticker_Row).Value = (CStr(Percent_Change) & "%")
                'Print the total stock volume to the summary table
                ws.Range("L" & Track_Ticker_Row).Value = Total_Stock_Volume
                'Add one to the summary row
                Track_Ticker_Row = Track_Ticker_Row + 1
                
                'Reset the yearly change
                Yearly_Change = 0
                
                'If the cell immediately following the row is the same ticker
                Else
                    'add to the Total Stock Volume
                    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                    
                End If
                
            Next i
     Next ws
End Sub
