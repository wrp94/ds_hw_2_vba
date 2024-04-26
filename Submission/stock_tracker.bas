Attribute VB_Name = "Module1"
Option Explicit
Sub SheetLoop()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Activate
        Call stock_tracker
    Next ws
End Sub

Sub stock_tracker()
    Dim ticker As String
    Dim open_val As Double
    Dim close_val As Double
    Dim vol As Long
    
    Dim total_vol As Double
    Dim val_change As Double
    Dim percent_change As Double
    Dim stock_number As Long: stock_number = 2
    
    Dim largest_increase_val As Double: largest_increase_val = 0
    Dim largest_increase_ticker As String
    Dim largest_decrease_val As Double: largest_decrease_val = 0
    Dim largest_decrease_ticker As String
    Dim largest_volume_val As Double: largest_volume_val = 0
    Dim largest_volume_ticker As String
    
    Dim row As Long: row = 2
    
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Quarterly Change"
    Cells(1, 12).Value = "Percent Change"
    Cells(1, 13).Value = "Total Stock Volume"
    
    
    ticker = Cells(row, 1).Value
    open_val = Cells(row, 3).Value
    total_vol = 0
    
    While Not (Cells(row, 1) = "")
        total_vol = total_vol + Cells(row, 7).Value
    
        If (Cells(row, 1).Value <> Cells(row + 1, 1)) Then
        
            close_val = Cells(row, 6).Value
            val_change = close_val - open_val
            percent_change = val_change / open_val
            Cells(stock_number, 10).Value = ticker
            Cells(stock_number, 11).Value = val_change
            Cells(stock_number, 12).Value = percent_change
            Cells(stock_number, 13).Value = total_vol
            
            If (val_change < 0) Then
                Cells(stock_number, 11).Interior.ColorIndex = 3
            ElseIf (val_change > 0) Then
                Cells(stock_number, 11).Interior.ColorIndex = 4
            End If
            
            If (total_vol > largest_volume_val) Then
                largest_volume_val = total_vol
                largest_volume_ticker = ticker
            End If
                
            If (percent_change > largest_increase_val) Then
                largest_increase_val = percent_change
                largest_increase_ticker = ticker
            ElseIf (percent_change < largest_decrease_val) Then
                largest_decrease_val = percent_change
                largest_decrease_ticker = ticker
            End If
        
            
            stock_number = stock_number + 1
            open_val = Cells(row + 1, 3).Value
            total_vol = 0
            ticker = Cells(row + 1, 1).Value
        End If
        
        row = row + 1
    Wend
        
    
    
    Cells(2, 15).Value = "Greatest % Increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest Total Volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    Cells(2, 16).Value = largest_increase_ticker
    Cells(3, 16).Value = largest_decrease_ticker
    Cells(4, 16).Value = largest_volume_ticker
    Cells(2, 17).Value = largest_increase_val
    Cells(3, 17).Value = largest_decrease_val
    Cells(4, 17).Value = largest_volume_val
    
    With Range("L:L, Q2:Q3")
        .NumberFormat = "0.00%"
    End With
    
    With Range("K:K")
        .NumberFormat = "0.00"
    End With
        
End Sub

