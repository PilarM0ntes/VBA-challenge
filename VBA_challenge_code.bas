Attribute VB_Name = "Module1"
Option Explicit

Sub start_program()
    Dim last_row As Long
    Dim i As Long
    Dim j As Long
    Dim old_ticker As String
    Dim volume As Double
    Dim open_value As Double
    Dim close_value As Double
    Dim dif_value As Double
    
    last_row = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Sorts out the information in the worksheet so the tickers & dates are always in ascending order
    With ActiveSheet.Sort
        .SortFields.Add Key:=Range("A1"), Order:=xlAscending
        .SortFields.Add Key:=Range("B1"), Order:=xlAscending
        .SetRange Range(Cells(1, 1), Cells(last_row, 7))
        .Header = xlYes
        .Apply
    End With
    
    'Initializes variables
    j = 2
    old_ticker = Cells(2, 1).Value
    volume = 0
    open_value = Cells(2, 3).Value
    
    'Prints the headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    
    For i = 2 To last_row + 1
        'Verifies if the current ticker is equals the old ticker
        If old_ticker = Cells(i, 1).Value Then
            volume = volume + Cells(i, 7).Value
        ElseIf i = last_row Then
            'Prints the last record
            close_value = Cells(i - 1, 6).Value
            Cells(j, 9).Value = old_ticker
            dif_value = close_value - open_value
            Cells(j, 10).Value = dif_value
            If dif_value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            If open_value = 0 Then
                Cells(j, 11).Value = 0
            Else
                Cells(j, 11).Value = (close_value / open_value) - 1
            End If
            Cells(j, 11).NumberFormat = "0.00%"
            Cells(j, 12).Value = volume
        Else
            'The current ticker changed, print the old ticker information
            close_value = Cells(i - 1, 6).Value
            Cells(j, 9).Value = old_ticker
            dif_value = close_value - open_value
            Cells(j, 10).Value = dif_value
            If dif_value > 0 Then
                Cells(j, 10).Interior.ColorIndex = 4
            Else
                Cells(j, 10).Interior.ColorIndex = 3
            End If
            If open_value = 0 Then
                Cells(j, 11).Value = 0
            Else
                Cells(j, 11).Value = (close_value / open_value) - 1
            End If
            Cells(j, 11).NumberFormat = "0.00%"
            Cells(j, 12).Value = volume
            j = j + 1
            'Initializes the variables for the new ticker
            old_ticker = Cells(i, 1).Value
            open_value = Cells(i, 3).Value
            volume = Cells(i, 7).Value
        End If
        
    Next i


    
    
    
End Sub

