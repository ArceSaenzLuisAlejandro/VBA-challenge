Attribute VB_Name = "Módulo1"
Sub VBAChallenge()

Dim ws As Worksheet
Dim ticker As String
Dim LastRow As Double
Dim TotalStock As Double
Dim OpeningPrice As Double
Dim ClosingPrice As Double
Dim Summary_Table_Row As Double
Dim flag_opening As Double
Dim flag_closing As Double

 
For Each ws In Worksheets
    
    'Definición de columnas de resultados
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "YearlyChange"
    ws.Range("K1").Value = "Percent Char"
    ws.Range("L1").Value = "Total Stock Volume"
    
    'Obtener rango de filas y nombre de hoja
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    flag_opening = ws.Cells(2, "C").Value
    flag_closing = 0
    Summary_Table_Row = 2
    
    For i = 2 To LastRow:
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            flag_closing = ws.Cells(i, "F").Value
            ticker = ws.Cells(i, 1).Value
            TotalStock = TotalStock + ws.Cells(i, "G").Value
            ws.Range("I" & Summary_Table_Row).Value = ticker
            ws.Range("L" & Summary_Table_Row).Value = TotalStock
            ws.Range("J" & Summary_Table_Row).Value = flag_closing - flag_opening
            If flag_opening <> 0 Then
                ws.Range("K" & Summary_Table_Row).Value = (flag_closing / flag_opening) - 1
            Else
                ws.Range("K" & Summary_Table_Row).Value = 0
            End If
            'Formato color
            If ws.Cells(Summary_Table_Row, "J").Value <= 0 Then
                ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
            Else
                ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
            End If
            
            Summary_Table_Row = Summary_Table_Row + 1
            TotalStock = 0
            flag_opening = ws.Cells(i + 1, "C").Value
               
        Else
            TotalStock = TotalStock + ws.Cells(i, "G").Value
            
        End If
    Next i
    
    ws.Range("K:K").NumberFormat = "0.00%"
    
    '------------Challenges--------------
    Dim ticker_increase As Double
    Dim ticker_decrease As Double
    Dim ticker_volume As Double
    Dim s_ticker_increase As String
    Dim s_ticker_decrease As String
    Dim s_ticker_volume As String
    Dim LastRow_challenge As Double
    
    LastRow_challenge = ws.Cells(Rows.Count, "I").End(xlUp).Row
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    ticker_increase = ws.Cells(2, "K").Value
    ticker_decrease = ws.Cells(2, "K").Value
    ticker_volume = ws.Cells(2, "L").Value
    
    For j = 2 To LastRow_challenge:
        If ws.Cells(j, "K").Value < ticker_decrease Then
            ticker_decrease = ws.Cells(j, "K").Value
            sticker_decrease = ws.Cells(j, "I").Value
        ElseIf ws.Cells(j, "K").Value > ticker_increase Then
            ticker_increase = ws.Cells(j, "K").Value
            sticker_increase = ws.Cells(j, "I").Value
        End If
        
        If ws.Cells(j, "L").Value > ticker_volume Then
            ticker_volume = ws.Cells(j, "L").Value
            sticker_volume = ws.Cells(j, "I").Value
        End If
    
    Next j
    
    ws.Range("P2").Value = sticker_increase
    ws.Range("P3").Value = sticker_decrease
    ws.Range("P4").Value = sticker_volume
    ws.Range("Q2").Value = ticker_increase
    ws.Range("Q3").Value = ticker_decrease
    ws.Range("Q4").Value = ticker_volume
    ws.Range("Q2:Q3").NumberFormat = "0.00%"

Next


End Sub
