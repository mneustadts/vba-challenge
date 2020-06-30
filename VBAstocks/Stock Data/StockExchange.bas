Attribute VB_Name = "Module1"
Sub StockExchange()

Dim ws As Worksheet
Dim ticker As String
Dim vol As Double
Dim year_open As Double
Dim year_close As Double
Dim yearly_change As Double
Dim pct_change As Double
Dim Summary_Table_Row As Double
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

On Error Resume Next

For Each ws In ThisWorkbook.Worksheets

    ws.Activate
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    
    Summary_Table_Row = 2
    
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow
             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            ticker = ws.Cells(i, 1).Value
            vol = ws.Cells(i, 7).Value

            year_open = ws.Cells(i, 3).Value
            year_close = ws.Cells(i, 6).Value

            yearly_change = year_close - year_open
            pct_change = (year_close - year_open) / year_close
        

            ws.Cells(Summary_Table_Row, 9).Value = ticker
            ws.Cells(Summary_Table_Row, 10).Value = yearly_change
            ws.Cells(Summary_Table_Row, 11).Value = pct_change
            ws.Cells(Summary_Table_Row, 12).Value = vol
            
            Summary_Table_Row = Summary_Table_Row + 1
            
            
            vol = 0
        
        End If


    Next i
    
ws.Columns("K").NumberFormat = "0.00%"

    Dim cell_color As Range
    Dim yg As Range
    Dim m As Long
    Dim n As Long
    
    Set yg = ws.Range("J2", Range("J2").End(xlDown))
    n = yg.Cells.Count
    
    For m = 1 To n
    Set cell_color = yg(m)
    Select Case cell_color
        Case Is >= 0
            With cell_color
                .Interior.Color = vbGreen
            End With
        Case Is < 0
            With cell_color
                .Interior.Color = vbRed
            End With
       End Select
    Next m


Next ws

starting_ws.Activate

End Sub


