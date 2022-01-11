# MJP_NW_BC_Homework
Sub Stocks():
    For Each ws In Worksheets
        Dim WorksheetName As String
        'Current row
        Dim i As Long
        'Start row of ticker
        Dim j As Long
        'Index counter to fill Ticker row
        Dim TickCount As Long
        'Last row column "A"
        Dim LastRowA As Long
        'Define the last row in column I
        Dim LastRowI As Long
        'Define variable for percent change calculation to hundredths decimal
        Dim PerChange As Double
        'Define variable for greatest total volume
        Dim GreatVol As Double
        
        'Pull WorksheetName
        WorksheetName = ws.Name
        
        'Create headers for new data
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Set Ticker Counter, row 2
        TickCount = 2
        'Ticker begins at row 2
        j = 2
        'Find the last non-blank cell the ticker column
        LastRowA = ws.Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox (Last row in column A is "& LastRowA")
            'Loop through all rows
            For i = 2 To LastRowA
                'Check if ticker name changed
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                'Write ticker in the ninth column (I)
                ws.Cells(TickCount, 9).Value = ws.Cells(i, 1).Value
                'Yearly Change in tenth column (J)
                ws.Cells(TickCount, 10).Value = ws.Cells(i, 6).Value - ws.Cells(j, 3).Value
                    'Conditional formatting
                    If ws.Cells(TickCount, 10).Value < 0 Then
                    'Set background red
                        ws.Cells(TickCount, 10).Interior.ColorIndex = 3
                    Else
                    'Set background green
                        ws.Cells(TickCount, 10).Interior.ColorIndex = 4
                    End If
                    'Calculate and write percent change in column K (#11)
                    If ws.Cells(j, 3).Value <> 0 Then
                        PerChange = ((ws.Cells(i, 6).Value - ws.Cells(j, 3).Value) / ws.Cells(j, 3).Value)
                    'Percent formatting
                        ws.Cells(TickCount, 11).Value = Format(PerChange, "Percent")
                    Else
                        ws.Cells(TickCount, 11).Value = Format(0, "Percent")
                    End If
                'Calculate and write total volume in column L (#12)
                    ws.Cells(TickCount, 12).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
                'Increase TickCount by 1
                    TickCount = TickCount + 1
                'Set new start row of the ticker block
                j = i + 1
                End If
            Next i
    Next ws
End Sub
