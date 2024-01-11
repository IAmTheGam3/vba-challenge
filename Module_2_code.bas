Attribute VB_Name = "Module1"
Sub Multiple_year()
        Dim NewTicker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim TotalVolume As Double
        Dim LastRow As Long
        Dim RowSummary As Double
        Dim ClosePrice As Double
        Dim OpenPrice As Double
        Dim GreatestInc As Double
        Dim GreatestDec As Double
        Dim GreatestVol As Double

        For Each ws In Worksheets
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"

            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            RowSummary = 2
            TotalVolume = 0
            OpenPrice = ws.Cells(2, 3)

            For i = 2 To LastRow
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                    NewTicker = ws.Cells(i, 1).Value

                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value()

                    ClosePrice = ws.Cells(i, 6).Value

                    YearlyChange = (ClosePrice - OpenPrice)

                        If (OpenPrice = 0) Then
                            PercentChange = 0

                        Else
                            PercentChange = YearlyChange / OpenPrice

                        End If

                    ws.Cells(RowSummary, 9).Value = NewTicker

                    ws.Cells(RowSummary, 10).Value = YearlyChange

                    ws.Cells(RowSummary, 11).Value = FormatPercent(PercentChange, 2)

                    ws.Cells(RowSummary, 12).Value = TotalVolume

                    RowSummary = RowSummary + 1

                    TotalVolume = 0

                    OpenPrice = ws.Cells(i + 1, 3).Value

                Else
                    TotalVolume = TotalVolume + ws.Cells(i, 7).Value()

                End If
            Next i

            LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            GreatestInc = 0
            GreatestDec = 0
            GreatestVol = 0

            For j = 2 To LastRow
                If ws.Cells(j, 11) > GreatestInc Then

                    NewTicker = ws.Cells(j, 9).Value

                    GreatestInc = ws.Cells(j, 11).Value
                End If
            Next j

            ws.Cells(2, 16).Value = NewTicker
            ws.Cells(2, 17).Value = FormatPercent(GreatestInc, 2)

            For k = 2 To LastRow
                If ws.Cells(k, 11) < GreatestDec Then

                    NewTicker = ws.Cells(k, 9).Value

                    GreatestDec = ws.Cells(k, 11).Value
                End If
            Next k

            ws.Cells(3, 16).Value = NewTicker

            ws.Cells(3, 17).Value = FormatPercent(GreatestDec, 2)

            For l = 2 To LastRow
                If ws.Cells(l, 12) > GreatestVol Then

                    NewTicker = ws.Cells(l, 9).Value

                    GreatestVol = ws.Cells(l, 12).Value
                End If
            Next l

            ws.Cells(4, 16).Value = NewTicker

            ws.Cells(4, 17).Value = GreatestVol

            For m = 2 To LastRow
                If ws.Cells(m, 10) < 0 Then
                    ws.Cells(m, 10).Interior.ColorIndex = 3
                Else
                    ws.Cells(m, 10).Interior.ColorIndex = 10
                End If
            Next m


        Next ws

End Sub
