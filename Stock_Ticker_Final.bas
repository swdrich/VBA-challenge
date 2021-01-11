Attribute VB_Name = "Module1"
Sub ticker_loop_ws():

'loop worksheets
' Declare ws as a worksheet object variable.
    Dim ws As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each ws In ActiveWorkbook.Worksheets
    
        ' This line displays the worksheet name in a message box.
        'MsgBox ws.Name
        'Debug.Print ws.Name
        
'--------------------------------------------------------------------

            'Make and label new columns
            ws.Range("I1").Value = "Ticker   "
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
        
            'Set up results table
            Dim SummaryTableRow As Integer
            SummaryTableRow = 2
        
            'Resize Columns
            ws.Columns("I:L").AutoFit
            
'---------------------------------------------------------------------

            'Calculate All Values
            
            Dim LastRow As Long
            LastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
                For i = 2 To (LastRow - 1) '1000
    
                    'Set Variables
                    Dim TickerName As String
                    Dim OpenPrice As Double
                    Dim ClosePrice As Double
                    Dim YearlyChng As Double
                    Dim PercentChgn As Double
                    Dim StockVolTotal As Double
                        StockVolTotal = 0
                    Dim CurrentRow As String
                        CurrentRow = ws.Cells(i, 1).Value
                    Dim PreviousRow As String
                        PreviousRow = ws.Cells(i - 1, 1).Value
                    Dim NextRow As String
                        NextRow = ws.Cells(i + 1, 1).Value
                    
                    
                    'Compare ticker symbols for opening info
                    If CurrentRow <> PreviousRow Then
                    
                        OpenPrice = ws.Cells(i, 3).Value
                        'record opening price - for troubleshooting purposes
                        'ws.Range("N" & SummaryTableRow) = OpenPrice
                        
                    End If
                    
                    'Compare ticker symbols for closing info
                    If CurrentRow <> NextRow Then
                    
                        'Define variables
                        TickerName = ws.Cells(i, 1).Value
                        ClosePrice = ws.Cells(i, 6).Value
                        
                        'record closing price - for troubleshooting purposes
                        'ws.Range("O" & SummaryTableRow) = ClosePrice
                        'Debug.Print (TickerName & ClosePrice)
                        
                        'subtract opening price from closing price
                        YearlyChng = ClosePrice - OpenPrice
                        
                        'calculate percentage change
                        If OpenPrice <> 0 Then
                            PercentChng = (YearlyChng / OpenPrice) * 100
                        Else
                            PercentChng = "N/A"
                        End If
                        
                        'calculate stock volume
                        StockVolTotal = StockVolTotal + ws.Cells(i, 7).Value
                    
                        'Debug.Print (TickerName & ClosePrice)
                        
'-------------------------------------------------------------------
                        
                        'Add Values to Table
                        
                        'Add ticker abbreviation
                        ws.Range("I" & SummaryTableRow).Value = TickerName
                        
                        'Add Yearly Change
                        ws.Range("J" & SummaryTableRow).Value = YearlyChng
                            
                            'Format Yearly Change
                            If YearlyChng >= 0 Then
                                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                            ElseIf YearlyChng < 0 Then
                                ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                            End If
                                
                        'Add Percent Change
                        ws.Range("K" & SummaryTableRow).Value = PercentChng
                        
                            'Format Percent Change
                            ws.Range("K" & SummaryTableRow).NumberFormat = "0.00\%"
                        
                        'Add Stock Volume Total
                        ws.Range("L" & SummaryTableRow).Value = StockVolTotal
                        
                        'Set up next row
                        SummaryTableRow = SummaryTableRow + 1
                        
                        'Reset Stock Total Counter
                        StockVolTotal = 0
                                                                       
                    Else
                    
                        'Add to stock volume total
                        StockVolTotal = StockVolTotal + ws.Cells(i, 7).Value
                    
                    End If
                    
                Next i
'-----------------------------------------------------------------------

        'Bonus Exercise
        'Make and label new table
        ws.Range("P1").Value = "Ticker      "
        ws.Range("Q1").Value = "Value          "
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Resize Columns
        ws.Columns("O:Q").AutoFit
        
            'Find Largest Percent Positive Change
            Dim myRangeMax As Range
                Set myRangeMax = ws.Range("K2:K" & LastRow)
            Dim AnswerMax As Double
                AnswerMax = Application.WorksheetFunction.Max(myRangeMax)
            'MsgBox (AnswerMax)
            
            'Find Largest Percent Negative Change
            Dim myRangeMin As Range
                Set myRangeMin = ws.Range("K2:K" & LastRow)
            Dim AnswerMin As Double
                AnswerMin = Application.WorksheetFunction.Min(myRangeMin)
            'MsgBox (AnswerMin)
            
            'Find Largest Total Volume
            Dim myRangeTotal As Range
                Set myRangeTotal = ws.Range("L2:L" & LastRow)
            Dim AnswerTotal As Double
                AnswerTotal = Application.WorksheetFunction.Max(myRangeTotal)
            'MsgBox (AnswerTotal)
            
            
            'Record Values
            
            'Record Largest Percent Positive Change
            ws.Range("Q2").Value = AnswerMax
                'Format Percent
                ws.Range("Q2").NumberFormat = "0.00\%"
                
            'Record Largest Percent Negative Change
            ws.Range("Q3").Value = AnswerMin
                'Format Percent
                ws.Range("Q3").NumberFormat = "0.00\%"
                
            'Record Largest Total Volume
            ws.Range("Q4").Value = AnswerTotal
            
            
            'Find and Record Ticker Values
            
            'Find Max
            Dim TickerMax As String
            Dim RowMax As Long
            ws.Range("R2").Value = Application.WorksheetFunction.Match(AnswerMax, ws.Range("K1:K" & LastRow), 0)
            RowMax = ws.Range("R2").Value
            TickerMax = ws.Range("I" & RowMax).Value
            'Debug.Print (TickerMax)
            'MsgBox (RowMax)
            ws.Range("P2").Value = TickerMax
            ws.Range("R2").ClearContents
            
            'Find Min
            Dim TickerMin As String
            Dim RowMin As Long
            ws.Range("R3").Value = Application.WorksheetFunction.Match(AnswerMin, ws.Range("K1:K" & LastRow), 0)
            RowMin = ws.Range("R3").Value
            TickerMin = ws.Range("I" & RowMin).Value
            ws.Range("P3").Value = TickerMin
            ws.Range("R3").ClearContents
            
            'Find Total
            Dim TickerTotal As String
            Dim RowTotal As Long
            ws.Range("R4").Value = Application.WorksheetFunction.Match(AnswerTotal, ws.Range("L1:L" & LastRow), 0)
            RowTotal = ws.Range("R4").Value
            TickerTotal = ws.Range("I" & RowTotal).Value
            ws.Range("P4").Value = TickerTotal
            ws.Range("R4").ClearContents
            
            
    Next ws
    
            
End Sub

Sub ClearColumns():

'loop worksheets
' Declare ws as a worksheet object variable.
    Dim ws As Worksheet

    ' Loop through all of the worksheets in the active workbook.
    For Each ws In ActiveWorkbook.Worksheets
    
        'Clear Columns I:O
        ws.Columns(9).ClearContents
        ws.Columns(10).ClearContents
        ws.Columns(10).Interior.ColorIndex = 0
        ws.Columns(11).ClearContents
        ws.Columns(12).ClearContents
        ws.Columns(13).ClearContents
        ws.Columns(14).ClearContents
        ws.Columns(15).ClearContents
    Next ws

End Sub
