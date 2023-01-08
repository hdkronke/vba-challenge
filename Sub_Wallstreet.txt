Sub WallStreet()

'Run Process on all worksheets
Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets
ws.Activate

'---------------------------------------------------------

    'This allegedly speeds up macro processing
    Application.ScreenUpdating = False

    'Set the variables
    Dim Ticker As String
    Dim Ticker_Volume As Double
        Ticker_Volume = 0
    Dim Ticker_Open As Double
        Ticker_Open = 0
    Dim Ticker_Close As Double
        Ticker_Close = 0
    Dim Yearly_Change As Double
        Yearly_Change = 0
    Dim Percent_Change As Double
        Percent_Change = 0
  
    'Set Variable for Last Row
    Dim LastRow As Long
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row

    'Keep track of the location for each Ticker Name in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Loop through all line items
    For i = 2 To LastRow

        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then

            'Set Ticker Open - this has to be separate from the rest of the loop or the info gets overwritten by the next data
            Ticker_Open = Cells(i, 3).Value
            
            'Set Ticker Volume - this has to be separate from the rest of the loop because it started acting weird until i did this
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

        'Check if value is within the same Ticker Name
        ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            
            'Set the Ticker Name
            Ticker = Cells(i, 1).Value

            'Set Ticker Close
            Ticker_Close = Ticker_Close + Cells(i, 6).Value
      
            'Set the Yearly Change
            Yearly_Change = Ticker_Close - Ticker_Open
      
            'Set the Percent Change
            Percent_Change = (((Ticker_Close - Ticker_Open) / Ticker_Open))

            'Add to the Ticker Volume
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

            'Print the Summary Table
            Range("I" & Summary_Table_Row).Value = Ticker
            Range("J" & Summary_Table_Row).Value = Yearly_Change
            Range("K" & Summary_Table_Row).Value = Percent_Change
            Range("L" & Summary_Table_Row).Value = Ticker_Volume

            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            'Reset the totals
            Ticker_Close = 0
            Ticker_Volume = 0

        Else

            'Add to the Ticker Volume
            Ticker_Volume = Ticker_Volume + Cells(i, 7).Value

        End If

    Next i

'---------------------------------------------------------

    'Format positive and negative yearly changes
    Dim YearRange As Range
    Set YearRange = Range("J:K")
    
    For Each Cell In YearRange
        If Cell.Value < 0 Then
            Cell.Interior.ColorIndex = 3
        ElseIf Cell.Value > 0 Then
            Cell.Interior.ColorIndex = 4
        End If
    Next

    'Format percent change as percentage
    Range("K:K").NumberFormat = "0.00%"
    
'---------------------------------------------------------

    'Create the other table with the highest and lowest
    Dim Per_Inc As Double
    Dim Per_Dec As Double
    Dim Tot_Vol As Double
    
    'Set range to determine searches
    Search_Percent = Range("k:k")
    Search_Volume = Range("L:L")

    'Use min max to find the values
    Per_Inc = Application.WorksheetFunction.Max(Search_Percent)
        Cells(2, 16).Value = Per_Inc
        
    Per_Dec = Application.WorksheetFunction.Min(Search_Percent)
        Cells(3, 16).Value = Per_Dec
        
    Tot_Vol = Application.WorksheetFunction.Max(Search_Volume)
        Cells(4, 16).Value = Tot_Vol

    'Use match function to find the ticker names
    Dim Tic_Inc As Double
    Tic_Inc = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
        Range("O2") = Cells(Tic_Inc + 1, 9)

    Dim Tic_Dec As Double
    Tic_Dec = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & LastRow)), Range("K2:K" & LastRow), 0)
        Range("O3") = Cells(Tic_Dec + 1, 9)

    Dim Tic_Tot As Double
    Tic_Tot = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & LastRow)), Range("L2:L" & LastRow), 0)
        Range("O4") = Cells(Tic_Tot + 1, 9)
    
    'Format Percents
    Range("P2:P3").NumberFormat = "0.00%"
    
'---------------------------------------------------------
    
    'Close the thing that speeds up macro processing
    Application.ScreenUpdating = True
    
'Loop process on all worksheets
Next ws
    
End Sub
