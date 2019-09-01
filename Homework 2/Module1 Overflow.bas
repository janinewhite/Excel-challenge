Attribute VB_Name = "Module1"
Option Explicit
'Define columns
Public Const OriginalTickerCol = 1
Public Const DateCol = 2
Public Const OpenCol = 3
Public Const CloseCol = 6
Public Const VolumeCol = 7
Public Const DistinctTickerCol As Integer = 9
Public Const YearlyChangeCol As Integer = 10
Public Const PercentChangeCol As Integer = 11
Public Const TotalVolumeCol As Integer = 12
Public Const YearOpenCol As Integer = 13
Public Const YearCloseCol As Integer = 14
Public Const GreatestLabelCol As Integer = 15
Public Const GreatestTickerCol As Integer = 16
Public Const GreatestVariableCol As Integer = 17
Public Const HeaderRow As Integer = 1
Public Const GreatestIncreaseRow As Integer = 2
Public Const GreatestDecreaseRow As Integer = 3
Public Const GreatestVolumeRow As Integer = 4



Sub AnalyzeStocks()

    Dim SheetNum, Stocksheets, Tickers, OriginalRowNum, DistinctRowNum, SearchRow, FoundRow, TickerRow, RowNum, StartTickerRow, EndTickerRow As Integer
    Dim ws As Worksheet
    Dim SearchRng, FoundRng As Range
    Dim SheetName, StockTicker, SearchTicker, SearchCells, SearchEnd, GreatestIncreaseTicker, GreatestDecreaseTicker, GreatestVolumeTicker As String
    Dim FoundTicker, FoundRange, FirstRow, LastRow As Boolean
    Dim YearOpen, YearClose, YearChange, YearPercentChange, PercentChange, GreatestIncrease, GreatestDecrease As Double
    Dim TickerDays, Volume, TotalVolume, GreatestVolume As LongLong

      
    'Loop through Worksheets
    Stocksheets = ActiveWorkbook.Worksheets.Count
    For SheetNum = 1 To Stocksheets
        Set ws = ActiveWorkbook.Worksheets(SheetNum)
        With ws
        
            SheetName = .Name
                               
            'Write required headers
            .Cells(HeaderRow, DistinctTickerCol).Value = "Ticker"
            .Cells(HeaderRow, YearlyChangeCol).Value = "Yearly Change"
            .Cells(HeaderRow, PercentChangeCol).Value = "Percent Change"
            .Cells(HeaderRow, TotalVolumeCol).Value = "Total Stock Volume"
            .Cells(HeaderRow, YearOpenCol).Value = "Year Open"
            .Cells(HeaderRow, YearCloseCol).Value = "Year Close"
            .Cells(HeaderRow + 1, GreatestLabelCol).Value = "Greatest % Increase"
            .Cells(HeaderRow + 2, GreatestLabelCol).Value = "Greatest % Decrease"
            .Cells(HeaderRow + 3, GreatestLabelCol).Value = "Greatest Total Volume"
            .Cells(HeaderRow, GreatestTickerCol).Value = "Ticker"
            .Cells(HeaderRow, GreatestVariableCol).Value = "Value"

            
            'Write a distinct list of all tickers in the original ticker column in the distinct ticker column
            
            'Initiailize loop variables
            OriginalRowNum = 2
            DistinctRowNum = 2
            FoundTicker = False
            StockTicker = .Cells(OriginalRowNum, 1).Value
            'Continue until all transactions have been read
            Do While Len(StockTicker) > 0
                           
                'Look for ticker in distinct list if any have been added to the list
                If DistinctRowNum > 2 Then
                    SearchCells = "I1:I" & Trim(Str(DistinctRowNum - 1))
                    Set SearchRng = .Range(SearchCells)
                    With SearchRng
                        Set FoundRng = .Find(StockTicker)
                        FoundTicker = Not (FoundRng Is Nothing)
                    End With
                End If
                'If ticker is not in list, add ticker to the list and increment last ticker row
                If Not FoundTicker Then
                    .Cells(DistinctRowNum, DistinctTickerCol).Value = StockTicker
                    .Range("H1").Value = StockTicker + " " + SheetName
                    DistinctRowNum = DistinctRowNum + 1
                End If
                'Prepare to find next ticker
                OriginalRowNum = OriginalRowNum + 1
                StockTicker = .Cells(OriginalRowNum, OriginalTickerCol).Value
            Loop
            'Set the total number of tckers found on the sheet
            Tickers = (DistinctRowNum - 2)
            'Set the total number of transactions found on the sheet
            TickerDays = OriginalRowNum - 2
        
            'Populate annual statistics
            
            'Loop through transactions using distinct tickers
            For TickerRow = 2 To (Tickers + 1)
                StockTicker = .Cells(TickerRow, DistinctTickerCol).Value
                .Range("H1").Value = StockTicker + " " + SheetName
                'Find First Ticker Row
                SearchCells = "A1:A" & Trim(Str(TickerDays + 1))
                Set SearchRng = .Range(SearchCells)
                With SearchRng
                    Set FoundRng = .Find(StockTicker)
                    StartTickerRow = FoundRng.Cells.Row
                End With
                'Initialize Volume and transaction row
                TotalVolume = 0
                RowNum = StartTickerRow
                
                'Loop through transactions for each ticker
                Do While (RowNum <= (TickerDays + 1)) And Not LastRow
                    'First row has a header or different ticker before it
                    FirstRow = (.Cells(RowNum, OriginalTickerCol).Value <> .Cells(RowNum - 1, OriginalTickerCol).Value)
                    'Last row has a header or different ticker after it
                    LastRow = (.Cells(RowNum, OriginalTickerCol).Value <> .Cells(RowNum + 1, OriginalTickerCol).Value)
                    'Get transaction volume
                    Volume = .Cells(RowNum, VolumeCol).Value
                    'Aggregate transaction volume
                    TotalVolume = TotalVolume + Volume
                    'Year Open is in first row
                    If FirstRow Then
                         YearOpen = .Cells(RowNum, OpenCol).Value
                         .Cells(TickerRow, YearOpenCol).Value = YearOpen
                    End If
                    'Year close is in last row, also calculate annual change and % change, and finish iterating current ticker.
                    If LastRow Then
                        EndTickerRow = RowNum
                        YearClose = .Cells(RowNum, CloseCol).Value
                        .Cells(TickerRow, YearCloseCol).Value = YearClose
                        .Cells(TickerRow, TotalVolumeCol).Value = TotalVolume
                        YearChange = YearClose - YearOpen
                        .Cells(TickerRow, YearlyChangeCol).Value = YearChange
                        YearPercentChange = YearChange / YearOpen
                        .Cells(TickerRow, PercentChangeCol).Value = YearPercentChange
                        Exit Do
                    End If
                    'Increment transaction row
                    RowNum = RowNum + 1
                Loop
                LastRow = False
             Next TickerRow

            MsgBox ("Aggregation Complete " + SheetName)
            
            'Find greatest chanages
            
            'Initialize loop variables
            TickerRow = 2
            GreatestIncrease = 0
            GreatestDecrease = 0
            GreatestVolume = 0
            StockTicker = .Cells(TickerRow, DistinctTickerCol).Value
            'Loop through distinct tickers
            Do While Len(StockTicker) > 0
                .Range("H1").Value = StockTicker + " " + SheetName
                'Get annual % change and total volume for ticker
                PercentChange = .Cells(TickerRow, PercentChangeCol).Value
                TotalVolume = .Cells(TickerRow, TotalVolumeCol).Value
                'Analyze % Change
                If PercentChange > GreatestIncrease Then
                    .Cells(GreatestIncreaseRow, GreatestTickerCol).Value = StockTicker
                    .Cells(GreatestIncreaseRow, GreatestVariableCol).Value = PercentChange
                    GreatestIncrease = PercentChange
                Else
                    If PercentChange < GreatestDecrease Then
                        .Cells(GreatestDecreaseRow, GreatestTickerCol).Value = StockTicker
                        .Cells(GreatestDecreaseRow, GreatestVariableCol).Value = PercentChange
                        GreatestDecrease = PercentChange
                    End If
                End If
                'Analyze Volume
                If TotalVolume > GreatestVolume Then
                    .Cells(GreatestVolumeRow, GreatestTickerCol).Value = StockTicker
                    .Cells(GreatestVolumeRow, GreatestVariableCol).Value = TotalVolume
                End If
                'Increment Loop
                TickerRow = TickerRow + 1
                StockTicker = .Cells(TickerRow, DistinctTickerCol).Value
            Loop
            
            'Clear debug displays
            .Range("H1").Value = ""
            .Range("M:M").Value = ""
            .Range("N:N").Value = ""
            .Range("S:S").Value = ""
            .Range("T:T").Value = ""
            
            MsgBox ("Greatest changes " + SheetName + " found.")
     
        End With
    Next SheetNum
    
    MsgBox ("Analysis Complete")
    

End Sub
