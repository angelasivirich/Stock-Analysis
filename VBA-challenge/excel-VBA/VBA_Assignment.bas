
Sub codeRunner()

    'Challenge report variables
    Dim greatestInc As Double
    Dim greatestDec As Double
    Dim greatestVol As Double
    Dim greatestIncName As String
    Dim greatestDecName As String
    Dim greatestVolName As String

    

    Dim wSheet As Worksheet

    '
    ' Interates over all the spreadsheets
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.application.worksheets
    '
    For Each wSheet In ActiveWorkbook.Worksheets
        
        ' Defines the first row of a ticker set
        Dim firstRowRange As Long
        ' Counter to summarize tickers
        Dim columnCreationCounter As Long
        ' Last row of the whole table
        Dim lastRow As Long
        Dim tickerName As String
        
    
        ' Gets the last filled row from the botton to the top
        lastRow = wSheet.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' Writes headers for the summarized results
        Call writeHeaders(wSheet)
        
        ' Sorts data. The sorting order is: <ticker> then <date>
        Call sortData(wSheet, lastRow)
        
        firstRowRange = 2
        columnCreationCounter = 2
        tickerName = wSheet.Cells(2, 1).Value
        For i = 2 To lastRow
            If (wSheet.Cells(i, 1).Value <> tickerName) Then
                
               Call writeColumnCreationValues(wSheet, columnCreationCounter, tickerName, firstRowRange, i - 1, greatestInc, greatestDec, greatestVol, greatestIncName, greatestDecName, greatestVolName)
                
                ' re-assigned values to the variables, to start the new ticker set
                firstRowRange = i
                tickerName = wSheet.Cells(i, 1).Value
                columnCreationCounter = columnCreationCounter + 1
            End If
        Next i
        
        ' Calls again function to calculate the last range
        Call writeColumnCreationValues(wSheet, columnCreationCounter, tickerName, firstRowRange, lastRow, greatestInc, greatestDec, greatestVol, greatestIncName, greatestDecName, greatestVolName)

    Next wSheet
    
    '
    ' writes the updated indexes
    '
    ActiveWorkbook.Worksheets("2016").Activate
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    Range("O2").Value = "Greatest % increase"
    Range("P2").Value = greatestIncName
    Range("Q2").Value = greatestInc
    
    Range("O3").Value = "Greatest % decrease"
    Range("P3").Value = greatestDecName
    Range("Q3").Value = greatestDec
    
    Range("O4").Value = "Greatest total volume"
    Range("P4").Value = greatestVolName
    Range("Q4").Value = greatestVol
    
    MsgBox ("All worksheets were sucessfully processed!")
    
End Sub

Private Sub writeColumnCreationValues(ByRef wSheet As Worksheet, ByVal columnCreationCounter As Long, _
        ByVal tickerName As String, ByVal firstRowRange As Long, ByVal lastRowRange As Long, _
        ByRef greatestInc As Double, ByRef greatestDec As Double, ByRef greatestVol As Double, _
        ByRef greatestIncName As String, ByRef greatestDecName As String, ByRef greatestVolName As String)
        
    Dim openPrice As Double
    Dim closePrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    
    
    ' The first row of the ticker range has the earliest date because the data was sorted before.
    openPrice = wSheet.Cells(firstRowRange, 3).Value
    ' the last row of the ticker range has the latest date because the data was sorted before.
    closePrice = wSheet.Cells(lastRowRange, 6).Value
    yearlyChange = closePrice - openPrice
    
    If yearlyChange <> 0 Then
        percentChange = openPrice / yearlyChange
    Else
        percentChange = 0
    End If
    
    totalVolume = Application.WorksheetFunction.Sum(wSheet.Range("G" & firstRowRange & ":G" & lastRowRange))
    
    wSheet.Cells(columnCreationCounter, 9).Value = tickerName
    wSheet.Cells(columnCreationCounter, 10).Value = yearlyChange
    wSheet.Cells(columnCreationCounter, 11).Value = percentChange
    wSheet.Cells(columnCreationCounter, 12).Value = totalVolume

    ' paint cell if yearly change is negative
    If yearlyChange < 0 Then
        wSheet.Cells(columnCreationCounter, 10).Interior.ColorIndex = 3
    Else
        wSheet.Cells(columnCreationCounter, 10).Interior.ColorIndex = 4
    End If
    
    ' update indexes
    If greatestIncName = "" Then
        ' if none of the variables were initialized then initialize with first row of the creation table.
        greatestInc = percentChange
        greatestDec = percentChange
        greatestVol = totalVolume
        greatestIncName = tickerName
        greatestDecName = tickerName
        greatestVolName = tickerName
    Else
        If percentChange > greatestInc Then
            greatestInc = percentChange
            greatestIncName = tickerName
        End If
        
        If percentChange < greatestDec Then
            greatestDec = percentChange
            greatestDecName = tickerName
        End If
        
        If totalVolume > greatestVol Then
            greatestVol = totalVolume
            greatestVolName = tickerName
        End If
    End If
    

End Sub

Private Sub writeHeaders(ByRef wSheet As Worksheet)
    wSheet.Range("I1").Value = "Ticker"
    wSheet.Range("J1").Value = "Yearly Change"
    wSheet.Range("K1").Value = "Percent Change"
    wSheet.Range("L1").Value = "Total Stock Volume"
End Sub

'
' Sorts column A first and column B
' Headers must be present in the given range
'
' https://docs.microsoft.com/en-us/office/vba/api/excel.range.sort
'
Private Sub sortData(ByRef wSheet As Worksheet, ByVal lastRow As Long)
    wSheet.Activate
    wSheet.Sort.SortFields.Clear
    ' Filter column A order A to Z
    wSheet.Sort.SortFields.Add2 Key:=Range("A2:A" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ' And then filter column B smallest to largest
    wSheet.Sort.SortFields.Add2 Key:=Range("B2:B" & lastRow), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With wSheet.Sort
        .SetRange Range("A1:G" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

End Sub






