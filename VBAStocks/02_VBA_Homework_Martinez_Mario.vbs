' CREATED BY Mario Martinez
' 31-03-2020
'----------------------------------------------------------------------------------------------------------
' RUN THIS  SUB FROM EXCEL WORKSHEET
' Below sub will provide the user with a some options to customize the I/O
'----------------------------------------------------------------------------------------------------------
Public Sub runScriptCalculateStocks()

    
    Dim wkSheetName As String, srcSheetName As String, outCIndex As String, outRIndex As String, inCIndex As String, inRIndex As String
    Dim wkSheet  As Worksheet, scrSheet As Worksheet
                
    srcSheetName = InputBox("Please type a worksheet name with source data", "Where is the data?", "2014")
    inRIndex = InputBox("Change this value if data do not start at row 2", "1st row of data", "2")
    inCIndex = InputBox("Change this value if data do not start at column 1", "1st column of data", "1")
    wkSheetName = InputBox("Please type the name of a  sheet to generate the report, if the sheet already exists the report will be generate on that sheet.", "Choose a name", "Report")
    outCIndex = InputBox("Please type the first column index of the result (for example if you type 1 the 1st column of the result will appear on A, if you type 2 on B and so on...) BE CAREFUL TO NOT OVERWRITE SOURCE DATA", "Column for output", 1)
    outRIndex = InputBox("Please type the first row index of the result (for example if you type 1 the headers wii be written on the 1st excel row)", "Row for output", 1)
    
    
    Set scrSheet = lookForSheet(srcSheetName)
    
    If Not scrSheet Is Nothing And validateNumeric(inCIndex & "," & inRIndex & "," & outCIndex & "," & outRIndex, ",") = True Then
    
        Set wkSheet = getSheet(wkSheetName)
        wkSheet.Activate
        
        stockCalculation scrSheet, wkSheet, CLng(inRIndex), CLng(inCIndex), CLng(outRIndex), CLng(outCIndex)
        
    Else
        MsgBox "The worksheet with source data do not exists, or you typed a wrong value (for example a letter or a sign) on a index row/column field. Please review the introduced parameters and try again"
    
    End If
End Sub
'----------------------------------------------------------------------------------------------------------
' RUN THIS SUB FROM EXCEL WORKSHEET
' This will generate the result on every page
'----------------------------------------------------------------------------------------------------------
Public Sub runScriptCalculateStocksEachPage()
     Dim xlSheet As Worksheet
    
    For Each xlSheet In Worksheets
        If xlSheet.Range("A1").Value = "<ticker>" Then
            stockCalculation xlSheet, xlSheet, 2, 1, 1, 10
        End If
    Next xlSheet


End Sub


'------------------------------------------------------------------------------------------------------------
' THIS CALCULATES THE STOCK PRICE
' Note: For the below sub to work as intented data should be order by ticker and date (from smaller date to longer)
'       and columns should follow this order: ticker, date, open, high, low, close, vol
'---------------------------------------------------------------------------------------------------------------
Private Sub stockCalculation(inWorkSheet As Worksheet, outWorkSheet As Worksheet, inRow As Long, inCol As Long, outRow As Long, outCol As Long)

    Dim currentTicker As String, nextTicker As String, grIncTicker As String, greDecTicker As String, greTotVolTicker As String
    Dim x As Long, upperDataBound As Long, lowerTickerBound As Long, upperTickerbound As Long, writeRow As Long, color As Long
    Dim openVal As Double, closeVal As Double, rateOfChange As Double, openCloseDiff As Double, gIncAm As Double, gDecAm As Double, gTotVol As Double, totalVol As Double
    Dim selRange As Range, tRange As Range
    
    With inWorkSheet
        'LOOK FOR THE LAST CELL IN EXCEL
        upperDataBound = .Cells(.Rows.Count, inCol).End(xlUp).Row
        'SET RANGE BASED ON LOWER/UPPER BOUNDS, USING +6 because the width of the data is 7 columns
        Set selRange = .Range(.Cells(inRow, inCol), .Cells(upperDataBound, inCol + 6))
    End With
        'RESIZING UPPER BOUNDS ACCORDING TO RANGE
        upperDataBound = selRange.Rows.Count
        'INITIALIZE THE LOWERTICKERBOUND (1ST DATA ROW IN RANGE)
        lowerTickerBound = 1
        
        'ADDING HEADERS TO OUTPUT FIELDS
        Set tRange = outWorkSheet.Cells(outRow, outCol)
        
        writeOnRow tRange, "Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume"

        'ADDING EXTRA ROWS TO WRITETABLE RANGE, +1 DUE TO THE HEADER THAT WE JUST WROTE
        writeRow = outRow + 1
        

    For x = 1 To upperDataBound
       
        currentTicker = selRange.Cells(x, 1).Value
        nextTicker = selRange.Cells(x, 1).Offset(1).Value
        
        If currentTicker <> nextTicker Then
            upperTickerbound = x
            
            openVal = selRange.Cells(lowerTickerBound, 3).Value
            closeVal = selRange.Cells(upperTickerbound, 6).Value
            openCloseDiff = closeVal - openVal
            rateOfChange = calculateRateOfChange(openCloseDiff, openVal)
            totalVol = Application.Sum(Range(selRange.Cells(lowerTickerBound, 7), selRange.Cells(upperTickerbound, 7)))
                      
            compareTicker grIncTicker, gIncAm, currentTicker, rateOfChange, True
            compareTicker greDecTicker, gDecAm, currentTicker, rateOfChange, False
            compareTicker greTotVolTicker, gTotVol, currentTicker, totalVol, True
         
            Set tRange = outWorkSheet.Cells(writeRow, outCol)

            
            writeOnRow tRange, currentTicker, openCloseDiff, rateOfChange, totalVol
            changeCellColor openCloseDiff, tRange.Offset(0, 1)
            
            writeRow = writeRow + 1
            lowerTickerBound = upperTickerbound + 1
        End If
      
    
    Next x
    
    'WRITE BONUS ROWS, IF TRANGE IS NULL IT MEANS  THAT WE DIDN'T HAVE RECORDS TO COMPARE
        Set tRange = outWorkSheet.Cells(outRow, outCol)
        With tRange
        .Offset(, 7).Value = "Ticker"
        .Offset(, 8).Value = "Value"
        
        .Offset(1, 6).Value = "Greatest % increase"
        .Offset(1, 7).Value = grIncTicker
        .Offset(1, 8).Value = gIncAm
        
        .Offset(2, 6).Value = "Greatest % decrease"
        .Offset(2, 7).Value = greDecTicker
        .Offset(2, 8).Value = gDecAm

        .Offset(3, 6).Value = "Greatest total volume"
        .Offset(3, 7).Value = greTotVolTicker
        .Offset(3, 8).Value = gTotVol
        
        .Offset(1, 6).Columns(1).AutoFit
        
        .Offset(1, 8).NumberFormat = "0.00%"
        .Offset(2, 8).NumberFormat = "0.00%"
        .Offset(3, 8).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
       End With
        
       'APPLY LAST FORMATS ON SOME CELLS
       With outWorkSheet
            .Columns(outCol + 2).NumberFormat = "0.00%"
            .Columns(outCol + 3).NumberFormat = "#,##0.00_);[Red](#,##0.00)"
            .Columns(outCol + 2).AutoFit
            .Columns(outCol + 3).AutoFit
        End With
End Sub


'COMPARE 2 TICKERS AND THEIR QUANTITIES USE TRUE TO KEEP THE HIGHEST VALUE AND FALSE TO KEEP THE LOWER VALUE
Private Sub compareTicker(ByRef actTicker As String, ByRef actVal As Double, compTicker As String, compVal As Double, method As Boolean)
    If method Then
        If actVal < compVal Then
            actTicker = compTicker
            actVal = compVal
        End If
    Else
        If actVal > compVal Then
            actTicker = compTicker
            actVal = compVal
        End If
    End If
End Sub


'WRITE A ROW OF DATA IN SELECTED RANGE
Private Sub writeOnRow(rng As Range, ticker As Variant, ychange As Variant, pchange As Variant, tStock As Variant)
    With rng
        .Value = ticker
        .Offset(0, 1).Value = ychange
        .Offset(0, 2).Value = pchange
        .Offset(0, 3).Value = tStock
    End With
End Sub


'APPLY DIFFERENT COLORS DEPENDING ON VALUE
Private Sub changeCellColor(val As Double, rng As Range)
    If val < 0 Then
        rng.Interior.ColorIndex = 3
    ElseIf val > 0 Then
        rng.Interior.ColorIndex = 10
    End If
End Sub

'FUNCTION TO AVOID DIVIDE BY ZERO ERROR WHILE CALCULATING RATE OF CHANGE
Private Function calculateRateOfChange(Dif As Double, InitVal As Double) As Double
    If InitVal <> 0 Then
        calculateRateOfChange = Dif / InitVal
    Else
        calculateRateOfChange = 0
    End If
End Function


'CREATES A NEW WORKSHEET IF IT DOESN'T EXISTS
Private Function getSheet(name As String) As Worksheet
    Dim objSheet As Worksheet
    
    Set objSheet = lookForSheet(name)
    
    If objSheet Is Nothing Then
        Set objSheet = Sheets.Add(After:=Sheets(Sheets.Count))
        objSheet.name = name
    End If
    Set getSheet = objSheet
    
End Function

'LOOKS FOR A WORKSHEET BY NAME ON THE WORKSHEET COLLECTION OF CURRENT WORKBOOK
Private Function lookForSheet(name As String) As Worksheet
     Dim xlSheet As Worksheet
    
    For Each xlSheet In Worksheets
        If xlSheet.name = name Then
            Set lookForSheet = xlSheet
        End If
    Next xlSheet
    
End Function

'VALIDATE THAT ALL  ELEMENTS ON A DELIMITED STRING ARE NUMERIC
Private Function validateNumeric(delimitedList As String, delimiter As String) As Boolean
    Dim strArray() As String
    Dim x As Integer
    Dim result As Boolean
    
    strArray = Split(delimitedList, delimiter)
    result = True
    
    For x = LBound(strArray) To UBound(strArray)
        If IsNumeric(strArray(x)) = False Then
            result = False
            Exit For
        End If
    Next x
    
    validateNumeric = result
    
End Function


