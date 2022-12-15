Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks From (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    
    tickers(1) = "CSIQ"
    
    tickers(2) = "DQ"
    
    tickers(3) = "ENPH"
    
    tickers(4) = "FSLR"
    
    tickers(5) = "HASI"
    
    tickers(6) = "JKS"
    
    tickers(7) = "RUN"
    
    tickers(8) = "SEDG"
    
    tickers(9) = "SPWR"
    
    tickers(10) = "TERP"
    
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    ''1a) Create a ticker Index
        tickerindex = 0

    '1b) Create three output arrays
        Dim tickerVolumes(12) As Long
        
        Dim tickerStartingPrices(12) As Single
        
        Dim tickerEndingPrices(12) As Single
    
    '2a) Create a for loop to initialize the tickerVolumes to zero.
    
    
        For i = 0 To 11
        
        tickerVolumes(i) = 0
    
    Next i
        
        '2b) Loop over all the rows in the spreadsheet.
       
        
    For i = 2 To RowCount
            If Cells(i, 1).Value = tickers(tickerindex) Then
             
        '3a) Increase volume for current ticker
                
            
            tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
    End If
        
            '3b) Check if the current row is the first row with the selected tickerIndex.
            
            If Cells(i - 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerStartingPrices(tickerindex) = Cells(i, 6).Value
    
    End If
           
            
            '3c) check if the current row is the last row with the selected ticker
            'If the next row’s ticker doesn’t match, increase the tickerIndex.
            'If  Then
                
            If Cells(i + 1, 1).Value <> tickers(tickerindex) And Cells(i, 1).Value = tickers(tickerindex) Then
            tickerEndingPrices(tickerindex) = Cells(i, 6).Value
                
                
                '3d) Increase the tickerIndex.
                
                tickerindex = tickerindex + 1
    End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    
         Worksheets("All Stocks Analysis").Activate
         
         For i = 0 To 11
         
         Cells(i + 4, 1).Value = tickers(i)
         
         Cells(i + 4, 2).Value = tickerVolumes(i)
         
         Cells(i + 4, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
         
    Next i
        

    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    'Special Shout out to Jordan Levy as his presentation helped clear mine up. https://github.com/jordanlevy001/challenge-vba-refactoring/blob/main/VBA_Challenge.xlsm

'
    Range("A18").Select
    ActiveCell.FormulaR1C1 = "Average Daily volume"
    Range("A19").Select
    ActiveCell.FormulaR1C1 = "Average Return"
    Range("B18").Select
    ActiveCell.Formula2R1C1 = "=sum"
    Range("B4:B15").Select
    Range("B18").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-14]C:R[-3]C)/12"
    Range("B18").Select
    Selection.AutoFill Destination:=Range("B18:B19"), Type:=xlFillDefault
    Range("B18:B19").Select
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-14]C:R[-3]C)/12"
    Range("B19").Select
    ActiveCell.FormulaR1C1 = "=SUM(R[-15]C[1]:R[-4]C[1])/12"
    Range("A18:A19").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("A18:B19").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B4").Select
    Selection.Copy
    Range("B18").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B18").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B19").Select
    Selection.Style = "Percent"
    Range("F23").Select
    
    
  endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)
  
    
    
    
    
    
End Sub



Sub Erase_Data()
'
' Erase_Data Macro
'

'
    Range("A1:E35").Select
    Selection.ClearContents
    Range("A24:B34").Select
    Selection.Copy
    Range("A1:c20").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Range("e1").Select
End Sub
