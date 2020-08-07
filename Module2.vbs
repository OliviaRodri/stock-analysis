Sub StaticFormatting()
    'Visual Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.Bold = True
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    'Numeric Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "$#,##0.00"
    Range("C4:C15").NumberFormat = "0.00%"
    Columns("B").Autofit
    Dim dataRowStart, dateRowEnd As Integer
    dataRowStart = 4
    dataRowEnd = 15
    For i = dataRowStart To dataRowEnd

        If Cells(i, 3) > 0 Then

            'Color the cell green
            Cells(i, 3).Interior.Color = vbGreen

        ElseIf Cells(i, 3) < 0 Then

            'Color the cell red
            Cells(i, 3).Interior.Color = vbRed

        Else

            'Clear the cell color
            Cells(i, 3).Interior.Color = xlNone

        End If

    Next i    

End Sub    

Sub CheckersSkillDrill()
    '8x8 checkers pattern
    Dim clmniseven, Rowiseven As Double
    For Row = 1 To 8
        'a line of code here will run 8 times

        For clmn = 1 To 8
            clmniseven = clmn Mod 2
            Rowiseven = Row Mod 2
            If clmniseven = 0 And Rowiseven = 0 Then
                Cells(Row, clmn).Interior.Color = vbBlack
            
            ElseIf clmniseven = 1 And Rowiseven = 1 Then
                Cells(Row, clmn).Interior.Color = vbBlack
            Else
                Cells(Row, clmn).Interior.Color = vbRed
       
            End If

        Next clmn
    Next Row
         
End Sub






Sub AllStocksAnalysis()
    '1. Format the output sheet on the "All Stocks Analysis" worksheet
    Worksheets("All Stocks Analysis").Activate
    Range("A1").Value = "All Stocks (2018)"
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    '2. Declare and Initialize an array of all tickers
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
    '3. Prepare for the analysis of tickers
    'Initialize variables for the starting price and ending price.
    Dim totalvolume, rowStart, rowEnd As Integer
    Dim startingPrice As Double
    Dim endingPrice As Double
    'Activate the data worksheet. Find the number of rows to loop over
    Worksheets("2018").Activate
    'Find number of rows (before both loops)
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    '4. Loop through the tickers
    For i = 0 To 11
        totalvolume = 0
        startingPrice = 0
        endingPrice = 0
        ticker = tickers(i)
        '5. Loop through rows in the data
        Worksheets("2018").Activate
        For j = 2 To rowEnd
            'Find the total volume for the current ticker
            If Cells(j, 1).Value = ticker Then
                totalvolume = totalvolume + Cells(j, 8).Value
            End If
            'Find the starting price for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then

                startingPrice = Cells(j, 6).Value

            End If
            'Find the ending price for the current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then

                endingPrice = Cells(j, 6).Value
            End If
        Next j
        '6. Output the data for the current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = "2018"
        Cells(4 + i, 2).Value = totalvolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    Next i
End Sub



'for loops*********************************
Sub volume()
	'Declare variables
	Dim totalvolume, rowStart, rowEnd As Integer
	Dim startingPrice As Double
	Dim endingPrice As Double
	'initialize variables
	startingPrice = 0
	rowStart = 2
	'rowEnd = 3013
	Worksheets("2018").Activate
	rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
	totalvolume = 0 
	For i = rowStart To rowEnd
		'increase totalVolume if ticker is "DQ"
		If Cells(i, 1).Value ="DQ" Then
			totalvolume = totalvolume + Cells(i, 8).Value
		End If
		If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then

			startingPrice = Cells(i, 6).Value 

		End If
		If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then

			endingPrice = Cells(i, 6).Value
		End If
	Next i
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1	
	'MsgBox (totalvolume)

End Sub

Sub AllStocksAnalysis()

    Worksheets("All Stocks Analysis").Activate

    Range("A1").Value = "All Stocks (2018)"

    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
End Sub


Sub SkillDrill()
	'write 1 into A1 - J10, A1 is (1, 1) J10 is (10, 10)
    For row = 1 To 10

        'a line of code here will run 10 times

    	For clmn = 1 To 10

            cells(row, clmn) = 1

        Next j

    Next i

End Sub


'If then Conditionals***************************


' instructor's example
Sub test()
    Dim totalCharged as Double
    totalCharged = 0
    Dim startRow, endRow as Integer
    startRow = 2
    endRow = 101
    Dim cc_index as Integer
    cc_index = 1
    For i = startRow to endRow
      totalCharged  = Cells(i, 3).Value + totalCharged
      If Cells(i, 1).Value <> Cells(i+1, 1).Value Then
        cc_index = cc_index + 1
        Range("G" & cc_index).Value = Cells(i, 1).Value
        Range("H" & cc_index).Value = totalCharged
        totalCharged = 0
      End If
    Next i
    ' For i = startRow to endRow
    '   totalCharged  = Cells(i, 3).Value + totalCharged
    '   If ((i-1) < startRow) && (Cells(i, 1).Value <> Cells(i-1, 1).Value) Then
    '     cc_index = cc_index + 1
    '     Range("G" & cc_index).Value = Cells(i, 1).Value
    '     Range("H" & cc_index).Value = totalCharged
    '     totalCharged = 0
    '   End If
    ' Next i
End Sub


Sub NewWorkbook()
    'Make a list of square numbers
    For i = 1 To 10
    
        Cells(1, i).Value = i * i
        
    Next i


End Sub


