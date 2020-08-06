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


