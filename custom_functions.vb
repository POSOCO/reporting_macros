Option Explicit

Function NAG_TB_VAL(rng As Range, tb As Double)
	Application.Volatile True
	Dim i As Integer
	If tb = 1 Then
		NAG_TB_VAL = rng.Cells(1, 1).Value
		Exit Function
	End If
	NAG_TB_VAL = 0
	For i = 0 To 14
		NAG_TB_VAL = NAG_TB_VAL + rng.Cells(1, 1).Offset((1 + (tb - 2) * 15 + i), 0).Value
	Next i
	NAG_TB_VAL = NAG_TB_VAL / 15
End Function

Function NAG_TB_MAX_VAL(rng As Range)
	Application.Volatile True
	Dim tempRes, tb As Double
	NAG_TB_MAX_VAL = NAG_TB_VAL(rng, 1)
	For tb = 2 To 96
		tempRes = NAG_TB_VAL(rng, tb)
		If NAG_TB_MAX_VAL < tempRes Then
			NAG_TB_MAX_VAL = tempRes
		End If
	Next tb
End Function

Function NAG_TB_MAX_TBLK(rng As Range)
	Application.Volatile True
	Dim tempRes, maxVal, tb As Double
	maxVal = NAG_TB_VAL(rng, 1)
	NAG_TB_MAX_TBLK = 1
	For tb = 2 To 96
		tempRes = NAG_TB_VAL(rng, tb)
		If maxVal < tempRes Then
			maxVal = tempRes
			NAG_TB_MAX_TBLK = tb
		End If
	Next tb
End Function

Function NAG_TB_MIN_VAL(rng As Range)
	Application.Volatile True
	Dim tempRes, tb As Double
	NAG_TB_MIN_VAL = NAG_TB_VAL(rng, 1)
	For tb = 2 To 96
		tempRes = NAG_TB_VAL(rng, tb)
		If NAG_TB_MIN_VAL > tempRes Then
			NAG_TB_MIN_VAL = tempRes
		End If
	Next tb
End Function

Function NAG_TB_MIN_TBLK(rng As Range)
	Application.Volatile True
	Dim tempRes, minVal, tb As Double
	minVal = NAG_TB_VAL(rng, 1)
	NAG_TB_MIN_TBLK = 1
	For tb = 2 To 96
		tempRes = NAG_TB_VAL(rng, tb)
		If minVal > tempRes Then
			minVal = tempRes
			NAG_TB_MIN_TBLK = tb
		End If
	Next tb
End Function

Function NAG_TB_AVG_VAL(rng As Range)
	Application.Volatile True
	Dim tb As Double
	NAG_TB_AVG_VAL = 0
	For tb = 1 To 96
		NAG_TB_AVG_VAL = NAG_TB_AVG_VAL + NAG_TB_VAL(rng, tb)
	Next tb
	NAG_TB_AVG_VAL = NAG_TB_AVG_VAL / 96
End Function

Function NAG_TB_MU_VAL(rng As Range)
	Application.Volatile True
	Dim tb As Double
	NAG_TB_MU_VAL = 0
	For tb = 1 To 96
		NAG_TB_MU_VAL = NAG_TB_MU_VAL + NAG_TB_VAL(rng, tb)
	Next tb
	NAG_TB_MU_VAL = NAG_TB_MU_VAL / 4000
End Function
''UI blocks functions
Function NAG_TB_UI_VAL(schRng As Range, actRng As Range, tb As Double)
	Application.Volatile True
	Dim i As Integer
	If tb = 1 Then
		NAG_TB_UI_VAL = actRng.Cells(1, 1).Value - schRng.Cells(1, 1).Value
		Exit Function
	End If
	NAG_TB_UI_VAL = 0
	For i = 0 To 14
		NAG_TB_UI_VAL = NAG_TB_UI_VAL + actRng.Cells(1, 1).Offset((1 + (tb - 2) * 15 + i), 0).Value - schRng.Cells(1, 1).Offset((1 + (tb - 2) * 15 + i), 0).Value
	Next i
	NAG_TB_UI_VAL = NAG_TB_UI_VAL / 15
End Function

Function NAG_TB_MAX_UI_VAL(schRng As Range, actRng As Range)
	Application.Volatile True
	Dim tempRes, tb As Double
	NAG_TB_MAX_UI_VAL = NAG_TB_UI_VAL(schRng, actRng, 1)
	For tb = 2 To 96
		tempRes = NAG_TB_UI_VAL(schRng, actRng, tb)
		If NAG_TB_MAX_UI_VAL < tempRes Then
			NAG_TB_MAX_UI_VAL = tempRes
		End If
	Next tb
End Function

Function NAG_TB_MAX_UI_TBLK(schRng As Range, actRng As Range)
	Application.Volatile True
	Dim tempRes, maxVal, tb As Double
	maxVal = NAG_TB_UI_VAL(schRng, actRng, 1)
	NAG_TB_MAX_UI_TBLK = 1
	For tb = 2 To 96
		tempRes = NAG_TB_UI_VAL(schRng, actRng, tb)
		If maxVal < tempRes Then
			maxVal = tempRes
			NAG_TB_MAX_UI_TBLK = tb
		End If
	Next tb
End Function

Function NAG_TB_MIN_UI_VAL(schRng As Range, actRng As Range)
	Application.Volatile True
	Dim tempRes, tb As Double
	NAG_TB_MIN_UI_VAL = NAG_TB_UI_VAL(schRng, actRng, 1)
	For tb = 2 To 96
		tempRes = NAG_TB_UI_VAL(schRng, actRng, tb)
		If NAG_TB_MIN_UI_VAL > tempRes Then
			NAG_TB_MIN_UI_VAL = tempRes
		End If
	Next tb
End Function

Function NAG_TB_MIN_UI_TBLK(schRng As Range, actRng As Range)
	Application.Volatile True
	Dim tempRes, minVal, tb As Double
	minVal = NAG_TB_UI_VAL(schRng, actRng, 1)
	NAG_TB_MIN_UI_TBLK = 1
	For tb = 2 To 96
		tempRes = NAG_TB_UI_VAL(schRng, actRng, tb)
		If minVal > tempRes Then
			minVal = tempRes
			NAG_TB_MIN_UI_TBLK = tb
		End If
	Next tb
End Function

Function NAG_TB_AVG_UI_VAL(schRng As Range, actRng As Range)
	Application.Volatile True
	Dim tb As Double
	NAG_TB_AVG_UI_VAL = 0
	For tb = 1 To 96
		NAG_TB_AVG_UI_VAL = NAG_TB_AVG_UI_VAL + NAG_TB_UI_VAL(schRng, actRng, tb)
	Next tb
	NAG_TB_AVG_UI_VAL = NAG_TB_AVG_UI_VAL / 96
End Function

Function NAG_TB_MU_UI_VAL(schRng As Range, actRng As Range)
	Application.Volatile True
	Dim tb As Double
	NAG_TB_MU_UI_VAL = 0
	For tb = 1 To 96
		NAG_TB_MU_UI_VAL = NAG_TB_MU_UI_VAL + NAG_TB_UI_VAL(schRng, actRng, tb)
	Next tb
	NAG_TB_MU_UI_VAL = NAG_TB_MU_UI_VAL / 4000
End Function

Function NAG_HSEARCH(rng As Range, str As String, vOffset As Double) As Range
	Application.Volatile True
	Dim i, sCol, sRow As Integer
	sCol = 0
	For i = 1 To rng.Columns.Count
		If i > 1000 Then
			Exit For
		End If
		If rng.Cells(1, 1).Offset(0, i - 1).Value = str Then
			sCol = i
			Exit For
		End If
	Next i
	Set NAG_HSEARCH = rng.Cells(vOffset + 1, sCol)
End Function

Function NAG_TABLE_SEARCH(hRng As Range, hStr As String, vRng As Range, vStr As String) As Range
	Application.Volatile True
	Dim i, sCol, sRow As Integer
	sCol = 0
	sRow = 0
	For i = 1 To hRng.Columns.Count
		If i > 1000 Then
			Exit For
		End If
		If hRng.Cells(1, 1).Offset(0, i - 1).Value = hStr Then
			sCol = hRng.Column + i - 1
			Exit For
		End If
	Next i
	For i = 1 To vRng.Rows.Count
		If i > 1000 Then
			Exit For
		End If
		If vRng.Cells(1, 1).Offset(i - 1, 0).Value = vStr Then
			sRow = vRng.Row + i - 1
			Exit For
		End If
	Next i
	Set NAG_TABLE_SEARCH = hRng.Worksheet.Cells(sRow, sCol)
End Function

Function NAG_HSEARCH_TWO(topRng As Range, topStr As String, botRng As Range, botStr As String, vOffset As Double) As Range
	Application.Volatile True
	Dim i, sCol, sRow As Integer
	sCol = 0
	sRow = 0
	For i = 1 To botRng.Columns.Count
		If i > 1000 Then
			Exit For
		End If
		If botRng.Cells(1, 1).Offset(0, i - 1).Value = botStr And topRng.Cells(1, 1).Offset(0, i - 1).Value = topStr Then
			sCol = i
			sRow = 0
			Exit For
		End If
	Next i
	Set NAG_HSEARCH_TWO = botRng.Cells(vOffset + 1, sCol)
End Function

Function NAG_TABLE_SEARCH_TWO(hRng As Range, hStr As String, hBRng As Range, hBStr As String, vRng As Range, vStr As String) As Range
	Application.Volatile True
	Dim i, sCol, sRow As Integer
	sCol = 0
	sRow = 0
	For i = 1 To hBRng.Columns.Count
		If i > 1000 Then
			Exit For
		End If
		If hBRng.Cells(1, 1).Offset(0, i - 1).Value = hBStr And hRng.Cells(1, 1).Offset(0, i - 1).Value = hStr Then
			sCol = hBRng.Column + i - 1
			Exit For
		End If
	Next i
	For i = 1 To vRng.Rows.Count
		If i > 1000 Then
			Exit For
		End If
		If vRng.Cells(1, 1).Offset(i - 1, 0).Value = vStr Then
			sRow = vRng.Row + i - 1
			Exit For
		End If
	Next i
	Set NAG_TABLE_SEARCH_TWO = hBRng.Worksheet.Cells(sRow, sCol)
End Function

Function NAG_VSEARCH(rng As Range, str As String, hOffset As Double) As Range
	Application.Volatile True
	Dim i, sCol, sRow As Integer
	sRow = 0
	For i = 1 To rng.Rows.Count
		If i > 1000 Then
			Exit For
		End If
		If rng.Cells(1, 1).Offset(i - 1, 0).Value = str Then
			sRow = i
			Exit For
		End If
	Next i
	Set NAG_VSEARCH = rng.Cells(sRow, hOffset + 1)
End Function

Function FileIsOpenTest(TargetWorkbook As String) As Boolean
	'Step 1: Declare your variables
	Dim TestBook As Workbook
	'Step 2: Tell Excel to resume on error
	On Error Resume Next
	'Step 3: Try to assign the target workbook to TestBook
	Set TestBook = Workbooks(TargetWorkbook)
	'Step 4: If no error occurred, workbook is already open
	If Err.Number = 0 Then
		FileIsOpenTest = True
	Else
		FileIsOpenTest = False
	End If
End Function
