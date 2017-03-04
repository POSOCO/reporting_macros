' General Controllers
Option Explicit

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
