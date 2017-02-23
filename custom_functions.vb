Option Explicit

Function NAG_TB_VAL(rng As Range, tb As Double)
Dim FirstRow, FirstCol, i As Integer
If tb = 0 Then
    NAG_TB_VAL = Cells(rng.Row, rng.Column).Value
    Exit Function
End If

FirstRow = rng.Row + 1 + (tb - 1) * 15
FirstCol = rng.Column

NAG_TB_VAL = 0
For i = 0 To 14
    NAG_TB_VAL = NAG_TB_VAL + Cells(FirstRow + i, FirstCol).Value
Next i
NAG_TB_VAL = NAG_TB_VAL / 15
End Function

Function NAG_TB_MAX_VAL(rng As Range)
Dim FirstRow, FirstCol, i As Integer
Dim tempRes, tb As Double
NAG_TB_MAX_VAL = 0
For tb = 0 To 95
    tempRes = NAG_TB_VAL(rng, tb)
    If NAG_TB_MAX_VAL < tempRes Then
        NAG_TB_MAX_VAL = tempRes
    End If
Next tb
End Function