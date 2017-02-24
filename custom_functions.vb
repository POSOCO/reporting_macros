Option Explicit

Function NAG_TB_VAL(rng As Range, tb As Double)
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
Dim i As Integer
Dim tempRes, tb As Double
NAG_TB_MAX_VAL = 0
For tb = 1 To 96
    tempRes = NAG_TB_VAL(rng, tb)
    If NAG_TB_MAX_VAL < tempRes Then
        NAG_TB_MAX_VAL = tempRes
    End If
Next tb
End Function

Function NAG_TB_MAX_TBLK(rng As Range)
Dim i As Integer
Dim tempRes, maxVal, tb As Double
maxVal = 0
NAG_TB_MAX_TBLK = 1
For tb = 1 To 96
    tempRes = NAG_TB_VAL(rng, tb)
    If maxVal < tempRes Then
        maxVal = tempRes
        NAG_TB_MAX_TBLK = tb
    End If
Next tb
End Function
