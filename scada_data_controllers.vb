' SCADA Data Controllers
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
