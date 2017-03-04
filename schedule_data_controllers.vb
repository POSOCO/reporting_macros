' Schedule Data Controllers
Option Explicit

Function getTableHRange(inp As String) As String
    If inp = "CONST_SCH" Then
        getTableHRange = "[DATA_MVC.xlsx]CONST_SCH!$1:$1"
    ElseIf inp = "ISGS_DC" Then
        getTableHRange = "[DATA_MVC.xlsx]ISGS_DC!$1:$1"
    ElseIf inp = "ISGS_SCH" Then
        getTableHRange = "[DATA_MVC.xlsx]ISGS_SCH!$1:$1"
    ElseIf inp = "FLOW_GATE_SCH" Then
        getTableHRange = "[DATA_MVC.xlsx]FLOW_GATE_SCH!$1:$1"
    Else
        getTableHRange = ""
    End If
End Function

Function getTableVRange(inp As String) As String
    If inp = "CONST_SCH" Then
        getTableVRange = "[DATA_MVC.xlsx]CONST_SCH!$A:$A"
    ElseIf inp = "ISGS_DC" Then
        getTableVRange = "[DATA_MVC.xlsx]ISGS_DC!$A:$A"
    ElseIf inp = "ISGS_SCH" Then
        getTableVRange = "[DATA_MVC.xlsx]ISGS_SCH!$A:$A"
    ElseIf inp = "FLOW_GATE_SCH" Then
        getTableVRange = "[DATA_MVC.xlsx]FLOW_GATE_SCH!$A:$A"
    Else
        getTableVRange = ""
    End If
End Function

Sub Test()
    Dim rng As Range

    Dim modelFileName As String

    modelFileName = "DATA_MVC.xlsx"

    If Not FileIsOpenTest(modelFileName) Then
        Workbooks.Open Filename:="" & modelFileName
    End If

    Set rng = Range(getTableHRange("ISGS_DC"))
    MsgBox (rng.Cells(1, 2))
End Sub

Function MVC_GET_STATE_SCH(state_Str As String, attr As String, timeBlkStr As String) As String
    Application.Volatile True

    Dim tHRng As Range
    Dim tVRng As Range
    Dim hSearchStr As String
    Dim hBSearchStr As String
    Dim vSearchStr As String

    Dim res As String

    Dim modelFileName As String

    hBSearchStr = attr
    hSearchStr = state_Str
    vSearchStr = timeBlkStr

    modelFileName = "DATA_MVC.xlsx"

    If Not FileIsOpenTest(modelFileName) Then
        Workbooks.Open Filename:=ActiveWorkbook.Path & "\" & modelFileName
    End If

    Set tHRng = Range(getTableHRange("CONST_SCH"))
    Set tVRng = Range(getTableVRange("CONST_SCH"))

    If timeBlkStr = "MU" Then
        Set tVRng = Range(getTableVRange("CONST_SCH")).Offset(ColumnOffset:=1)
        vSearchStr = "MWHR"
    End If

    If attr = "OA" Then
        hBSearchStr = "STOA"
    End If

    res = NAG_TABLE_SEARCH_TWO(tHRng, hSearchStr, tHRng.Offset(RowOffset:=1), hBSearchStr, tVRng, vSearchStr).Cells(1, 1).Value

    If attr = "OA" Then
        hBSearchStr = "IEX"
        res = res + NAG_TABLE_SEARCH_TWO(tHRng, hSearchStr, tHRng.Offset(RowOffset:=1), hBSearchStr, tVRng, vSearchStr).Cells(1, 1).Value
        hBSearchStr = "PXI"
        res = res + NAG_TABLE_SEARCH_TWO(tHRng, hSearchStr, tHRng.Offset(RowOffset:=1), hBSearchStr, tVRng, vSearchStr).Cells(1, 1).Value
    End If

    If timeBlkStr = "MU" Then
        res = res / 1000
    End If

    MVC_GET_STATE_SCH = res

End Function
