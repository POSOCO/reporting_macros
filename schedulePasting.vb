Function GetFileType(fileString As String) As String
'
' Macro1 Macro
' http://www.exceltrick.com/formulas_macros/vba-substring-function/ - String extraxtions
'
Dim startPos As Integer
If (fileString Like "Report-RLDC-Dec-WEST(*") Then
    GetFileType = "ISGS-DC"
        Exit Function
ElseIf (fileString Like "FullSchedule-InjectionSummary-ALL_Seller(*") Then
    GetFileType = "ISGS-SCH"
        Exit Function
ElseIf (fileString Like "FlowGate-Schedule-RevNo(*") Then
    GetFileType = "Flow Gate Schedule"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-ALL_Buyer(*") Then
    GetFileType = "AllConsSchdule"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-CSEB_State(*") Then
    GetFileType = "CSEB"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-DD_State(*") Then
    GetFileType = "DD"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-DNH_State(*") Then
    GetFileType = "DNH"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-ESIL_WR_State(*") Then
    GetFileType = "ESIL"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-GEB_State(*") Then
    GetFileType = "GUJ"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-GOA_State(*") Then
    GetFileType = "GOA"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-MP_State(*") Then
    GetFileType = "MP"
        Exit Function
ElseIf (fileString Like "NetSchedule-Summary-MSEB_State(*") Then
    GetFileType = "MSEB"
        Exit Function
End If
GetFileType = "NA"
'
End Function

Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com
Dim endMessage As String
endMessage = "Pasted Sheets are "
Dim thisFileName As String
thisFileName = "SCHEDULE COMPUTATION_WBES.xls"

Dim wb As Workbook
Dim myPath As String
Dim myFile As String
Dim myExtension As String
Dim FldrPicker As FileDialog

'Optimize Macro Speed
  Application.ScreenUpdating = False
  Application.EnableEvents = False
  Application.Calculation = xlCalculationManual

'In Case of Cancel
  myPath = ActiveWorkbook.Path & "\files\"
  If myPath = "" Then GoTo ResetSettings

'Target File Extension (must include wildcard "*")
  myExtension = "*"

'Target Path with Ending Extention
  myFile = Dir(myPath & myExtension)

'Loop through each Excel file in folder
  Do While myFile <> ""
    'MsgBox (myFile)
    If myFile Like "*.csv" Or myFile Like "*.xlsx" Then
        'MsgBox (GetFileType(myFile))
        'Set variable equal to opened workbook
          Set wb = Workbooks.Open(Filename:=myPath & myFile)
        
        'Ensure Workbook has opened before moving on to next line of code
          DoEvents
        Application.DisplayAlerts = False
        Dim sheetName As String
        sheetName = GetFileType(myFile)
        If Not sheetName = "NA" Then
            wb.Worksheets(1).UsedRange.Copy
            Windows(thisFileName).Activate
            Sheets(sheetName).Select
            Range("A1").Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        endMessage = endMessage & sheetName & ", "
        End If

        'Save and Close Workbook
          wb.Close SaveChanges:=True
          
        'Ensure Workbook has closed before moving on to next line of code
          DoEvents
    End If
    'Get next file name
      myFile = Dir
  Loop
Application.DisplayAlerts = True
'Message Box when tasks are completed
  MsgBox "Task Complete!" & endMessage
  'Windows(thisFileName).Worksheets("BUTTON").Cells(1, 1).Value = endMessage
ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
