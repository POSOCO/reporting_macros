Function GetFileType(fileString As String) As String
'
' Macro1 Macro
' http://www.exceltrick.com/formulas_macros/vba-substring-function/ - String extraxtions
'
Dim startPos As Integer
If (fileString Like "Report-RLDC-Dec-WEST(*") Then
    GetFileType = "Decleration"
        Exit Function
End If
If (fileString Like "FullSchedule-InjectionSummary-ALL_Seller(*") Then
    GetFileType = "InjectionSchedule"
        Exit Function
End If
If (fileString Like "FlowGate-Schedule-RevNo(*") Then
    GetFileType = "FlowGateSchedule"
        Exit Function
End If
If (fileString Like "NetSchedule-Summary-ALL_Buyer(*") Then
    GetFileType = "AllConsSchdule"
        Exit Function
End If
GetFileType = "NA"
'
End Function

Sub LoopAllExcelFilesInFolder()
'PURPOSE: To loop through all Excel files in a user specified folder and perform a set task on them
'SOURCE: www.TheSpreadsheetGuru.com

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
        
        If GetFileType(myFile) = "Decleration" Then
            wb.Worksheets(1).UsedRange.Copy
            Sheets("DC").Range("A1").Select
            
        ElseIf GetFileType(myFile) = "InjectionSchedule" Then
            wb.Worksheets(1).UsedRange.Copy
            Sheets("SCH").Range("A1").Select
        End If
        
        
        'Save and Close Workbook
          wb.Close SaveChanges:=True
          
        'Ensure Workbook has closed before moving on to next line of code
          DoEvents
    End If
    'Get next file name
      myFile = Dir
  Loop

'Message Box when tasks are completed
  MsgBox "Task Complete!"

ResetSettings:
  'Reset Macro Optimization Settings
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub
