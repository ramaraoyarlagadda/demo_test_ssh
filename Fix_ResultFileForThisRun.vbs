'***************************************************************************************************************
' QUICK FIX for UFT error:
'   The environment parameter 'ResultFileForThisRun' was not found
'   at Report_EndOfTCReport -> Environment("ResultFileForThisRun")
'
' CAUSE:
'   You started from the FX21 section (~line 60) and skipped the setup that creates
'   Environment("ResultFileForThisRun").
'
' PASTE THIS BLOCK BEFORE any Report_* call / before Report_EndOfTCReport
'***************************************************************************************************************

Set FileSystem = CreateObject("Scripting.FileSystemObject")

' Helper: check if env param exists (VBScript Or does NOT short-circuit)
Function EnvParamExists(paramName)
    On Error Resume Next
    Dim dummy
    dummy = Environment(paramName)
    EnvParamExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' 1) Result folder (use your real framework path if you already have one)
If Not EnvParamExists("ResultFolder") Then
    Environment("ResultFolder") = "C:\Temp\UFT_Results\"
ElseIf Trim(CStr(Environment("ResultFolder"))) = "" Then
    Environment("ResultFolder") = "C:\Temp\UFT_Results\"
End If
If Right(Environment("ResultFolder"), 1) <> "\" Then
    Environment("ResultFolder") = Environment("ResultFolder") & "\"
End If
If Not FileSystem.FolderExists(Environment("ResultFolder")) Then
    FileSystem.CreateFolder Environment("ResultFolder")
End If

moduleName = "NewModule"

' 2) Create detail/summary report files
Environment("ResultFileForThisRun_Summary") = Environment("ResultFolder") & moduleName & ".html"
Environment("ResultFileForThisRun_Detail") = Environment("ResultFolder") & moduleName & "Details.html"

' 3) THIS is the missing variable that Report_EndOfTCReport needs
Environment("ResultFileForThisRun") = Environment("ResultFileForThisRun_Detail")

' 4) Create the physical files expected by shared library helpers
Call createReportFileDetails(FileSystem, Environment("ResultFileForThisRun_Summary"))
Call report_UpdateSummaryDataHeader(FileSystem, Environment("ResultFileForThisRun_Summary"))
Call createReportFileDetails(FileSystem, Environment("ResultFileForThisRun_Detail"))

' 5) Screenshots folder
If Not FileSystem.FolderExists(Environment("ResultFolder") & "Screenshots_" & moduleName) Then
    FileSystem.CreateFolder Environment("ResultFolder") & "Screenshots_" & moduleName
End If
Environment("ScreenShots") = Environment("ResultFolder") & "Screenshots_" & moduleName & "\"
Environment("Screenshots") = Environment("ScreenShots")

' Now this will work:
' Call Report_EndOfTCReport(FileSystem, "</table><br><br><br>")

'***************************************************************************************************************
' MINIMUM one-liner if ResultFolder already exists and you only need the env var:
'
'   Environment("ResultFileForThisRun") = Environment("ResultFolder") & "NewModuleDetails.html"
'
'***************************************************************************************************************
