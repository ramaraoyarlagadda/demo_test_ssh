'***************************************************************************************************************
' FX21 ONLY - Temporary executable block to run from ~line 60 in UFT
'
' REQUIRED BEFORE THIS BLOCK:
'   Paste Fix_ResultFileForThisRun.vbs first, OR keep full setup that sets:
'     Environment("ResultFileForThisRun")
' Otherwise Report_EndOfTCReport fails with:
'   The environment parameter 'ResultFileForThisRun' was not found
'***************************************************************************************************************

Function EnvParamExists(paramName)
    On Error Resume Next
    Dim dummy
    dummy = Environment(paramName)
    EnvParamExists = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

' --- Ensure report env exists (safe to run every time) ---
Set FileSystem = CreateObject("Scripting.FileSystemObject")
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
Environment("ResultFileForThisRun_Detail") = Environment("ResultFolder") & moduleName & "Details.html"
Environment("ResultFileForThisRun") = Environment("ResultFileForThisRun_Detail")
If Not FileSystem.FileExists(Environment("ResultFileForThisRun")) Then
    Call createReportFileDetails(FileSystem, Environment("ResultFileForThisRun"))
End If

channelSource = "FX21"
TestStep = 1

Select Case channelSource

    Case "FX21"

        TestStep = TestStep + 1

        ' keyDealUsingFX21 is a Sub - use Call (NOT assignment)
        Call keyDealUsingFX21(TestStep)

        ' >>> Continue with FX21 / FM33M / FX41 / FX78 logic here <<<

End Select

' Safe end-of-TC report
Call Report_EndOfTCReport(FileSystem, "</table><br><br><br>")
