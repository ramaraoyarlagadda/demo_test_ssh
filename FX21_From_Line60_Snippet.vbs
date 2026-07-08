'***************************************************************************************************************
' FX21 ONLY - Temporary executable block to run from ~line 60 in UFT
'
' HOW TO USE:
'  1. Keep your existing setup code (FileSystem, DataTable import, report files, screenshots folder).
'  2. Paste this block in place of the broken Select Case / FX21 section.
'  3. Comment out "On Error Resume Next" while debugging.
'  4. keyDealUsingFX21 is a Sub - use Call (NOT assignment). That fixes:
'       Type mismatch: 'keyDealUsingFX21'
'***************************************************************************************************************

channelSource = "FX21"
TestStep = 1

Select Case channelSource

    Case "FX21"

        TestStep = TestStep + 1

        ' CORRECT - keyDealUsingFX21 is a Sub (no return value)
        Call keyDealUsingFX21(TestStep)

        ' WRONG - causes Type mismatch: 'keyDealUsingFX21'
        ' dealKeyResult = keyDealUsingFX21(TestStep)
        ' If keyDealUsingFX21(TestStep) Then

        ' >>> Continue with FX21 / FM33M / FX41 / FX78 logic here <<<
        ' (Use the full cleaned script: FX21_Deal_Validation_Executable.vbs)

End Select

'***************************************************************************************************************
' REQUIRED CLOSURE ORDER for the FULL script (not just this snippet):
'
'   End Select  ' closes Select Case channelSource
'   End If      ' closes If ProcessThisRow
'   Next        ' closes For RowNum = 1 To RowCount
'
' Do NOT write: Next RowNum   (can cause UFT parse issues)
' Write only:   Next
'***************************************************************************************************************
