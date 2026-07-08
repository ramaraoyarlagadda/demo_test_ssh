'***************************************************************************************************************
' FX21 ONLY - Temporary executable block to run from ~line 60 in UFT
'
' HOW TO USE:
'  1. Keep your existing setup code (FileSystem, DataTable import, report files, screenshots folder).
'  2. Paste this block in place of the broken Select Case / FX21 section.
'  3. Comment out "On Error Resume Next" while debugging.
'  4. Confirm keyDealUsingFX21 is a Function returning True/False (see notes at bottom).
'***************************************************************************************************************

channelSource = "FX21"
TestStep = 1

Select Case channelSource

    Case "FX21"

        TestStep = TestStep + 1

        ' --- Option A: Function returns Boolean ---
        dealKeyResult = keyDealUsingFX21(TestStep)
        If dealKeyResult = True Then

            ' >>> Put your FX21 / FM33M / FX41 / FX78 logic here <<<
            ' (Use the full cleaned script: FX21_Deal_Validation_Executable.vbs)

        End If  ' dealKeyResult

        ' --- Option B: If keyDealUsingFX21 is a Sub, use this instead of Option A ---
        ' Call keyDealUsingFX21(TestStep)
        ' Then continue with FX21 logic (no If dealKeyResult)

        ' --- Option C: If function returns "PASS"/"FAIL" ---
        ' dealKeyResult = keyDealUsingFX21(TestStep)
        ' If UCase(CStr(dealKeyResult)) = "PASS" Then
        '     ' FX21 logic
        ' End If

End Select

'***************************************************************************************************************
' REQUIRED CLOSURE ORDER for the FULL script (not just this snippet):
'
'   End If      ' closes If keyDealUsingFX21 / dealKeyResult
'   End Select  ' closes Select Case channelSource
'   End If      ' closes If ProcessThisRow
'   Next        ' closes For RowNum = 1 To RowCount
'
' Do NOT write: Next RowNum   (can cause UFT parse issues)
' Write only:   Next
'***************************************************************************************************************
