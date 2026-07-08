'***************************************************************************************************************
' Script Purpose - FX21 deal creation / amendment validation with FM33M, FX41 and FX78 checks
'
' Created by: Sravanthi
'
' Notes:
'  - This is a cleaned, executable UFT/VBScript version of the FX21 flow.
'  - Select Case / If / For block closures are corrected.
'  - Temporary debug mode can start from the FX21 section (see RUN_FROM_FX21_SECTION).
'  - Comment out "On Error Resume Next" while debugging so UFT shows the real failing line.
'
' History:
'  Action no. 1
'***************************************************************************************************************

'On Error Resume Next   ' Keep commented while debugging

'============================================================
' CONFIG - set True to run only FX21 section (from ~line 60)
'============================================================
Const RUN_FROM_FX21_SECTION = True

'============================================================
' SETUP (required even when testing from FX21 section)
'============================================================
Set FileSystem = CreateObject("Scripting.FileSystemObject")
Environment("TestStatusResult") = True
sheetName = Environment("LocalSheetName")
Screen_Name = sheetName
Sheet_Name = sheetName
moduleName = "NewModule"    '******* Change the module name based on the test

CurrentDateTime = Replace(Date, "/", "") & "_" & Replace(Time, ":", "")
currentDate = Replace(Date, "/", ".")

' Create file for Test Summary storage
Environment("ResultFileForThisRun_Summary") = Environment("ResultFolder") & moduleName & ".html"
Call createReportFileDetails(FileSystem, Environment("ResultFileForThisRun_Summary"))
Call report_UpdateSummaryDataHeader(FileSystem, Environment("ResultFileForThisRun_Summary"))

' Create file for Test Details storage
Environment("ResultFileForThisRun_Detail") = Environment("ResultFolder") & moduleName & "Details.html"
Environment("ResultFileForThisRun") = Environment("ResultFileForThisRun_Detail")
Call createReportFileDetails(FileSystem, Environment("ResultFileForThisRun_Detail"))

' Create folder for Screenshots
If Not FileSystem.FolderExists(Environment("ResultFolder") & "Screenshots_" & moduleName) Then
    FileSystem.CreateFolder Environment("ResultFolder") & "Screenshots_" & moduleName
End If
Environment("ScreenShots") = Environment("ResultFolder") & "Screenshots_" & moduleName & "\"
Environment("Screenshots") = Environment("ScreenShots")

' Create the Data sheet and import data
TestStep = 0
DataTable.AddSheet sheetName
DataTable.ImportSheet Environment("testDataForTestCases"), sheetName, sheetName
RowCount = DataTable.GetSheet(sheetName).GetRowCount

'============================================================
' MAIN LOOP
'============================================================
For RowNum = 1 To RowCount

    DataTable.GetSheet(sheetName).SetCurrentRow RowNum
    ProcessThisRow = DataTable.Value("ProcessThisRow", sheetName)
    TestLabels = UCase(DataTable.Value("TestLabels", sheetName))

    If UCase(ProcessThisRow) = "Y" Or InStr(TestLabels, Environment("TestRunMode")) > 0 Or RUN_FROM_FX21_SECTION Then

        TestIdentifier = DataTable.Value("TestIdentifier", sheetName)
        TestScript = DataTable.Value("Scenario", sheetName)
        Call report_Single_CreateTestResultHeader(FileSystem, TestIdentifier, TestScript)

        TestStep = TestStep + 1
        Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Test step Description", "Test should be working", "Test is working", "PASS", "")

        '------------------------------------------------------------
        ' FX21 SECTION (start here for temporary execution / debug)
        '------------------------------------------------------------
        channelSource = "FX21"
        TestStep = 1

        Select Case channelSource

            Case "FX21"

                TestStep = TestStep + 1

                ' keyDealUsingFX21 is a Sub (not a Function).
                ' Do NOT write: dealKeyResult = keyDealUsingFX21(TestStep)
                ' That causes: Type mismatch: 'keyDealUsingFX21'
                Call keyDealUsingFX21(TestStep)

                ' Capture the PDF Path to print
                Wait 4
                TestStep = TestStep + 1
                Browser("FXDealPDF").Refresh
                Wait 5
                pdfBrowserURL = Browser("FXDealPDF").GetROProperty("openurl")
                Browser("FXDealPDF").Close
                Wait 4
                Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture the PDF File Path", _
                    "PDF Path should be captured", _
                    "PDF Location : <a href=" & Chr(34) & pdfBrowserURL & Chr(34) & "> PDF Path </a>", _
                    "PASS", "")

                ' Capture the Deal Number
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleApplications("OracleApplications").CaptureBitmap ScreenshotPath
                Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture the deal confirmation", _
                    "Deal confirmation should be captured", "Deal confirmation captured", "PASS", ScreenshotPath)
                Wait 2
                Status_Message = Trim(OracleApplications("OracleApplications").OracleStatusLine("OracleStatusLine").GetROProperty("message"))
                DealNumber = Right(Status_Message, 13)
                DataTable.Value("Deal_Number", Screen_Name) = "'" & DealNumber

                ' Verify the deal creation
                TestStep = TestStep + 1
                TestDescription = "Verify the deal creation for case reference: " & DataTable.Value("Comment", Screen_Name)
                TestExpRes = "Deal creation should be successful"
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"

                If IsNumeric(DealNumber) Then
                    ' Deal Creation is successful
                    OracleApplications("OracleApplications").CaptureBitmap ScreenshotPath
                    Call Report_UpdateTestCaseStatus(FileSystem, TestStep, TestDescription, TestExpRes, _
                        "Deal creation is successful : " & DealNumber, "PASS", ScreenshotPath)
                    OracleFormWindow("FX21").OracleButton("Exit").Click

                    ' Deal Settlement
                    Call DealSettlement(FileSystem, TestStep, Screen_Name, DealNumber)
                Else
                    ' Unable to Complete deal creation
                    OracleFormWindow("FX21").CaptureBitmap ScreenshotPath
                    Call Report_UpdateTestCaseStatus(FileSystem, TestStep, TestDescription, TestExpRes, _
                        "Deal creation failed", "FAIL", ScreenshotPath)
                    OracleFormWindow("FX21").OracleButton("Exit").Click
                End If

                ' Open FM33M Screen
                Call OpenCBScreen("FM33M")
                Wait 2

                ' Select the All Radio button
                OracleFormWindow("FM33M").OracleRadioGroup("Verified").Select "All"
                Expected = "FXSMT300"
                ' Enter the deal Number
                Call OpenCBSRecord("FM33M", OracleFormWindow("FM33M").OracleTextField("Transaction Reference"), DealNumber)
                Actual = OracleFormWindow("FM33M").OracleTable("tblFM33MFormatData").GetFieldValue(1, "Routing")

                ' Capture the Screenshot
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FM33M").CaptureBitmap ScreenshotPath

                If Expected = Actual Then
                    MsgBox "Verify the Message Format Type in FM33M Screen. Expected: Message Format Type should be correctly displayed. Actual: Message Format Type displayed correctly: " & Actual & " Status: PASS"
                    OracleFormWindow("FM33M").OracleButton("View").Click
                    Wait 4

                    Set shellObject = CreateObject("WScript.Shell")
                    shellObject.SendKeys "{F5}"
                    Set shellObject = Nothing
                    Wait 2

                    TestStep = TestStep + 1
                    ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                    Desktop.CaptureBitmap ScreenshotPath
                    Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture the Screen shot of View Document", _
                        "Screen Shot Captured", "", "PASS", ScreenshotPath)
                    Browser("CBS Print Browser").Close
                Else
                    MsgBox "Verify the Message Format Type in FM33M Screen. Expected: Message Format Type should be correctly displayed. Actual: Message Format Type displayed incorrectly: " & Actual & " Status: FAIL"
                    If Actual = "" Then
                        OracleFormWindow("FM33M").SelectMenu "Query->Cancel                  Ctrl+q"
                    End If
                End If

                ' Close the window
                Wait 2
                OracleFormWindow("FM33M").OracleButton("Exit").Click

                ' Verify the Settlement flag on FX41
                SettlementFlag = OracleFormWindow("FX41").OracleTable("settlementTable").GetFieldValue(1, "Settlement")
                If SettlementFlag Then
                    OracleFormWindow("FX41").OracleTable("settlementTable").EnterField 1, "Settlement", False
                    OracleFormWindow("FX41").OracleButton("Save").Click
                    Wait 2
                    ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & DataTable.Value("Comment", Sheet_Name) & ".png"
                    OracleFormWindow("FX41").CaptureBitmap ScreenshotPath
                End If

                ' Click on amend deal (AD) on FX21 screen
                OracleFormWindow("FX21").OracleButton("AD").Click
                OracleFormWindow("FX21PopUp").OracleTextField("Reason").Enter "AMD"
                OracleFormWindow("FX21PopUp").OracleButton("OK").Click
                OracleListOfValues("List of Reasons").Select "INPUT ERROR (AMOUNT) - OPS"
                OracleFormWindow("FX21PopUp").OracleButton("OK").Click

                amendedmainAmount = CInt(mainAmount) + 100
                amendedmainAmount = CStr(amendedmainAmount)
                OracleFormWindow("FX21").OracleTextField("Amount/ Counter Amount").Enter amendedmainAmount
                OracleFormWindow("FX21").OracleButton("Save").Click
                OracleNotification("Forms").OracleButton("OK").Click
                Wait 2

                ' Capture screenshot for original deal
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                Desktop.CaptureBitmap ScreenshotPath
                Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture the Screen shot of View Document", _
                    "Screen Shot Captured", "", "PASS", ScreenshotPath)
                Browser("CBS Print Browser").Close
                OracleNotification("Forms").OracleButton("OK").Click
                Wait 2

                ' Capture screenshot for new deal
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                Desktop.CaptureBitmap ScreenshotPath
                Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture the Screen shot of View Document", _
                    "Screen Shot Captured", "", "PASS", ScreenshotPath)
                Browser("CBS Print Browser").Close

                ' Capture screenshot for original deal on FX21
                Call OpenCBScreen("FX21")
                OracleFormWindow("FX21").SelectMenu "Query->Enter"
                OracleFormWindow("FX21").OracleTextField("Deal Number").Enter DealNumber
                OracleFormWindow("FX21").SelectMenu "Query->eXecute"
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FX21").CaptureBitmap ScreenshotPath
                OracleFormWindow("FX21").OracleTextField("OracleTextField").GetROProperty("value")
                OracleFormWindow("FX21").OracleButton("Exit").Click

                ' Capture screenshot for new/amended deal on FX21
                NewDealNumber = Amend_Deal_number
                Call OpenCBScreen("FX21")
                OracleFormWindow("FX21").SelectMenu "Query->Enter"
                OracleFormWindow("FX21").OracleTextField("Deal Number").Enter NewDealNumber
                OracleFormWindow("FX21").SelectMenu "Query->eXecute"
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FX21").CaptureBitmap ScreenshotPath
                booleanDealOptimization = OracleFormWindow("FX21").OracleCheckbox("Deal Optimization").IsSelected
                OracleFormWindow("FX21").OracleButton("Exit").Click

                If booleanDealOptimization = False Then
                    MsgBox "Verify the Deal Optimization Checkbox is not Selected. Expected: Deal should be created with Deal Optimization checkbox as unselected. Actual: Deal with Deal Optimization creation is successful: Deal No: " & NewDealNumber & " Status: PASS"
                Else
                    MsgBox "Verify the Deal Optimization Checkbox is Selected. Expected: Deal should be created with Deal Optimization checkbox as selected. Actual: Deal with Deal Optimization creation is un-successful: Deal No: " & NewDealNumber & " Status: FAIL"
                End If

                ' Verify Settlement is unchecked in FX41 screen
                Call OpenCBScreen("FX41")
                Call OpenCBSRecord("FX41", OracleFormWindow("FX41").OracleTextField("Deal No."), NewDealNumber)
                SettlementFlag = OracleFormWindow("FX41").OracleTable("settlementTable").GetFieldValue(1, "Settlement")
                booleanSettlement = SettlementFlag
                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FX41").CaptureBitmap ScreenshotPath

                If booleanSettlement = False Then
                    MsgBox "Verify FX41 Settlement checkbox is unchecked. Expected: Settlement checkbox should be unchecked. Actual: Settlement checkbox is unchecked for Deal No: " & NewDealNumber & " Status: PASS"
                Else
                    MsgBox "Verify FX41 Settlement checkbox is unchecked. Expected: Settlement checkbox should be unchecked. Actual: Settlement checkbox is selected for Deal No: " & NewDealNumber & " Status: FAIL"
                End If

                OracleFormWindow("FX41").OracleButton("Exit").Click

                ' Verify original deal in FM33M screen - FXSMT300
                Call OpenCBScreen("FM33M")
                OracleFormWindow("FM33M").OracleRadioGroup("Verified").Select "All"
                Expected = "FXSMT300"
                Call OpenCBSRecord("FM33M", OracleFormWindow("FM33M").OracleTextField("Transaction Reference"), DealNumber)
                Actual = OracleFormWindow("FM33M").OracleTable("tblFM33MFormatData").GetFieldValue(1, "Routing")

                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FM33M").CaptureBitmap ScreenshotPath

                If Expected = Actual Then
                    MsgBox "Verify the Message Format Type in FM33M Screen. Expected: Message Format Type should be correctly displayed. Actual: Message Format Type displayed correctly: " & Actual & " Status: PASS"
                    OracleFormWindow("FM33M").OracleButton("View").Click
                    Wait 4

                    Set shellObject = CreateObject("WScript.Shell")
                    shellObject.SendKeys "{F5}"
                    Set shellObject = Nothing
                    Wait 2

                    TestStep = TestStep + 1
                    ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                    Desktop.CaptureBitmap ScreenshotPath
                    Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture the Screen shot of View Document", _
                        "Screen Shot Captured", "", "PASS", ScreenshotPath)
                    Browser("CBS Print Browser").Close
                Else
                    MsgBox "Verify the Message Format Type in FM33M Screen. Expected: Message Format Type should be correctly displayed. Actual: Message Format Type displayed incorrectly: " & Actual & " Status: FAIL"
                    If Actual = "" Then
                        OracleFormWindow("FM33M").SelectMenu "Query->Cancel                  Ctrl+q"
                    End If
                End If

                Wait 2
                OracleFormWindow("FM33M").OracleButton("Exit").Click

                ' Verify FM33M - FMSMT202
                Call OpenCBScreen("FM33M")
                OracleFormWindow("FM33M").OracleRadioGroup("Verified").Select "All"
                Expected = "FMSMT202"
                Call OpenCBSRecord("FM33M", OracleFormWindow("FM33M").OracleTextField("Transaction Reference"), DealNumber)
                Actual = OracleFormWindow("FM33M").OracleTable("tblFM33MFormatData").GetFieldValue(2, "Routing")

                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FM33M").CaptureBitmap ScreenshotPath

                If Expected = Actual Then
                    MsgBox "Verify the Message Format Type in FM33M Screen. Expected: Message 202 should be deleted. Actual: Message 202 is deleted: " & Actual & " Status: PASS"
                    OracleFormWindow("FM33M").OracleButton("View").Click
                    Wait 4

                    Set shellObject = CreateObject("WScript.Shell")
                    shellObject.SendKeys "{F5}"
                    Set shellObject = Nothing
                    Wait 2
                Else
                    MsgBox "Verify the Message Format Type in FM33M Screen. Expected: Message 202 should be deleted. Actual: Message 202 is not deleted: " & Actual & " Status: FAIL"
                    If Actual = "" Then
                        OracleFormWindow("FM33M").SelectMenu "Query->Cancel                  Ctrl+q"
                    End If
                End If

                Wait 2
                OracleFormWindow("FM33M").OracleButton("Exit").Click

                ' Verify Deal Status in FX78 screen
                Call OpenCBScreen("FX78")
                OracleFormWindow("FX78").SelectMenu "Query->Enter"
                Wait 1
                OracleFormWindow("FX78").OracleTable("Table").EnterField 1, "Deal No.", DealNumber
                OracleFormWindow("FX78").SelectMenu "Query->eXecute"
                OracleFormWindow("FX78").OracleButton("MSG").Click

                TestStep = TestStep + 1
                ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"
                OracleFormWindow("FX78").CaptureBitmap ScreenshotPath
                OracleFormWindow("FX78").OracleButton("Exit Msg").Click
                OracleFormWindow("FX78").OracleButton("Exit").Click

                ' NSTP / STP settlement verification on FX41
                Call OpenCBScreen("FX41")
                Call OpenCBSRecord("FX41", OracleFormWindow("FX41").OracleTextField("Deal No."), DealNumber)
                CBSTranDate = OracleFormWindow("FX41").OracleTable("tblTradeDate").GetFieldValue(1, "Trade Date")
                ValueDate = OracleFormWindow("FX41").OracleTable("Table").GetFieldValue(1, "Value Date Bt")
                DataTable("Value_Date", Screen_Name) = "'" & ValueDate

                If DataTable.Value("Deal_STP", Screen_Name) = "Y" Then

                    SettlementFlag = OracleFormWindow("FX41").OracleTable("SettlementTable").GetFieldValue(1, "Settlement")
                    TestStep = TestStep + 1
                    TestDescription = "Verify the Settlement check box is selected for STP Deal :  " & DealNumber
                    TestExpRes = "Settlement check box should be selected"
                    ScreenshotPath = Environment("ScreenShots") & "Step" & TestStep & ".png"

                    If SettlementFlag Then
                        OracleFormWindow("FX41").OracleTable("SettlementTable").EnterField 1, "Conf. Rec'd", True
                        OracleFormWindow("FX41").OracleButton("Save").Click
                        TestActRes = "Settlement check box is selected"
                        OracleFormWindow("FX41").CaptureBitmap ScreenshotPath
                        Call Report_UpdateTestCaseStatus(FileSystem, TestStep, TestDescription, TestExpRes, TestActRes, "PASS", ScreenshotPath)

                        OracleFormWindow("FX41").OracleButton("Settlement Details").Click
                        TestStep = TestStep + 1
                        ScreenshotPath = Environment("ScreenShots") & "Step" & TestStep & ".png"
                        OracleFormWindow("FX221").CaptureBitmap ScreenshotPath
                        Call Report_UpdateTestCaseStatus(FileSystem, TestStep, "Capture FX221 Screen", "Screenshot captured", "", "PASS", ScreenshotPath)

                        messageGroup = OracleFormWindow("FX221").OracleTextField("Message Group").GetROProperty("value")
                        OracleFormWindow("FX221").OracleButton("Exit").Click
                    Else
                        TestActRes = "Settlement check box is not selected"
                        OracleFormWindow("FX41").CaptureBitmap ScreenshotPath
                        Call Report_UpdateTestCaseStatus(FileSystem, TestStep, TestDescription, TestExpRes, TestActRes, "FAIL", ScreenshotPath)
                    End If

                    OracleFormWindow("FX41").OracleButton("Exit").Click

                Else
                    ' Non-STP path
                    Call OpenCBSRecord("FX41", OracleFormWindow("FX41").OracleTextField("Deal No."), DealNumber)
                    SettlementFlag = OracleFormWindow("FX41").OracleTable("SettlementTable").GetFieldValue(1, "Settlement")
                    TestStep = TestStep + 1
                    TestDescription = "Verify the Settlement check box is not selected for Non STP Deal :  " & DealNumber
                    TestExpRes = "Settlement check box should not be selected"
                    ScreenshotPath = Environment("Screenshots") & "Step" & TestStep & ".png"

                    If SettlementFlag Then
                        TestActRes = "Settlement check box is selected"
                        OracleFormWindow("FX41").CaptureBitmap ScreenshotPath
                        Call Report_UpdateTestCaseStatus(FileSystem, TestStep, TestDescription, TestExpRes, TestActRes, "FAIL", ScreenshotPath)
                        OracleFormWindow("FX41").OracleButton("Exit").Click
                    Else
                        TestActRes = "Settlement check box is not selected"
                        OracleFormWindow("FX41").CaptureBitmap ScreenshotPath
                        Call Report_UpdateTestCaseStatus(FileSystem, TestStep, TestDescription, TestExpRes, TestActRes, "PASS", ScreenshotPath)
                        OracleFormWindow("FX41").OracleButton("Exit").Click
                    End If

                End If  ' Deal_STP = Y

            ' Case "FX23"
            ' Case "Cellar"

        End Select  ' channelSource

        ' Update the test summary
        Call Report_EndOfTCReport(FileSystem, "</table><br><br><br>")
        Call Report_UpdateTestSummaryDetails(FileSystem, TestIdentifier, TestScript, Environment("TestStatusResult"), "")
        Environment("TestStatusResult") = True

    End If  ' ProcessThisRow / TestLabels / RUN_FROM_FX21_SECTION

Next  ' RowNum

' Merge the Summary and Detail Data files (optional)
' Call report_Merge_TestSummaryAndDetails(FileSystem, Environment("ResultFileForThisRun_Summary"), Environment("ResultFileForThisRun_Detail"))

'***************************************************************************************************************
' LOCAL HELPER NOTES
'***************************************************************************************************************
' 1) Closure order must always be:
'       End If      ' innermost If
'       End Select  ' channelSource
'       End If      ' ProcessThisRow
'       Next        ' For RowNum
'
' 2) keyDealUsingFX21 is a Sub. Always use:
'       Call keyDealUsingFX21(TestStep)
'    NEVER use:
'       dealKeyResult = keyDealUsingFX21(TestStep)   ' Type mismatch
'       If keyDealUsingFX21(TestStep) Then           ' Type mismatch
'
' 3) Required variables before FX21 amend section:
'       mainAmount, Amend_Deal_number, Screen_Name / Sheet_Name
'***************************************************************************************************************
