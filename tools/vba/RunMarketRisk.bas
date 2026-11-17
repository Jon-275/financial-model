Attribute VB_Name = "MarketRiskRunner"
Option Explicit

' Financial rules remain in the visible Trial_* worksheets. This module
' selects trials, checks their outputs and publishes one completed batch.
Private Const EXPECTED_SCHEMA As String = "retirement-risk-native-v1"
Private Const EXPECTED_ENGINE As String = "Excel 1.0"
Private Const MAX_PAIRED_PATHS As Long = 1000
Private Const STRESS_CASES As Long = 33
Private Const VALID_FIELD As Long = 2
Private Const FAILURE_AGE_FIELD As Long = 4
Private Const MAX_UNPAID_FIELD As Long = 5
Private Const ENDING_REAL_FIELD As Long = 6
Private Const ENDING_NOMINAL_FIELD As Long = 7
Private Const CUT_YEARS_FIELD As Long = 8
Private Const REFRESH_DATE_FIELD As Long = 27
Private Const FIRST_YEAR_FIELD As Long = 33
Private Const ANNUAL_CAPACITY As Long = 126
Private Const FUNDING_YEAR_WIDTH As Long = 38
Private Const FUNDING_BALANCE_FIELD As Long = 28
Private Const RUN_ERROR As Long = vbObjectError + 5200

Private Type ExcelState
    Calculation As XlCalculation
    EventsEnabled As Boolean
    ScreenUpdating As Boolean
    CancelKey As XlEnableCancelKey
    StatusVisible As Boolean
    StatusText As Variant
    Cursor As XlMousePointer
End Type

Private mRunning As Boolean
Private mCancelRequested As Boolean

Public Sub CancelMarketRisk()
    If mRunning Then mCancelRequested = True
End Sub

Public Sub RunMarketRisk()
    Dim book As Workbook, trials As Worksheet, results As Worksheet, pathSheet As Worksheet
    Dim settings As ExcelState, savedSettings As Boolean
    Dim trialCurrent As Range, trialResults As Range, trialIDs As Range
    Dim pathCurrent As Range, pathResults As Range, pathIDs As Range
    Dim selector As Range, previousSelector As Variant, savedSelector As Boolean
    Dim trialProtection As Variant, resultProtection As Variant
    Dim preparedTrials As Boolean, preparedResults As Boolean, preparedPaths As Boolean
    Dim pathProtection As Variant, selectorRestored As Boolean
    Dim runBegan As Boolean, selectorTouched As Boolean
    Dim batch As Variant, ids As Variant, paths As Variant, chosenIDs As Variant
    Dim oldBatch As Variant, oldIDs As Variant, oldPaths As Variant
    Dim oldPathIDs As Variant, oldMetadata As Variant, metadata As Variant
    Dim captured As Variant, order() As Long, pairedRow As Long
    Dim count As Long, total As Long, width As Long, pathWidth As Long
    Dim seed As Double, context As String, code As Long, index As Long
    Dim rank As Long, caseIndex As Long, column As Long, started As Double
    Dim completed As Boolean, publishing As Boolean, failureNumber As Long
    Dim failureText As String, cleanupText As String, note As String

    If mRunning Then
        MsgBox "Market risk is already running. Use Cancel or press Esc to stop it.", vbInformation, "Market risk"
        Exit Sub
    End If
    mRunning = True
    mCancelRequested = False
    On Error GoTo RunFailed

    Set book = ThisWorkbook
    Set trials = book.Worksheets("Risk_Trials")
    Set results = book.Worksheets("Risk_Data")
    RequireText results.Range("Z2").Value2, EXPECTED_SCHEMA, "Use the workbook and runner from the same release."
    RequireText results.Range("Z4").Value2, EXPECTED_ENGINE, "The workbook's risk engine version does not match the runner."
    Set trialCurrent = LocalRange(trials, "RiskTrialCurrent")
    Set trialResults = LocalRange(trials, "RiskTrialResults")
    Set trialIDs = LocalRange(trials, "RiskTrialIDs")
    Set pathCurrent = LocalRange(trials, "RiskPathCurrent")
    Set pathResults = LocalRange(trials, "RiskPathResults")
    Set pathIDs = LocalRange(trials, "RiskPathTrials")
    ValidateRanges trialCurrent, trialResults, trialIDs, pathCurrent, pathResults, pathIDs
    Set pathSheet = pathResults.Parent
    Set selector = trials.Range("B3")
    RequireNoFormulas selector
    previousSelector = selector.Value2
    savedSelector = True

    SaveExcelState settings
    savedSettings = True
    Application.EnableCancelKey = xlErrorHandler
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = True
    Application.Cursor = xlWait
    Application.Calculation = xlCalculationManual
    PrepareProtection trials, trialProtection, preparedTrials
    PrepareProtection results, resultProtection, preparedResults
    If Not pathSheet Is trials And Not pathSheet Is results Then PrepareProtection pathSheet, pathProtection, preparedPaths
    runBegan = True
    results.Range("Z6").Value2 = "running"
    results.Range("Z9").Value2 = "Running market risk. Use Cancel or press Esc to stop."
    selectorTouched = True
    selector.Value2 = 0
    Application.Calculate
    RequireText book.Worksheets("Market_Paths").Range("AL3").Value2, "OK", _
                "Review the plan and risk assumptions before running market risk."
    count = WholeNumber(book.Worksheets("Risk_Settings").Range("B29").Value2, 1, MAX_PAIRED_PATHS, "Enter 1 to 1,000 paired paths on Risk_Settings.")
    seed = SeedNumber(book.Worksheets("Risk_Settings").Range("B30").Value2)
    context = AssumptionRecord(book)
    total = 2 * count + STRESS_CASES
    width = trialCurrent.Columns.Count
    pathWidth = pathCurrent.Columns.Count
    If trialResults.Rows.Count < total Then RaiseRunError "The trial result area is too small for this run."
    ReDim batch(1 To trialResults.Rows.Count, 1 To width)
    ReDim ids(1 To trialResults.Rows.Count, 1 To 1)
    ReDim paths(1 To 3, 1 To pathWidth)
    ReDim chosenIDs(1 To 3, 1 To 1)
    ReDim order(1 To count)
    started = Timer

    For index = 1 To total
        code = TrialCode(index, count)
        pairedRow = 0
        If code >= 1001 And code <= 2000 Then pairedRow = index - count
        If code >= 2101 And code <= 2113 Then pairedRow = index - 13
        captured = CaptureTrial(book, selector, trialCurrent, code, context, batch, pairedRow, index - 1)
        For column = 1 To width
            batch(index, column) = captured(1, column)
        Next column
        ids(index, 1) = code
        If index <= count Then order(index) = index
        ShowProgress index, total + 3, started
    Next index

    ' Rank complete fixed-spending paths before touching any saved result.
    SortPaths order, batch, ids, 1, count
    For caseIndex = 1 To 3
        Select Case caseIndex
            Case 1: rank = 1 + Int(0.1 * (count - 1) + 0.5)
            Case 2: rank = 1 + Int(0.5 * (count - 1) + 0.5)
            Case 3: rank = 1 + Int(0.9 * (count - 1) + 0.5)
        End Select
        code = CLng(ids(order(rank), 1))
        captured = CaptureTrial(book, selector, pathCurrent, code, context, batch, order(rank), 0)
        For column = 1 To pathWidth
            paths(caseIndex, column) = captured(1, column)
        Next column
        chosenIDs(caseIndex, 1) = code
        ShowProgress total + caseIndex, total + 3, started
    Next caseIndex
    CheckCancelled
    selector.Value2 = previousSelector
    selectorRestored = True
    Application.Calculate
    RequireText AssumptionRecord(book), context, "The assumptions changed during the run. Run market risk again."

    ' Keep backups until the report confirms the complete replacement batch.
    oldBatch = trialResults.Value2
    oldIDs = trialIDs.Value2
    oldPaths = pathResults.Value2
    oldPathIDs = pathIDs.Value2
    oldMetadata = results.Range("Z6:Z11").Value2
    ReDim metadata(1 To 6, 1 To 1)
    metadata(1, 1) = "complete"
    metadata(2, 1) = CDbl(Now)
    If book.Date1904 Then metadata(2, 1) = metadata(2, 1) - 1462
    metadata(3, 1) = context
    metadata(4, 1) = "Updated " & Format$(count, "#,##0") & " paired paths and 33 stress cases."
    metadata(5, 1) = count
    metadata(6, 1) = seed

    ' No user events occur between these bounded writes and verification.
    Application.EnableCancelKey = xlDisabled
    publishing = True
    trialResults.Value2 = batch
    trialIDs.Value2 = ids
    pathResults.Value2 = paths
    pathIDs.Value2 = chosenIDs
    results.Range("Z6:Z11").Value2 = metadata
    Application.Calculate
    RequireText results.Range("Z1").Value2, "current", "The completed run did not pass the workbook's result checks."
    completed = True
    publishing = False
    GoTo CleanExit

RunFailed:
    failureNumber = Err.Number
    If failureNumber = 18 Then
        failureText = "Run cancelled. No new results were published."
    Else
        failureText = "Run stopped: " & Err.Description
    End If
    Resume CleanExit

CleanExit:
    ' Every cleanup operation reports its own failure. Do not leave a healthy
    ' status when publication or restoration could not be completed.
    On Error Resume Next
    If savedSettings Then
        Application.EnableCancelKey = xlDisabled
        If Err.Number <> 0 Then cleanupText = "Excel's cancellation setting could not be restored safely.": Err.Clear
    End If
    If publishing Then
        note = RestoreResults(trialResults, oldBatch, trialIDs, oldIDs, pathResults, oldPaths, pathIDs, oldPathIDs, results, oldMetadata)
        AddNote cleanupText, note
    End If
    If savedSelector And selectorTouched And Not selectorRestored Then
        selector.Value2 = previousSelector
        If Err.Number <> 0 Then AddNote cleanupText, "The inspection selector could not be restored.": Err.Clear
    End If
    If Not completed Or Len(cleanupText) > 0 Then
        If Len(failureText) = 0 Then failureText = "Run stopped during cleanup."
        If Len(cleanupText) > 0 Then failureText = failureText & " " & cleanupText
        If runBegan And Not results Is Nothing Then
            results.Range("Z6").Value2 = IIf(failureNumber = 18, "cancelled", "failed")
            results.Range("Z9").Value2 = failureText
            If Err.Number <> 0 Then AddNote cleanupText, "The run status could not be written. Rerun market risk before using its report.": Err.Clear
        End If
    End If
    If savedSettings Then
        Application.Calculate
        If Err.Number <> 0 Then AddNote cleanupText, "The restored workbook could not be recalculated.": Err.Clear
        note = RestoreExcelState(settings)
        AddNote cleanupText, note
    End If
    If preparedTrials Then
        note = RestoreProtection(trials, trialProtection)
        AddNote cleanupText, note
    End If
    If preparedPaths Then
        note = RestoreProtection(pathSheet, pathProtection)
        AddNote cleanupText, note
    End If
    If runBegan And Len(cleanupText) > 0 And Not results Is Nothing Then
        results.Range("Z6").Value2 = "failed"
        results.Range("Z9").Value2 = "Review this run before using its report. " & cleanupText
        Application.Calculate
        If Err.Number <> 0 Then AddNote cleanupText, "The cleanup warning could not be written.": Err.Clear
    End If
    If preparedResults Then
        note = RestoreProtection(results, resultProtection)
        AddNote cleanupText, note
        If runBegan And Len(note) > 0 Then
            results.Range("Z6").Value2 = "failed"
            results.Range("Z9").Value2 = "Review this run before using its report. " & cleanupText
            Application.Calculate
            If Err.Number <> 0 Then AddNote cleanupText, "The protection warning could not be written.": Err.Clear
        End If
    End If
    If savedSettings Then Application.EnableCancelKey = settings.CancelKey
    If Err.Number <> 0 Then AddNote cleanupText, "Excel's cancellation setting needs review.": Err.Clear
    mCancelRequested = False
    mRunning = False
    On Error GoTo 0
    If Not completed Or Len(cleanupText) > 0 Then
        If Len(failureText) = 0 Then failureText = "The run completed, but Excel's prior settings could not all be restored."
        If Len(cleanupText) > 0 And InStr(1, failureText, cleanupText, vbBinaryCompare) = 0 Then failureText = failureText & " " & cleanupText
        MsgBox failureText, vbExclamation, "Market risk"
    End If
End Sub

Private Function CaptureTrial(ByVal book As Workbook, ByVal selector As Range, ByVal source As Range, ByVal code As Long, ByVal context As String, ByRef batch As Variant, ByVal pairedRow As Long, ByVal previousRow As Long) As Variant
    Dim values As Variant, column As Long, attempt As Long, primary As Boolean
    CheckCancelled
    selector.Value2 = code
    primary = (source.Parent Is selector.Parent)
    For attempt = 1 To 2
        ' Complete the workbook dependency chain before refreshing its terminal output row.
        FinishTrialCalculation (attempt = 2)
        source.Calculate
        CheckCancelled
        RequireText AssumptionRecord(book), context, "The assumptions changed during the run. Run market risk again."
        values = source.Value2
        If CaptureBalancesAgree(book, values, primary, batch, pairedRow) Then
            If attempt = 2 Or Not primary Then Exit For
            If Not RepeatsPriorEnding(values, batch, previousRow) Then Exit For
        End If
    Next attempt
    If attempt > 2 Then RaiseRunError "Trial " & CStr(code) & " did not reconcile after a full recalculation. No results were published."
    For column = 1 To source.Columns.Count
        If IsError(values(1, column)) Then RaiseRunError "Trial " & CStr(code) & " contains a calculation error. Inspect Trial_Checks and the trial schedules."
    Next column
    RequireText values(1, 1), context, "A trial's assumptions do not match this run."
    RequireText values(1, source.Columns.Count), context, "A trial's closing assumption record does not match this run."
    If Not IsNumeric(values(1, VALID_FIELD)) Then RaiseRunError "Trial " & CStr(code) & " has no valid calculation status."
    If CDbl(values(1, VALID_FIELD)) <> 1 Then RaiseRunError "Trial " & CStr(code) & " did not pass its calculation checks. Inspect Trial_Checks and healthcare inputs."
    CaptureTrial = values
End Function

Private Sub FinishTrialCalculation(ByVal full As Boolean)
    Dim started As Double, elapsed As Double
    started = Timer
    If full Then
        Application.CalculateFull
    Else
        Application.Calculate
    End If
    ' In manual mode xlPending can remain set for unrelated dirty cells.
    ' The synchronous calculation above has returned. Wait only while Excel is
    ' actively calculating, then verify our own ledger and terminal row.
    Do While Application.CalculationState = xlCalculating
        DoEvents
        CheckCancelled
        elapsed = Timer - started
        If elapsed < 0 Then elapsed = elapsed + 86400
        If elapsed > 120 Then RaiseRunError "Excel did not finish calculating the trial. No results were published."
    Loop
End Sub

Private Function CaptureBalancesAgree(ByVal book As Workbook, ByRef values As Variant, ByVal primary As Boolean, ByRef batch As Variant, ByVal pairedRow As Long) As Boolean
    Dim ledger As Variant, years As Long, period As Long, field As Long
    years = WholeNumber(book.Worksheets("Trial_Inputs").Range("B70").Value2, 2, ANNUAL_CAPACITY, "The trial horizon is invalid.")
    ledger = book.Worksheets("Trial_Projection").Range("BH5").Resize(years, 7).Value2
    If primary Then
        ' BH is the closing portfolio; BN is the same balance in opening-year dollars.
        For period = 1 To years
            field = FIRST_YEAR_FIELD + period - 1
            If Not SameCaptureValue(values(1, field), ledger(period, 7)) Then Exit Function
            If Not SameCaptureValue(values(1, field + ANNUAL_CAPACITY), ledger(period, 1)) Then Exit Function
        Next period
        If Not SameCaptureValue(values(1, ENDING_REAL_FIELD), ledger(years, 7)) Then Exit Function
        If Not SameCaptureValue(values(1, ENDING_NOMINAL_FIELD), ledger(years, 1)) Then Exit Function
        If pairedRow > 0 And SameCaptureValue(values(1, CUT_YEARS_FIELD), 0#) Then
            ' Identical returns and no spending cuts must reproduce the complete fixed path.
            For field = VALID_FIELD To UBound(values, 2) - 1
                If field <> REFRESH_DATE_FIELD Then
                    If Not SameCaptureValue(values(1, field), batch(pairedRow, field)) Then Exit Function
                End If
            Next field
        End If
    Else
        ' Each representative funding ledger must reproduce its already ranked fixed path.
        For period = 1 To years
            field = 2 + FUNDING_BALANCE_FIELD + (period - 1) * FUNDING_YEAR_WIDTH
            If Not SameCaptureValue(values(1, field), ledger(period, 1)) Then Exit Function
            If Not SameCaptureValue(values(1, field), batch(pairedRow, FIRST_YEAR_FIELD + ANNUAL_CAPACITY + period - 1)) Then Exit Function
        Next period
    End If
    CaptureBalancesAgree = True
End Function

Private Function RepeatsPriorEnding(ByRef values As Variant, ByRef batch As Variant, ByVal previousRow As Long) As Boolean
    ' A nonzero identical ending after a different trajectory can be genuine.
    ' Verify it once with a full calculation before publishing it.
    Dim field As Long
    If previousRow < 1 Then Exit Function
    If Not IsNumeric(values(1, ENDING_REAL_FIELD)) Then Exit Function
    If values(1, ENDING_REAL_FIELD) <= 0 Then Exit Function
    If Not SameCaptureValue(values(1, ENDING_REAL_FIELD), batch(previousRow, ENDING_REAL_FIELD)) Then Exit Function
    If Not SameCaptureValue(values(1, ENDING_NOMINAL_FIELD), batch(previousRow, ENDING_NOMINAL_FIELD)) Then Exit Function
    For field = FIRST_YEAR_FIELD To FIRST_YEAR_FIELD + 2 * ANNUAL_CAPACITY - 1
        If Not SameCaptureValue(values(1, field), batch(previousRow, field)) Then
            RepeatsPriorEnding = True
            Exit Function
        End If
    Next field
End Function

Private Function SameCaptureValue(ByVal first As Variant, ByVal second As Variant) As Boolean
    If IsError(first) Or IsError(second) Then Exit Function
    If VarType(first) <> VarType(second) Then Exit Function
    SameCaptureValue = (first = second)
End Function

Private Function TrialCode(ByVal index As Long, ByVal count As Long) As Long
    Dim stress As Long
    If index <= count Then
        TrialCode = index
    ElseIf index <= 2 * count Then
        TrialCode = 1000 + index - count
    Else
        stress = index - 2 * count
        If stress <= 13 Then
            TrialCode = 2000 + stress
        ElseIf stress <= 26 Then
            TrialCode = 2100 + stress - 13
        Else
            TrialCode = 2200 + stress - 26
        End If
    End If
End Function

Private Sub ShowProgress(ByVal index As Long, ByVal total As Long, ByVal started As Double)
    Dim elapsed As Double
    elapsed = Timer - started
    If elapsed < 0 Then elapsed = elapsed + 86400
    Application.StatusBar = "Market risk: " & Format$(index, "#,##0") & " of " & Format$(total, "#,##0") & _
                            " calculations, " & Format$(elapsed, "0") & " sec. Use Cancel or press Esc to stop."
    DoEvents
    CheckCancelled
End Sub

Private Sub CheckCancelled()
    If mCancelRequested Then Err.Raise 18, "MarketRiskRunner", "Run cancelled."
End Sub

Private Sub SortPaths(ByRef order() As Long, ByRef data As Variant, ByRef ids As Variant, ByVal low As Long, ByVal high As Long)
    Dim left As Long, right As Long, pivot As Long, saved As Long
    If low >= high Then Exit Sub
    left = low
    right = high
    pivot = order((low + high) \ 2)
    Do While left <= right
        Do While ComparePaths(order(left), pivot, data, ids) < 0
            left = left + 1
        Loop
        Do While ComparePaths(order(right), pivot, data, ids) > 0
            right = right - 1
        Loop
        If left <= right Then
            saved = order(left): order(left) = order(right): order(right) = saved
            left = left + 1: right = right - 1
        End If
    Loop
    If low < right Then SortPaths order, data, ids, low, right
    If left < high Then SortPaths order, data, ids, left, high
End Sub

Private Function ComparePaths(ByVal a As Long, ByVal b As Long, ByRef data As Variant, ByRef ids As Variant) As Long
    Dim difference As Double
    If Not IsNumeric(data(a, ENDING_REAL_FIELD)) Or Not IsNumeric(data(b, ENDING_REAL_FIELD)) Then RaiseRunError "A fixed path has no numeric ending balance."
    difference = CDbl(data(a, ENDING_REAL_FIELD)) - CDbl(data(b, ENDING_REAL_FIELD))
    If difference = 0 Then difference = FailureAge(data(a, FAILURE_AGE_FIELD)) - FailureAge(data(b, FAILURE_AGE_FIELD))
    If difference = 0 Then
        If Not IsNumeric(data(a, MAX_UNPAID_FIELD)) Or Not IsNumeric(data(b, MAX_UNPAID_FIELD)) Then RaiseRunError "A fixed path has no numeric unpaid-cost measure."
        difference = CDbl(data(b, MAX_UNPAID_FIELD)) - CDbl(data(a, MAX_UNPAID_FIELD))
    End If
    If difference = 0 Then difference = CLng(ids(a, 1)) - CLng(ids(b, 1))
    If difference < 0 Then
        ComparePaths = -1
    ElseIf difference > 0 Then
        ComparePaths = 1
    Else
        ComparePaths = 0
    End If
End Function

Private Function FailureAge(ByVal value As Variant) As Double
    If IsNumeric(value) And Len(CStr(value)) > 0 Then FailureAge = CDbl(value) Else FailureAge = 1000000000#
End Function

Private Function LocalRange(ByVal sheet As Worksheet, ByVal name As String) As Range
    Dim area As Range, definition As Excel.Name, owner As Workbook, shortName As String
    For Each definition In sheet.Names
        shortName = Mid$(definition.Name, InStrRev(definition.Name, "!") + 1)
        If StrComp(shortName, name, vbTextCompare) = 0 Then
            Set area = definition.RefersToRange
            Set owner = area.Parent.Parent
            If Not owner Is ThisWorkbook Then RaiseRunError "A risk output name points to another workbook."
            If area.Areas.Count <> 1 Then RaiseRunError "Risk output names must refer to one continuous area."
            Set LocalRange = area
            Exit Function
        End If
    Next definition
    RaiseRunError "A required risk output name is missing: " & name & ". Use the current workbook release."
End Function

Private Sub ValidateRanges(ByVal current As Range, ByVal results As Range, ByVal ids As Range, ByVal pathCurrent As Range, ByVal paths As Range, ByVal pathIDs As Range)
    Dim areas(0 To 5) As Range, left As Long, right As Long
    Dim leftSheet As Worksheet, rightSheet As Worksheet
    If current.Rows.Count <> 1 Or current.Columns.Count < 32 Then RaiseRunError "The trial output vector has an unsupported layout."
    If results.Columns.Count <> current.Columns.Count Or ids.Columns.Count <> 1 Or ids.Rows.Count <> results.Rows.Count Then RaiseRunError "Trial output and result areas have different dimensions."
    If results.Rows.Count <> 2 * MAX_PAIRED_PATHS + STRESS_CASES Then RaiseRunError "The trial result capacity does not match the runner."
    If pathCurrent.Rows.Count <> 1 Or pathCurrent.Columns.Count < 3 Then RaiseRunError "The funding-path output vector has an unsupported layout."
    If paths.Rows.Count <> 3 Or paths.Columns.Count <> pathCurrent.Columns.Count Or pathIDs.Rows.Count <> 3 Or pathIDs.Columns.Count <> 1 Then RaiseRunError "The funding-path result areas have different dimensions."
    RequireNoFormulas results
    RequireNoFormulas ids
    RequireNoFormulas paths
    RequireNoFormulas pathIDs
    Set areas(0) = current: Set areas(1) = results: Set areas(2) = ids
    Set areas(3) = pathCurrent: Set areas(4) = paths: Set areas(5) = pathIDs
    Set leftSheet = paths.Parent: Set rightSheet = pathIDs.Parent
    If Not leftSheet Is rightSheet Then RaiseRunError "Funding-path identifiers and results must share a worksheet."
    Set leftSheet = current.Parent: Set rightSheet = results.Parent
    If Not leftSheet Is rightSheet Then RaiseRunError "Trial outputs and results must share a worksheet."
    Set leftSheet = results.Parent: Set rightSheet = ids.Parent
    If Not leftSheet Is rightSheet Then RaiseRunError "Trial identifiers and results must share a worksheet."
    For left = LBound(areas) To UBound(areas) - 1
        For right = left + 1 To UBound(areas)
            Set leftSheet = areas(left).Parent: Set rightSheet = areas(right).Parent
            If leftSheet Is rightSheet Then
                If Not Application.Intersect(areas(left), areas(right)) Is Nothing Then RaiseRunError "Risk input and result areas overlap."
            End If
        Next right
    Next left
End Sub

Private Sub RequireNoFormulas(ByVal area As Range)
    Dim hasFormula As Variant
    hasFormula = area.HasFormula
    If IsNull(hasFormula) Then RaiseRunError "A result area overlaps calculation formulas. Use the current workbook release."
    If CBool(hasFormula) Then RaiseRunError "A result area overlaps calculation formulas. Use the current workbook release."
End Sub

Private Function WholeNumber(ByVal value As Variant, ByVal low As Long, ByVal high As Long, ByVal message As String) As Long
    If IsError(value) Then RaiseRunError message
    If Not IsNumeric(value) Or VarType(value) = vbBoolean Then RaiseRunError message
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < low Or CDbl(value) > high Then RaiseRunError message
    WholeNumber = CLng(value)
End Function

Private Function SeedNumber(ByVal value As Variant) As Double
    If IsError(value) Then RaiseRunError "Enter a whole random seed from 0 to 4,294,967,295."
    If Not IsNumeric(value) Or VarType(value) = vbBoolean Then RaiseRunError "Enter a whole random seed from 0 to 4,294,967,295."
    If CDbl(value) <> Fix(CDbl(value)) Or CDbl(value) < 0 Or CDbl(value) > 4294967295# Then RaiseRunError "Enter a whole random seed from 0 to 4,294,967,295."
    SeedNumber = CDbl(value)
End Function

Private Function AssumptionRecord(ByVal book As Workbook) As String
    Dim value As Variant
    value = book.Worksheets("Market_Paths").Range("BB3").Value2
    If IsError(value) Then RaiseRunError "The assumption record contains a calculation error."
    If VarType(value) <> vbString Then RaiseRunError "The assumption record is missing."
    If Len(value) = 0 Or Len(value) > 32767 Then RaiseRunError "The assumption record has an unsupported length."
    AssumptionRecord = CStr(value)
End Function

Private Sub RequireText(ByVal actual As Variant, ByVal expected As String, ByVal message As String)
    If IsError(actual) Then RaiseRunError message
    If VarType(actual) <> vbString Then RaiseRunError message
    If StrComp(CStr(actual), expected, vbBinaryCompare) <> 0 Then RaiseRunError message
End Sub

Private Sub RaiseRunError(ByVal message As String)
    Err.Raise RUN_ERROR, "MarketRiskRunner", message
End Sub

Private Sub SaveExcelState(ByRef state As ExcelState)
    state.Calculation = Application.Calculation
    state.EventsEnabled = Application.EnableEvents
    state.ScreenUpdating = Application.ScreenUpdating
    state.CancelKey = Application.EnableCancelKey
    state.StatusVisible = Application.DisplayStatusBar
    state.StatusText = Application.StatusBar
    state.Cursor = Application.Cursor
End Sub

Private Function RestoreExcelState(ByRef state As ExcelState) As String
    Dim notes As String
    On Error Resume Next
    Application.Calculation = state.Calculation
    If Err.Number <> 0 Then AddNote notes, "The prior calculation mode could not be restored.": Err.Clear
    Application.Cursor = state.Cursor
    If Err.Number <> 0 Then AddNote notes, "The cursor could not be restored.": Err.Clear
    If VarType(state.StatusText) = vbBoolean Then
        ' False gives the status bar back to Excel after it displayed progress.
        Application.StatusBar = False
    Else
        Application.StatusBar = state.StatusText
    End If
    If Err.Number <> 0 Then AddNote notes, "The prior status bar text could not be restored.": Err.Clear
    Application.DisplayStatusBar = state.StatusVisible
    If Err.Number <> 0 Then AddNote notes, "The status bar visibility could not be restored.": Err.Clear
    Application.ScreenUpdating = state.ScreenUpdating
    If Err.Number <> 0 Then AddNote notes, "Screen updating could not be restored.": Err.Clear
    Application.EnableEvents = state.EventsEnabled
    If Err.Number <> 0 Then AddNote notes, "Excel events could not be restored.": Err.Clear
    On Error GoTo 0
    RestoreExcelState = notes
End Function

Private Function RestoreResults(ByVal results As Range, ByRef oldResults As Variant, ByVal ids As Range, ByRef oldIDs As Variant, ByVal paths As Range, ByRef oldPaths As Variant, ByVal pathIDs As Range, ByRef oldPathIDs As Variant, ByVal sheet As Worksheet, ByRef oldMetadata As Variant) As String
    Dim notes As String
    On Error Resume Next
    results.Value2 = oldResults
    If Err.Number <> 0 Then AddNote notes, "The prior trial results could not be restored.": Err.Clear
    ids.Value2 = oldIDs
    If Err.Number <> 0 Then AddNote notes, "The prior trial identifiers could not be restored.": Err.Clear
    paths.Value2 = oldPaths
    If Err.Number <> 0 Then AddNote notes, "The prior funding paths could not be restored.": Err.Clear
    pathIDs.Value2 = oldPathIDs
    If Err.Number <> 0 Then AddNote notes, "The prior funding-path identifiers could not be restored.": Err.Clear
    sheet.Range("Z6:Z11").Value2 = oldMetadata
    If Err.Number <> 0 Then AddNote notes, "The prior run record could not be restored.": Err.Clear
    On Error GoTo 0
    RestoreResults = notes
End Function

Private Sub PrepareProtection(ByVal sheet As Worksheet, ByRef state As Variant, ByRef prepared As Boolean)
    state = Array(sheet.ProtectDrawingObjects, sheet.ProtectContents, sheet.ProtectScenarios, _
        sheet.ProtectionMode, sheet.EnableSelection, sheet.Protection.AllowFormattingCells, _
        sheet.Protection.AllowFormattingColumns, sheet.Protection.AllowFormattingRows, _
        sheet.Protection.AllowInsertingColumns, sheet.Protection.AllowInsertingRows, _
        sheet.Protection.AllowInsertingHyperlinks, sheet.Protection.AllowDeletingColumns, _
        sheet.Protection.AllowDeletingRows, sheet.Protection.AllowSorting, _
        sheet.Protection.AllowFiltering, sheet.Protection.AllowUsingPivotTables)
    If state(0) Or state(1) Or state(2) Then
        sheet.Unprotect Password:=""
        prepared = True
        ApplyProtection sheet, state, True
    End If
End Sub

Private Sub ApplyProtection(ByVal sheet As Worksheet, ByRef state As Variant, ByVal macrosOnly As Boolean)
    sheet.Protect Password:="", DrawingObjects:=state(0), Contents:=state(1), _
        Scenarios:=state(2), UserInterfaceOnly:=macrosOnly, AllowFormattingCells:=state(5), _
        AllowFormattingColumns:=state(6), AllowFormattingRows:=state(7), _
        AllowInsertingColumns:=state(8), AllowInsertingRows:=state(9), _
        AllowInsertingHyperlinks:=state(10), AllowDeletingColumns:=state(11), _
        AllowDeletingRows:=state(12), AllowSorting:=state(13), AllowFiltering:=state(14), _
        AllowUsingPivotTables:=state(15)
    sheet.EnableSelection = state(4)
End Sub

Private Function RestoreProtection(ByVal sheet As Worksheet, ByRef state As Variant) As String
    On Error GoTo CouldNotRestore
    If state(0) Or state(1) Or state(2) Then
        sheet.Unprotect Password:=""
        ApplyProtection sheet, state, CBool(state(3))
    End If
    Exit Function
CouldNotRestore:
    RestoreProtection = "The prior protection settings for " & sheet.Name & " could not be restored."
End Function

Private Sub AddNote(ByRef text As String, ByVal note As String)
    If Len(note) = 0 Then Exit Sub
    If Len(text) > 0 Then text = text & " "
    text = text & note
End Sub
