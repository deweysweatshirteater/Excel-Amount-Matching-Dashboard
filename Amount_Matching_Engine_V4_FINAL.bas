Attribute VB_Name = "AmountMatchingEngineV4"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Matching Dashboard"
Private Const BUTTON_NAME As String = "btnRunAmountMatcherV4"
Private Const DEFAULT_VARIANCE As Currency = 150000@
Private Const DEFAULT_MAX_RESULTS As Long = 25
Private Const DEFAULT_MAX_ITEMS As Long = 8
Private Const DEFAULT_SEARCH_LIMIT As Long = 5000000
Private Const DEFAULT_SOLUTIONS As Long = 5
Private Const DEFAULT_MAX_USES As Long = 1

Private mAmounts() As Currency
Private mSourceRows() As Long
Private mUseLimit() As Long
Private mUsedCount() As Long
Private mPick() As Long

Private mSolutionTotal() As Currency
Private mSolutionVariance() As Currency
Private mSolutionItemCount() As Long
Private mSolutionPick() As Long
Private mSolutionCount As Long
Private mTopSolutionLimit As Long

Private mTarget As Currency
Private mAllowance As Currency
Private mVisits As Long
Private mSearchLimit As Long
Private mLimitReached As Boolean

Public Sub InstallAmountMatcherV4()
    Dim ws As Worksheet
    On Error GoTo InstallFail

    Set ws = GetOrCreateDashboard()
    PrepareDashboard ws
    AddRunButton ws
    MsgBox "Version 4 is installed. Use the Run Amount Matcher button on the dashboard.", vbInformation
    Exit Sub

InstallFail:
    MsgBox "Version 4 could not be installed: " & Err.Description, vbCritical
End Sub

Public Sub AnalyzeMatches()
    Dim ws As Worksheet
    Dim lastAvailable As Long, lastTarget As Long
    Dim amountCount As Long, targetRow As Long, resultRow As Long
    Dim maxResults As Long, maxItems As Long
    Dim processed As Long, solutionIndex As Long
    Dim oldResultLast As Long
    Dim targetValue As Variant, targetAmount As Currency

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    PrepareDashboard ws
    AddRunButton ws

    mAllowance = Abs(CCur(ws.Range("K2").Value2))
    maxResults = PositiveLong(ws.Range("K3").Value2, DEFAULT_MAX_RESULTS)
    maxItems = PositiveLong(ws.Range("K4").Value2, DEFAULT_MAX_ITEMS)
    mSearchLimit = PositiveLong(ws.Range("K5").Value2, DEFAULT_SEARCH_LIMIT)
    mTopSolutionLimit = PositiveLong(ws.Range("K6").Value2, DEFAULT_SOLUTIONS)
    If mTopSolutionLimit > 25 Then mTopSolutionLimit = 25

    lastAvailable = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastTarget = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastAvailable < 2 Or lastTarget < 2 Then
        MsgBox "Enter available amounts in column A and targets in column B.", vbExclamation
        Exit Sub
    End If

    LoadAmounts ws, lastAvailable, amountCount, PositiveLong(ws.Range("K7").Value2, DEFAULT_MAX_USES)
    If amountCount = 0 Then
        MsgBox "No numeric available amounts were found in column A.", vbExclamation
        Exit Sub
    End If
    If maxItems > amountCount Then maxItems = amountCount

    oldResultLast = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
    If oldResultLast >= 2 Then
        ws.Range("D2:I" & oldResultLast).ClearContents
        ws.Range("D2:I" & oldResultLast).Interior.Pattern = xlNone
        ws.Range("D2:I" & oldResultLast).Font.Bold = False
    End If
    ws.Range("L2:L" & ws.Rows.Count).ClearContents
    ws.Range("D1:I1").Value = Array("Target", "Rank", "Match Found", "Matched Total", "Variance", "Quality")
    resultRow = 2

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Analyzing amount matches..."

    For targetRow = 2 To lastTarget
        targetValue = ws.Cells(targetRow, "B").Value2
        If Not IsError(targetValue) Then
            If IsNumeric(targetValue) Then
                If CCur(targetValue) <> 0 Then
                    targetAmount = Abs(CCur(targetValue))
                    FindTopMatches targetAmount, amountCount, maxItems

                    If mSolutionCount = 0 Then
                        WriteNoMatch ws, resultRow, CCur(targetValue)
                        resultRow = resultRow + 1
                    Else
                        For solutionIndex = 1 To mSolutionCount
                            WriteSolution ws, resultRow, CCur(targetValue), solutionIndex
                            resultRow = resultRow + 1
                        Next solutionIndex
                        CommitPrimarySolution
                    End If

                    processed = processed + 1
                    If processed >= maxResults Then Exit For
                End If
            End If
        End If
    Next targetRow

    WriteUsageCounts ws, amountCount
    FormatResults ws, resultRow

CleanExit:
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    If Err.Number = 0 Then MsgBox CStr(processed) & " target(s) analyzed.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "The analysis stopped: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Function GetOrCreateDashboard() As Worksheet
    On Error Resume Next
    Set GetOrCreateDashboard = ThisWorkbook.Worksheets(DASHBOARD_SHEET)
    On Error GoTo 0

    If GetOrCreateDashboard Is Nothing Then
        Set GetOrCreateDashboard = ThisWorkbook.Worksheets.Add
        GetOrCreateDashboard.Name = DASHBOARD_SHEET
    End If
End Function

Private Sub PrepareDashboard(ByVal ws As Worksheet)
    ws.Range("A1").Value = "Available Amounts"
    ws.Range("B1").Value = "Amounts To Resolve"
    ws.Range("C1").Value = "Maximum Uses"
    ws.Range("D1:I1").Value = Array("Target", "Rank", "Match Found", "Matched Total", "Variance", "Quality")
    ws.Range("J1").Value = "Setting"
    ws.Range("K1").Value = "Value"
    ws.Range("L1").Value = "Uses Consumed"

    EnsureSetting ws, 2, "Variance allowance", DEFAULT_VARIANCE
    EnsureSetting ws, 3, "Maximum targets", DEFAULT_MAX_RESULTS
    EnsureSetting ws, 4, "Maximum transactions per match", DEFAULT_MAX_ITEMS
    EnsureSetting ws, 5, "Search limit per target", DEFAULT_SEARCH_LIMIT
    EnsureSetting ws, 6, "Solutions shown per target", DEFAULT_SOLUTIONS
    EnsureSetting ws, 7, "Default maximum uses", DEFAULT_MAX_USES

    With ws.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
    End With
    ws.Columns("A:B").ColumnWidth = 18
    ws.Columns("C:C").ColumnWidth = 15
    ws.Columns("D:D").ColumnWidth = 16
    ws.Columns("E:E").ColumnWidth = 8
    ws.Columns("F:F").ColumnWidth = 42
    ws.Columns("G:H").ColumnWidth = 16
    ws.Columns("I:I").ColumnWidth = 28
    ws.Columns("J:J").ColumnWidth = 32
    ws.Columns("K:K").ColumnWidth = 14
    ws.Columns("L:L").ColumnWidth = 16
    ws.Range("A:A,B:B,D:D,G:G,H:H").NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("C:C,E:E,K3:K7,L:L").NumberFormat = "0"
End Sub

Private Sub EnsureSetting(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal label As String, ByVal defaultValue As Variant)
    ws.Cells(rowNumber, "J").Value = label
    If Len(ws.Cells(rowNumber, "K").Value2) = 0 Then ws.Cells(rowNumber, "K").Value = defaultValue
    If Not IsNumeric(ws.Cells(rowNumber, "K").Value2) Then ws.Cells(rowNumber, "K").Value = defaultValue
End Sub

Private Sub AddRunButton(ByVal ws As Worksheet)
    Dim button As Shape
    Dim buttonLeft As Double, buttonTop As Double
    Dim buttonWidth As Double, buttonHeight As Double
    Dim macroName As String

    On Error Resume Next
    Set button = ws.Shapes(BUTTON_NAME)
    On Error GoTo 0

    If button Is Nothing Then
        buttonLeft = ws.Range("J9").Left
        buttonTop = ws.Range("J9").Top
        buttonWidth = ws.Range("J9:K10").Width
        buttonHeight = ws.Range("J9:K10").Height

        Set button = ws.Shapes.AddFormControl(xlButtonControl, buttonLeft, buttonTop, buttonWidth, buttonHeight)
        button.Name = BUTTON_NAME
        button.TextFrame.Characters.Text = "Run Amount Matcher"
    End If
    macroName = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AnalyzeMatches"
    button.OnAction = macroName
End Sub

Private Function PositiveLong(ByVal value As Variant, ByVal fallback As Long) As Long
    If IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) <= 2147483647# Then
            PositiveLong = CLng(value)
            Exit Function
        End If
    End If
    PositiveLong = fallback
End Function

Private Function NonNegativeLong(ByVal value As Variant, ByVal fallback As Long) As Long
    If IsNumeric(value) Then
        If CDbl(value) >= 0 And CDbl(value) <= 2147483647# Then
            NonNegativeLong = CLng(value)
            Exit Function
        End If
    End If
    NonNegativeLong = fallback
End Function

Private Sub LoadAmounts(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef count As Long, ByVal defaultMaxUses As Long)
    Dim r As Long, value As Variant, usesValue As Variant
    ReDim mAmounts(1 To lastRow - 1)
    ReDim mSourceRows(1 To lastRow - 1)
    ReDim mUseLimit(1 To lastRow - 1)
    ReDim mUsedCount(1 To lastRow - 1)

    For r = 2 To lastRow
        value = ws.Cells(r, "A").Value2
        If Not IsError(value) Then
            If IsNumeric(value) Then
                If CCur(value) <> 0 Then
                    count = count + 1
                    mAmounts(count) = CCur(value)
                    mSourceRows(count) = r
                    usesValue = ws.Cells(r, "C").Value2
                    If IsError(usesValue) Then
                        mUseLimit(count) = defaultMaxUses
                    ElseIf Len(usesValue) = 0 Then
                        mUseLimit(count) = defaultMaxUses
                    Else
                        mUseLimit(count) = NonNegativeLong(usesValue, defaultMaxUses)
                    End If
                End If
            End If
        End If
    Next r

    If count > 0 Then
        ReDim Preserve mAmounts(1 To count)
        ReDim Preserve mSourceRows(1 To count)
        ReDim Preserve mUseLimit(1 To count)
        ReDim Preserve mUsedCount(1 To count)
    End If
End Sub

Private Sub FindTopMatches(ByVal target As Currency, ByVal amountCount As Long, ByVal maxItems As Long)
    Dim size As Long
    mTarget = target
    mSolutionCount = 0
    mVisits = 0
    mLimitReached = False

    ReDim mPick(1 To maxItems)
    ReDim mSolutionTotal(1 To mTopSolutionLimit)
    ReDim mSolutionVariance(1 To mTopSolutionLimit)
    ReDim mSolutionItemCount(1 To mTopSolutionLimit)
    ReDim mSolutionPick(1 To mTopSolutionLimit, 1 To maxItems)

    For size = 1 To maxItems
        SearchCombinations 1, 1, size, 0, amountCount
        If mLimitReached Then Exit For
        If HaveEnoughExactSolutions(size) Then Exit For
    Next size
End Sub

Private Sub SearchCombinations(ByVal startIndex As Long, ByVal depth As Long, ByVal requiredCount As Long, ByVal runningTotal As Currency, ByVal amountCount As Long)
    Dim i As Long, newTotal As Currency, variance As Currency
    If mLimitReached Then Exit Sub

    For i = startIndex To amountCount - (requiredCount - depth)
        If mUsedCount(i) < mUseLimit(i) Then
            mPick(depth) = i
            newTotal = runningTotal + mAmounts(i)

            If depth = requiredCount Then
                mVisits = mVisits + 1
                variance = Abs(newTotal - mTarget)
                If variance <= mAllowance Then InsertCandidate depth, newTotal, variance
                If mVisits >= mSearchLimit Then
                    mLimitReached = True
                    Exit Sub
                End If
            Else
                SearchCombinations i + 1, depth + 1, requiredCount, newTotal, amountCount
                If mLimitReached Then Exit Sub
            End If
        End If
    Next i
End Sub

Private Sub InsertCandidate(ByVal itemCount As Long, ByVal total As Currency, ByVal variance As Currency)
    Dim position As Long, shiftIndex As Long, pickIndex As Long

    position = 1
    Do While position <= mSolutionCount
        If IsCandidateBetter(variance, itemCount, mSolutionVariance(position), mSolutionItemCount(position)) Then Exit Do
        position = position + 1
    Loop

    If mSolutionCount = mTopSolutionLimit And position > mSolutionCount Then Exit Sub
    If mSolutionCount < mTopSolutionLimit Then mSolutionCount = mSolutionCount + 1

    For shiftIndex = mSolutionCount To position + 1 Step -1
        mSolutionTotal(shiftIndex) = mSolutionTotal(shiftIndex - 1)
        mSolutionVariance(shiftIndex) = mSolutionVariance(shiftIndex - 1)
        mSolutionItemCount(shiftIndex) = mSolutionItemCount(shiftIndex - 1)
        For pickIndex = 1 To UBound(mSolutionPick, 2)
            mSolutionPick(shiftIndex, pickIndex) = mSolutionPick(shiftIndex - 1, pickIndex)
        Next pickIndex
    Next shiftIndex

    mSolutionTotal(position) = total
    mSolutionVariance(position) = variance
    mSolutionItemCount(position) = itemCount
    For pickIndex = 1 To itemCount
        mSolutionPick(position, pickIndex) = mPick(pickIndex)
    Next pickIndex
    For pickIndex = itemCount + 1 To UBound(mSolutionPick, 2)
        mSolutionPick(position, pickIndex) = 0
    Next pickIndex
End Sub

Private Function IsCandidateBetter(ByVal newVariance As Currency, ByVal newCount As Long, ByVal oldVariance As Currency, ByVal oldCount As Long) As Boolean
    If newVariance < oldVariance Then
        IsCandidateBetter = True
    ElseIf newVariance = oldVariance And newCount < oldCount Then
        IsCandidateBetter = True
    End If
End Function

Private Function HaveEnoughExactSolutions(ByVal currentSize As Long) As Boolean
    If mSolutionCount < mTopSolutionLimit Then Exit Function
    HaveEnoughExactSolutions = (mSolutionVariance(mTopSolutionLimit) = 0 And mSolutionItemCount(mTopSolutionLimit) = currentSize)
End Function

Private Sub WriteSolution(ByVal ws As Worksheet, ByVal resultRow As Long, ByVal originalTarget As Currency, ByVal solutionIndex As Long)
    Dim quality As String
    If mSolutionVariance(solutionIndex) = 0 Then
        quality = "EXACT"
    Else
        quality = "WITHIN ALLOWANCE"
    End If
    If mLimitReached Then quality = quality & " (SEARCH LIMITED)"

    ws.Cells(resultRow, "D").Value = originalTarget
    ws.Cells(resultRow, "E").Value = solutionIndex
    ws.Cells(resultRow, "F").Value = BuildMatchText(solutionIndex)
    ws.Cells(resultRow, "G").Value = mSolutionTotal(solutionIndex)
    ws.Cells(resultRow, "H").Value = mSolutionVariance(solutionIndex)
    ws.Cells(resultRow, "I").Value = quality

    If solutionIndex = 1 Then
        ws.Range("D" & resultRow & ":I" & resultRow).Interior.Color = RGB(226, 239, 218)
        ws.Range("D" & resultRow & ":I" & resultRow).Font.Bold = True
    Else
        ws.Range("D" & resultRow & ":I" & resultRow).Interior.Pattern = xlNone
        ws.Range("D" & resultRow & ":I" & resultRow).Font.Bold = False
    End If
End Sub

Private Sub WriteNoMatch(ByVal ws As Worksheet, ByVal resultRow As Long, ByVal originalTarget As Currency)
    ws.Cells(resultRow, "D").Value = originalTarget
    ws.Cells(resultRow, "F").Value = "No acceptable match found"
    If mLimitReached Then
        ws.Cells(resultRow, "I").Value = "NO MATCH (SEARCH LIMITED)"
    Else
        ws.Cells(resultRow, "I").Value = "NO MATCH"
    End If
    ws.Range("D" & resultRow & ":I" & resultRow).Interior.Pattern = xlNone
    ws.Range("D" & resultRow & ":I" & resultRow).Font.Bold = False
End Sub

Private Function BuildMatchText(ByVal solutionIndex As Long) As String
    Dim pickIndex As Long, sourceIndex As Long, text As String
    For pickIndex = 1 To mSolutionItemCount(solutionIndex)
        sourceIndex = mSolutionPick(solutionIndex, pickIndex)
        If pickIndex > 1 Then text = text & " + "
        text = text & Format$(mAmounts(sourceIndex), "#,##0.00") & " [A" & CStr(mSourceRows(sourceIndex)) & "]"
    Next pickIndex
    BuildMatchText = text
End Function

Private Sub CommitPrimarySolution()
    Dim pickIndex As Long, sourceIndex As Long
    For pickIndex = 1 To mSolutionItemCount(1)
        sourceIndex = mSolutionPick(1, pickIndex)
        mUsedCount(sourceIndex) = mUsedCount(sourceIndex) + 1
    Next pickIndex
End Sub

Private Sub WriteUsageCounts(ByVal ws As Worksheet, ByVal amountCount As Long)
    Dim i As Long
    For i = 1 To amountCount
        ws.Cells(mSourceRows(i), "L").Value = mUsedCount(i)
    Next i
End Sub

Private Sub FormatResults(ByVal ws As Worksheet, ByVal nextResultRow As Long)
    Dim lastResultRow As Long
    lastResultRow = nextResultRow - 1
    If lastResultRow < 2 Then Exit Sub

    ws.Range("D2:D" & lastResultRow).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("G2:H" & lastResultRow).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("D1:I" & lastResultRow).Borders.LineStyle = xlContinuous
    ws.Range("D1:I" & lastResultRow).Borders.Color = RGB(217, 217, 217)
End Sub
