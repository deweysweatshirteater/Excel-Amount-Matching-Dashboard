Attribute VB_Name = "AmountMatchingEngine"
Option Explicit

Private Const DASHBOARD_SHEET As String = "Matching Dashboard"
Private Const DEFAULT_VARIANCE As Currency = 150000@
Private Const DEFAULT_MAX_RESULTS As Long = 25
Private Const DEFAULT_MAX_ITEMS As Long = 8
Private Const DEFAULT_SEARCH_LIMIT As Long = 250000

Private mAmounts() As Currency
Private mRows() As Long
Private mPick() As Long
Private mBestPick() As Long
Private mBestCount As Long
Private mBestTotal As Currency
Private mBestVariance As Currency
Private mTarget As Currency
Private mAllowance As Currency
Private mVisits As Long
Private mSearchLimit As Long
Private mLimitReached As Boolean

Public Sub AnalyzeMatches()
    Dim ws As Worksheet
    Dim lastAvailable As Long, lastTarget As Long
    Dim amountCount As Long, targetRow As Long, resultRow As Long
    Dim maxResults As Long, maxItems As Long
    Dim processed As Long, found As Boolean
    Dim targetValue As Variant, targetAmount As Currency

    On Error GoTo CleanFail
    Set ws = ThisWorkbook.Worksheets(DASHBOARD_SHEET)

    EnsureSettings ws
    mAllowance = Abs(CCur(ws.Range("K2").Value2))
    maxResults = PositiveLong(ws.Range("K3").Value2, DEFAULT_MAX_RESULTS)
    maxItems = PositiveLong(ws.Range("K4").Value2, DEFAULT_MAX_ITEMS)
    mSearchLimit = PositiveLong(ws.Range("K5").Value2, DEFAULT_SEARCH_LIMIT)

    lastAvailable = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastTarget = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    If lastAvailable < 2 Or lastTarget < 2 Then
        MsgBox "Enter available amounts in column A and targets in column B.", vbExclamation
        Exit Sub
    End If

    LoadAmounts ws, lastAvailable, amountCount
    If amountCount = 0 Then
        MsgBox "No numeric available amounts were found in column A.", vbExclamation
        Exit Sub
    End If
    If maxItems > amountCount Then maxItems = amountCount

    ws.Range("D2:H" & ws.Rows.Count).ClearContents
    ws.Range("D1:H1").Value = Array("Target", "Match Found", "Matched Total", "Variance", "Quality")
    resultRow = 2

    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing amount matches..."

    For targetRow = 2 To lastTarget
        targetValue = ws.Cells(targetRow, "B").Value2
        If Not IsError(targetValue) And IsNumeric(targetValue) Then
            If CCur(targetValue) <> 0 Then
                targetAmount = Abs(CCur(targetValue))
                found = FindBestMatch(targetAmount, amountCount, maxItems)

                ws.Cells(resultRow, "D").Value = CCur(targetValue)
                If found Then
                    ws.Cells(resultRow, "E").Value = BuildMatchText()
                    ws.Cells(resultRow, "F").Value = mBestTotal
                    ws.Cells(resultRow, "G").Value = mBestVariance
                    If mBestVariance = 0 Then
                        ws.Cells(resultRow, "H").Value = "EXACT"
                    ElseIf mLimitReached Then
                        ws.Cells(resultRow, "H").Value = "WITHIN ALLOWANCE (SEARCH LIMITED)"
                    Else
                        ws.Cells(resultRow, "H").Value = "WITHIN ALLOWANCE"
                    End If
                Else
                    ws.Cells(resultRow, "E").Value = "No acceptable match found"
                    If mLimitReached Then
                        ws.Cells(resultRow, "H").Value = "NO MATCH (SEARCH LIMITED)"
                    Else
                        ws.Cells(resultRow, "H").Value = "NO MATCH"
                    End If
                End If

                resultRow = resultRow + 1
                processed = processed + 1
                If processed >= maxResults Then Exit For
            End If
        End If
    Next targetRow

    ws.Range("D2:D" & resultRow - 1).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("F2:G" & resultRow - 1).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Columns("D:H").AutoFit

CleanExit:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If Err.Number = 0 Then MsgBox CStr(processed) & " target(s) analyzed.", vbInformation
    Exit Sub

CleanFail:
    MsgBox "The analysis stopped: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub EnsureSettings(ByVal ws As Worksheet)
    If Len(ws.Range("J2").Value2) = 0 Then ws.Range("J2").Value = "Variance allowance"
    If Len(ws.Range("K2").Value2) = 0 Or Not IsNumeric(ws.Range("K2").Value2) Then ws.Range("K2").Value = DEFAULT_VARIANCE
    If Len(ws.Range("J3").Value2) = 0 Then ws.Range("J3").Value = "Maximum results"
    If Len(ws.Range("K3").Value2) = 0 Or Not IsNumeric(ws.Range("K3").Value2) Then ws.Range("K3").Value = DEFAULT_MAX_RESULTS
    If Len(ws.Range("J4").Value2) = 0 Then ws.Range("J4").Value = "Maximum transactions per match"
    If Len(ws.Range("K4").Value2) = 0 Or Not IsNumeric(ws.Range("K4").Value2) Then ws.Range("K4").Value = DEFAULT_MAX_ITEMS
    If Len(ws.Range("J5").Value2) = 0 Then ws.Range("J5").Value = "Search limit per target"
    If Len(ws.Range("K5").Value2) = 0 Or Not IsNumeric(ws.Range("K5").Value2) Then ws.Range("K5").Value = DEFAULT_SEARCH_LIMIT
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

Private Sub LoadAmounts(ByVal ws As Worksheet, ByVal lastRow As Long, ByRef count As Long)
    Dim r As Long, value As Variant
    ReDim mAmounts(1 To lastRow - 1)
    ReDim mRows(1 To lastRow - 1)

    For r = 2 To lastRow
        value = ws.Cells(r, "A").Value2
        If Not IsError(value) And IsNumeric(value) Then
            If CCur(value) <> 0 Then
                count = count + 1
                mAmounts(count) = CCur(value)
                mRows(count) = r
            End If
        End If
    Next r

    If count > 0 Then
        ReDim Preserve mAmounts(1 To count)
        ReDim Preserve mRows(1 To count)
    End If
End Sub

Private Function FindBestMatch(ByVal target As Currency, ByVal amountCount As Long, ByVal maxItems As Long) As Boolean
    Dim size As Long
    mTarget = target
    mBestCount = 0
    mBestTotal = 0
    mBestVariance = 922337203685477.5807@
    mVisits = 0
    mLimitReached = False
    ReDim mPick(1 To maxItems)
    ReDim mBestPick(1 To maxItems)

    For size = 1 To maxItems
        SearchCombinations 1, 1, size, 0, amountCount
        If mBestVariance = 0 Or mLimitReached Then Exit For
    Next size

    FindBestMatch = (mBestCount > 0 And mBestVariance <= mAllowance)
End Function

Private Sub SearchCombinations(ByVal startIndex As Long, ByVal depth As Long, ByVal requiredCount As Long, ByVal runningTotal As Currency, ByVal amountCount As Long)
    Dim i As Long, newTotal As Currency, variance As Currency
    If mLimitReached Then Exit Sub

    For i = startIndex To amountCount - (requiredCount - depth)
        mPick(depth) = i
        newTotal = runningTotal + mAmounts(i)

        If depth = requiredCount Then
            mVisits = mVisits + 1
            If mVisits >= mSearchLimit Then mLimitReached = True
            variance = Abs(newTotal - mTarget)
            If variance < mBestVariance Then SaveBest depth, newTotal, variance
            If mBestVariance = 0 Or mLimitReached Then Exit Sub
        Else
            SearchCombinations i + 1, depth + 1, requiredCount, newTotal, amountCount
            If mBestVariance = 0 Or mLimitReached Then Exit Sub
        End If
    Next i
End Sub

Private Sub SaveBest(ByVal count As Long, ByVal total As Currency, ByVal variance As Currency)
    Dim i As Long
    mBestCount = count
    mBestTotal = total
    mBestVariance = variance
    For i = 1 To count
        mBestPick(i) = mPick(i)
    Next i
End Sub

Private Function BuildMatchText() As String
    Dim i As Long, text As String
    For i = 1 To mBestCount
        If i > 1 Then text = text & " + "
        text = text & Format$(mAmounts(mBestPick(i)), "#,##0.00") & " [A" & CStr(mRows(mBestPick(i))) & "]"
    Next i
    BuildMatchText = text
End Function
