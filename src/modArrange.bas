Attribute VB_Name = "modArrange"
Option Explicit

' ============================================================================
' modArrange — Tier 1 distribution/grid + Tier 2 stack (zero-gap abut).
' Sorting is positional (left-to-right or top-to-bottom), so visual order is
' preserved regardless of selection order — except GridArrange, which places
' shapes in SELECTION order (documented in its screentip).
' ============================================================================

' Returns the indices of sr sorted by Left (byLeft=True) or Top (insertion sort;
' selections are small).
Private Function SortedIdx(ByVal sr As ShapeRange, ByVal byLeft As Boolean) As Long()
    Dim idx() As Long, i As Long, j As Long, tmp As Long
    Dim ki As Double, kj As Double
    ReDim idx(1 To sr.Count)
    For i = 1 To sr.Count
        idx(i) = i
    Next i
    For i = 2 To sr.Count
        For j = i To 2 Step -1
            If byLeft Then
                ki = sr.Item(idx(j)).Left: kj = sr.Item(idx(j - 1)).Left
            Else
                ki = sr.Item(idx(j)).Top: kj = sr.Item(idx(j - 1)).Top
            End If
            If ki < kj Then
                tmp = idx(j): idx(j) = idx(j - 1): idx(j - 1) = tmp
            Else
                Exit For
            End If
        Next j
    Next i
    SortedIdx = idx
End Function

' Core placement: order shapes along one axis, first shape stays put, each
' following shape starts gap points after the previous one ends.
' gap = NUM_CANCELLED means "equal gaps": keep the outer bounds, spread evenly.
Private Sub Distribute(ByVal sr As ShapeRange, ByVal horizontal As Boolean, ByVal gap As Double)
    Dim idx() As Long, i As Long, n As Long
    Dim sumSize As Double, span As Double, pos As Double
    n = sr.Count
    idx = SortedIdx(sr, horizontal)
    If gap = NUM_CANCELLED Then
        For i = 1 To n
            sumSize = sumSize + IIf(horizontal, sr.Item(idx(i)).Width, sr.Item(idx(i)).Height)
        Next i
        If horizontal Then
            span = sr.Item(idx(n)).Left + sr.Item(idx(n)).Width - sr.Item(idx(1)).Left
        Else
            span = sr.Item(idx(n)).Top + sr.Item(idx(n)).Height - sr.Item(idx(1)).Top
        End If
        gap = (span - sumSize) / (n - 1)   ' may be negative: shapes overlap evenly
    End If
    If horizontal Then
        pos = sr.Item(idx(1)).Left
        For i = 1 To n
            sr.Item(idx(i)).Left = pos
            pos = pos + sr.Item(idx(i)).Width + gap
        Next i
    Else
        pos = sr.Item(idx(1)).Top
        For i = 1 To n
            sr.Item(idx(i)).Top = pos
            pos = pos + sr.Item(idx(i)).Height + gap
        Next i
    End If
End Sub

' Tag-dispatched: "H" | "V" — equal gaps, outer shapes keep their edges.
Public Sub DistributeEqual(control As Object)
    Const OP As String = "Distribute with equal gaps"
    On Error GoTo Oops
    Dim sr As ShapeRange
    Set sr = GuardShapes(3, OP)
    If sr Is Nothing Then Exit Sub
    Distribute sr, (control.Tag = "H"), NUM_CANCELLED
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "H" | "V" — fixed gap in cm, first shape anchors the row.
Public Sub DistributeFixedGap(control As Object)
    Const OP As String = "Distribute with fixed gap"
    On Error GoTo Oops
    Dim sr As ShapeRange, gapCm As Double
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    gapCm = AskNumber("Gap between shapes in cm:", "0.20")
    If gapCm = NUM_CANCELLED Then Exit Sub
    Distribute sr, (control.Tag = "H"), CmToPt(gapCm)
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "H" | "V" — Tier 2 stack/abut with zero gap.
Public Sub StackShapes(control As Object)
    Const OP As String = "Stack shapes (zero gap)"
    On Error GoTo Oops
    Dim sr As ShapeRange
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    Distribute sr, (control.Tag = "H"), 0#
    Exit Sub
Oops:
    ReportError OP
End Sub

' Arrange the selection into a grid of N columns. Shapes are placed in
' SELECTION order; the reference (first-selected) shape's top-left corner is
' the grid origin; every cell is as big as the largest shape plus the gap.
Public Sub GridArrange(control As Object)
    Const OP As String = "Arrange in grid"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, i As Long
    Dim cols As Double, gapCm As Double, nCols As Long
    Dim cellW As Double, cellH As Double, gapPt As Double
    Dim originX As Double, originY As Double, r As Long, c As Long
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    cols = AskNumber("Number of columns:", "3")
    If cols = NUM_CANCELLED Then Exit Sub
    nCols = CLng(cols)
    If nCols < 1 Or nCols > 50 Then
        MsgBox "Columns must be between 1 and 50.", vbExclamation, APP_NAME
        Exit Sub
    End If
    gapCm = AskNumber("Gap between cells in cm:", "0.20")
    If gapCm = NUM_CANCELLED Then Exit Sub
    gapPt = CmToPt(gapCm)
    For i = 1 To sr.Count
        If sr.Item(i).Width > cellW Then cellW = sr.Item(i).Width
        If sr.Item(i).Height > cellH Then cellH = sr.Item(i).Height
    Next i
    originX = RefShape(sr).Left
    originY = RefShape(sr).Top
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        r = (i - 1) \ nCols
        c = (i - 1) Mod nCols
        shp.Left = originX + c * (cellW + gapPt)
        shp.Top = originY + r * (cellH + gapPt)
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub
