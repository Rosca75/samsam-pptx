Attribute VB_Name = "modGuides"
Option Explicit

' ============================================================================
' modGuides — Tier 3 guide tools. Uses the Presentation.Guides API
' (PowerPoint 2016+). Guide.Position is measured in points from the slide's
' top-left corner — VERIFY on first manual test (see testing checklist); if
' guides land mirrored around the centre, positions are centre-relative in
' this build and OffsetGuidePos below is the one place to fix it.
' ============================================================================

' Single adjustment point in case Guide.Position turns out centre-relative.
Private Function OffsetGuidePos(ByVal pos As Double, ByVal extent As Double) As Double
    OffsetGuidePos = pos          ' assume edge-relative (documented behaviour)
End Function

Private Sub AddGuide(ByVal vertical As Boolean, ByVal posPt As Double)
    Dim orient As Long, extent As Double
    orient = IIf(vertical, 2, 1)  ' ppVerticalGuide=2, ppHorizontalGuide=1
    extent = IIf(vertical, ActivePres.PageSetup.SlideWidth, _
                           ActivePres.PageSetup.SlideHeight)
    If posPt < 0 Or posPt > extent Then Exit Sub
    ActivePres.Guides.Add orient, OffsetGuidePos(posPt, extent)
End Sub

' Margin + column grid: vertical guides at the margins and between columns
' (gutter-aware), horizontal guides at top/bottom margins.
Public Sub BuildGuideGrid(control As Object)
    Const OP As String = "Build guide grid"
    On Error GoTo Oops
    Dim marginCm As Double, colsIn As Double, gutterCm As Double
    Dim nCols As Long, i As Long
    Dim sw As Double, margin As Double, gutter As Double
    Dim contentW As Double, colW As Double, x As Double
    marginCm = AskNumber("Margin on all four sides, in cm:", "1.00")
    If marginCm = NUM_CANCELLED Then Exit Sub
    colsIn = AskNumber("Number of columns (1 = margins only):", "3")
    If colsIn = NUM_CANCELLED Then Exit Sub
    nCols = CLng(colsIn)
    If nCols < 1 Or nCols > 24 Then
        MsgBox "Columns must be between 1 and 24.", vbExclamation, APP_NAME
        Exit Sub
    End If
    If nCols > 1 Then
        gutterCm = AskNumber("Gutter between columns, in cm:", "0.40")
        If gutterCm = NUM_CANCELLED Then Exit Sub
    End If

    sw = ActivePres.PageSetup.SlideWidth
    margin = CmToPt(marginCm)
    gutter = CmToPt(gutterCm)

    ' margins
    AddGuide True, margin
    AddGuide True, sw - margin
    AddGuide False, margin
    AddGuide False, ActivePres.PageSetup.SlideHeight - margin

    ' column boundaries (pair of guides per gutter)
    If nCols > 1 Then
        contentW = sw - 2 * margin
        colW = (contentW - (nCols - 1) * gutter) / nCols
        If colW <= 0 Then
            MsgBox "Margins + gutters leave no room for columns.", vbExclamation, APP_NAME
            Exit Sub
        End If
        x = margin
        For i = 1 To nCols - 1
            x = x + colW
            AddGuide True, x
            AddGuide True, x + gutter
            x = x + gutter
        Next i
    End If
    Exit Sub
Oops:
    ReportError OP
End Sub

' Four guides on the selection's bounding box.
Public Sub GuidesFromSelection(control As Object)
    Const OP As String = "Guides from selection"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long, shp As Shape
    Dim l As Double, t As Double, r As Double, b As Double
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    l = 1E+99: t = 1E+99: r = -1E+99: b = -1E+99
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If shp.Left < l Then l = shp.Left
        If shp.Top < t Then t = shp.Top
        If shp.Left + shp.Width > r Then r = shp.Left + shp.Width
        If shp.Top + shp.Height > b Then b = shp.Top + shp.Height
    Next i
    AddGuide True, l
    AddGuide True, r
    AddGuide False, t
    AddGuide False, b
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub DeleteAllGuides(control As Object)
    Const OP As String = "Delete all guides"
    On Error GoTo Oops
    Dim i As Long, n As Long
    n = ActivePres.Guides.Count
    If n = 0 Then
        Inform "No guides to delete."
        Exit Sub
    End If
    For i = n To 1 Step -1
        ActivePres.Guides(i).Delete
    Next i
    Inform n & " guide(s) deleted."
    Exit Sub
Oops:
    ReportError OP
End Sub
