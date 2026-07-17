Attribute VB_Name = "modSizing"
Option Explicit

' ============================================================================
' modSizing — Tier 1 sizing tools + Tier 3 stretch-edge-to-reference.
' Reference shape = FIRST selected. Top-left corners stay put when resizing.
' ============================================================================

' Tag-dispatched: control.Tag = "MAX" | "MIN" — size every shape like the
' largest/smallest (by area) shape in the selection.
Public Sub SizeToExtreme(control As Object)
    Const OP As String = "Size to largest/smallest"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, pick As Shape, i As Long
    Dim area As Double, best As Double, wasLocked As Boolean
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    Set pick = sr.Item(1)
    best = pick.Width * pick.Height
    For i = 2 To sr.Count
        area = sr.Item(i).Width * sr.Item(i).Height
        If (control.Tag = "MAX" And area > best) Or _
           (control.Tag = "MIN" And area < best) Then
            best = area
            Set pick = sr.Item(i)
        End If
    Next i
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If Not shp Is pick Then
            wasLocked = UnlockAspect(shp)
            shp.Width = pick.Width
            shp.Height = pick.Height
            RestoreAspect shp, wasLocked
        End If
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tier 3 — stretch one edge of every other shape to the reference shape's
' matching edge. Tag = "L" | "R" | "T" | "B". The opposite edge stays fixed.
Public Sub StretchEdgeToReference(control As Object)
    Const OP As String = "Stretch edge to reference"
    On Error GoTo Oops
    Dim sr As ShapeRange, ref As Shape, shp As Shape, i As Long
    Dim newVal As Double, wasLocked As Boolean
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    Set ref = RefShape(sr)
    For i = 2 To sr.Count
        Set shp = sr.Item(i)
        wasLocked = UnlockAspect(shp)
        Select Case control.Tag
            Case "L"   ' left edge moves to ref's left edge; right edge fixed
                newVal = shp.Left + shp.Width - ref.Left
                If newVal > 1 Then
                    shp.Left = ref.Left
                    shp.Width = newVal
                End If
            Case "R"   ' right edge moves to ref's right edge; left edge fixed
                newVal = (ref.Left + ref.Width) - shp.Left
                If newVal > 1 Then shp.Width = newVal
            Case "T"
                newVal = shp.Top + shp.Height - ref.Top
                If newVal > 1 Then
                    shp.Top = ref.Top
                    shp.Height = newVal
                End If
            Case "B"
                newVal = (ref.Top + ref.Height) - shp.Top
                If newVal > 1 Then shp.Height = newVal
        End Select
        RestoreAspect shp, wasLocked
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub
