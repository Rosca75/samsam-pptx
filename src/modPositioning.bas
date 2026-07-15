Attribute VB_Name = "modPositioning"
Option Explicit

' ============================================================================
' modPositioning — Tier 1 align/centre/nudge/exact-XY + Tier 2 THOR-style
' pick-up/apply position & size, and swap.
' Align mode comes from modRibbon.AlignToSlideMode (ribbon toggle):
'   * align to SLIDE:      every selected shape aligns to the slide bounds
'   * align to SELECTION:  shapes align to the REFERENCE (first-selected) shape
' Module-level state below MUST stay above the first Sub (VBA rule).
' ============================================================================

' Nudge step (cm) — session state, initialised lazily.
Private mNudgeStepCm As Double
Private mNudgeInit As Boolean

' THOR pick-up/apply store — session state.
Private mHasPickup As Boolean
Private mPickLeft As Double
Private mPickTop As Double
Private mPickWidth As Double
Private mPickHeight As Double

Private Function NudgeStepPt() As Double
    If Not mNudgeInit Then
        mNudgeStepCm = 0.1
        mNudgeInit = True
    End If
    NudgeStepPt = CmToPt(mNudgeStepCm)
End Function

' Tag-dispatched: "L" | "C" | "R" | "T" | "M" | "B".
Public Sub AlignShapes(control As Object)
    Const OP As String = "Align"
    On Error GoTo Oops
    Dim sr As ShapeRange, ref As Shape, shp As Shape, i As Long, first As Long
    Dim sw As Double, sh As Double
    If modRibbon.AlignToSlideMode() Then
        Set sr = GuardShapes(1, OP & " to slide")
        If sr Is Nothing Then Exit Sub
        sw = ActivePres.PageSetup.SlideWidth
        sh = ActivePres.PageSetup.SlideHeight
        For i = 1 To sr.Count
            Set shp = sr.Item(i)
            Select Case control.Tag
                Case "L": shp.Left = 0
                Case "C": shp.Left = (sw - shp.Width) / 2
                Case "R": shp.Left = sw - shp.Width
                Case "T": shp.Top = 0
                Case "M": shp.Top = (sh - shp.Height) / 2
                Case "B": shp.Top = sh - shp.Height
            End Select
        Next i
    Else
        Set sr = GuardShapes(2, OP & " to reference shape")
        If sr Is Nothing Then Exit Sub
        Set ref = RefShape(sr)
        For i = 2 To sr.Count
            Set shp = sr.Item(i)
            Select Case control.Tag
                Case "L": shp.Left = ref.Left
                Case "C": shp.Left = ref.Left + (ref.Width - shp.Width) / 2
                Case "R": shp.Left = ref.Left + ref.Width - shp.Width
                Case "T": shp.Top = ref.Top
                Case "M": shp.Top = ref.Top + (ref.Height - shp.Height) / 2
                Case "B": shp.Top = ref.Top + ref.Height - shp.Height
            End Select
        Next i
    End If
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "H" | "V" | "B" — centre each shape on the slide.
Public Sub CenterOnSlide(control As Object)
    Const OP As String = "Centre on slide"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, i As Long
    Dim sw As Double, sh As Double
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    sw = ActivePres.PageSetup.SlideWidth
    sh = ActivePres.PageSetup.SlideHeight
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If control.Tag = "H" Or control.Tag = "B" Then shp.Left = (sw - shp.Width) / 2
        If control.Tag = "V" Or control.Tag = "B" Then shp.Top = (sh - shp.Height) / 2
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "L" | "R" | "U" | "D" — precise nudge by the configured step.
Public Sub NudgeShapes(control As Object)
    Const OP As String = "Nudge"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, i As Long, stepPt As Double
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    stepPt = NudgeStepPt()
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        Select Case control.Tag
            Case "L": shp.Left = shp.Left - stepPt
            Case "R": shp.Left = shp.Left + stepPt
            Case "U": shp.Top = shp.Top - stepPt
            Case "D": shp.Top = shp.Top + stepPt
        End Select
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub SetNudgeStep(control As Object)
    Const OP As String = "Set nudge step"
    On Error GoTo Oops
    Dim v As Double
    NudgeStepPt ' ensure initialised so the default shows the live value
    v = AskNumber("Nudge step in cm:", Format(mNudgeStepCm, "0.00"))
    If v = NUM_CANCELLED Then Exit Sub
    If v <= 0 Or v > 10 Then
        MsgBox "Step must be between 0.01 and 10 cm.", vbExclamation, APP_NAME
        Exit Sub
    End If
    mNudgeStepCm = v
    Exit Sub
Oops:
    ReportError OP
End Sub

' Exact X/Y (cm, top-left corner) for all selected shapes. Blank keeps.
Public Sub SetExactPosition(control As Object)
    Const OP As String = "Set exact position"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, i As Long, x As Double, y As Double
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    x = AskNumber("X — distance from slide LEFT edge in cm (empty = keep):", _
                  Format(PtToCm(RefShape(sr).Left), "0.00"))
    y = AskNumber("Y — distance from slide TOP edge in cm (empty = keep):", _
                  Format(PtToCm(RefShape(sr).Top), "0.00"))
    If x = NUM_CANCELLED And y = NUM_CANCELLED Then Exit Sub
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If x <> NUM_CANCELLED Then shp.Left = CmToPt(x)
        If y <> NUM_CANCELLED Then shp.Top = CmToPt(y)
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' --- Tier 2: THOR pattern — pick up position & size once, apply anywhere ----

Public Sub PickUpPositionSize(control As Object)
    Const OP As String = "Pick up position && size"
    On Error GoTo Oops
    Dim sr As ShapeRange, ref As Shape
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    Set ref = RefShape(sr)
    mPickLeft = ref.Left
    mPickTop = ref.Top
    mPickWidth = ref.Width
    mPickHeight = ref.Height
    mHasPickup = True
    Inform "Picked up: X " & FmtCm(mPickLeft) & ", Y " & FmtCm(mPickTop) & _
           ", W " & FmtCm(mPickWidth) & ", H " & FmtCm(mPickHeight) & vbCrLf & _
           "Now select shapes (on any slide) and use one of the Apply buttons."
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "P" (position) | "S" (size) | "B" (both).
Public Sub ApplyPositionSize(control As Object)
    Const OP As String = "Apply position/size"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, i As Long, wasLocked As Boolean
    If Not mHasPickup Then
        MsgBox "Nothing picked up yet. Select a source shape and click " & _
               "'Pick up pos+size' first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If control.Tag = "S" Or control.Tag = "B" Then
            wasLocked = UnlockAspect(shp)
            shp.Width = mPickWidth
            shp.Height = mPickHeight
            RestoreAspect shp, wasLocked
        End If
        If control.Tag = "P" Or control.Tag = "B" Then
            shp.Left = mPickLeft
            shp.Top = mPickTop
        End If
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' Swap the positions of exactly two shapes (centre-anchored).
Public Sub SwapShapes(control As Object)
    Const OP As String = "Swap positions"
    On Error GoTo Oops
    Dim sr As ShapeRange, a As Shape, b As Shape
    Dim ax As Double, ay As Double, bx As Double, by As Double
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    If sr.Count <> 2 Then
        MsgBox "Swap needs EXACTLY two shapes selected (you have " & sr.Count & ").", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    Set a = sr.Item(1)
    Set b = sr.Item(2)
    ax = a.Left + a.Width / 2:  ay = a.Top + a.Height / 2
    bx = b.Left + b.Width / 2:  by = b.Top + b.Height / 2
    a.Left = bx - a.Width / 2:  a.Top = by - a.Height / 2
    b.Left = ax - b.Width / 2:  b.Top = ay - b.Height / 2
    Exit Sub
Oops:
    ReportError OP
End Sub
