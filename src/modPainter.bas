Attribute VB_Name = "modPainter"
Option Explicit

' ============================================================================
' modPainter — Tier 3 format painting with memory.
' FRAGILITY QUARANTINE: the sticky painter needs Application events
' (clsAppEvents, WithEvents + WindowSelectionChange). Everything event-related
' is isolated HERE + clsAppEvents so a failure cannot break other modules —
' keep it that way. Every event-path procedure swallows its own errors and, on
' repeated failure, switches the sticky painter off instead of cascading.
'
' Pick up / apply uses the native Shape.PickUp / Shape.Apply pair, which
' captures the complete format (fill, line, effects, text formatting).
' ============================================================================

Private mSticky As Boolean          ' sticky painter armed?
Private mEvents As clsAppEvents     ' event sink (Nothing when inactive)
Private mHavePickup As Boolean      ' PickUp done at least once this session
Private mApplying As Boolean        ' re-entrancy guard for the event handler

' ----------------------------------------------------------------------------
' One-shot pick up / apply (safe, no events involved)
' ----------------------------------------------------------------------------

Public Sub PickUpFormat(control As Object)
    Const OP As String = "Pick up format"
    On Error GoTo Oops
    Dim sr As ShapeRange
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    RefShape(sr).PickUp
    mHavePickup = True
    Inform "Format picked up from '" & RefShape(sr).Name & "'." & vbCrLf & _
           "Select any shapes (any slide) and click 'Apply format' — as often " & _
           "as you like. Or switch on the Sticky painter."
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub ApplyFormat(control As Object)
    Const OP As String = "Apply format"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long
    If Not mHavePickup Then
        MsgBox "Nothing picked up yet. Select the source shape and click " & _
               "'Pick up format' first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        sr.Item(i).Apply
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' ----------------------------------------------------------------------------
' Sticky painter (event-driven — the fragile part)
' ----------------------------------------------------------------------------

Public Function StickyActive() As Boolean
    StickyActive = mSticky
End Function

' Called by the ribbon toggle (modRibbon.OnPainterToggle).
Public Sub SetSticky(ByVal turnOn As Boolean)
    Const OP As String = "Sticky painter"
    On Error GoTo Oops
    If turnOn Then
        If Not mHavePickup Then
            ' arm directly from the current selection when possible
            Dim sr As ShapeRange
            Set sr = GuardShapes(1, "Sticky painter (pick up source)")
            If sr Is Nothing Then
                mSticky = False
                If Not modRibbon.gRibbon Is Nothing Then modRibbon.gRibbon.Invalidate
                Exit Sub
            End If
            RefShape(sr).PickUp
            mHavePickup = True
        End If
        Set mEvents = New clsAppEvents
        Set mEvents.App = Application
        mSticky = True
        Inform "Sticky painter ON — every shape you now click gets the " & _
               "picked-up format. Click the toggle again to stop."
    Else
        StopSticky
    End If
    Exit Sub
Oops:
    StopSticky
    ReportError OP
End Sub

Public Sub StopSticky()
    On Error Resume Next
    mSticky = False
    If Not mEvents Is Nothing Then Set mEvents.App = Nothing
    Set mEvents = Nothing
End Sub

' Called by clsAppEvents on every selection change. MUST never raise: the
' event sink runs outside any user-facing call chain.
Public Sub OnSelectionChanged(ByVal sel As Selection)
    On Error GoTo Bail
    If Not mSticky Or mApplying Then Exit Sub
    If sel.Type <> ppSelectionShapes Then Exit Sub
    mApplying = True
    Dim i As Long
    For i = 1 To sel.ShapeRange.Count
        sel.ShapeRange.Item(i).Apply
    Next i
    mApplying = False
    Exit Sub
Bail:
    ' fail SAFE: disarm rather than error-loop on every click
    mApplying = False
    StopSticky
    On Error Resume Next
    If Not modRibbon.gRibbon Is Nothing Then modRibbon.gRibbon.Invalidate
    MsgBox "Sticky painter hit an error and switched itself off (" & _
           Err.Description & ").", vbExclamation, APP_NAME
End Sub

' ----------------------------------------------------------------------------
' Make-same adjustments (corner radius etc.) — safe, no events
' ----------------------------------------------------------------------------

' Copy ALL adjustment-handle values (corner radius, chevron angle, ...) from
' the reference to every same-type shape in the selection.
Public Sub MatchAdjustments(control As Object)
    Const OP As String = "Match shape adjustments"
    On Error GoTo Oops
    Dim sr As ShapeRange, ref As Shape, shp As Shape, i As Long, a As Long
    Dim skipped As Long
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    Set ref = RefShape(sr)
    If ref.Adjustments.Count = 0 Then
        MsgBox "The reference (first-selected) shape has no adjustment handles.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    For i = 2 To sr.Count
        Set shp = sr.Item(i)
        If shp.Type = ref.Type And shp.AutoShapeType = ref.AutoShapeType Then
            For a = 1 To ref.Adjustments.Count
                On Error Resume Next
                shp.Adjustments(a) = ref.Adjustments(a)
                On Error GoTo Oops
            Next a
        Else
            skipped = skipped + 1
        End If
    Next i
    If skipped > 0 Then Inform skipped & " shape(s) skipped (different shape " & _
                               "type from the reference)."
    Exit Sub
Oops:
    ReportError OP
End Sub

' Corner radius convenience: rounded rectangles only, radius prompt in cm.
' (Adjustment 1 of a rounded rectangle = radius / smaller side.)
Public Sub SetCornerRadius(control As Object)
    Const OP As String = "Set corner radius"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, i As Long
    Dim radCm As Double, small As Double, done As Long
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    radCm = AskNumber("Corner radius in cm (applied to every selected " & _
                      "rounded rectangle):", "0.20")
    If radCm = NUM_CANCELLED Then Exit Sub
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If shp.Type = msoAutoShape And shp.AutoShapeType = msoShapeRoundedRectangle Then
            small = IIf(shp.Width < shp.Height, shp.Width, shp.Height)
            If small > 0 Then
                shp.Adjustments(1) = CmToPt(radCm) / small
                done = done + 1
            End If
        End If
    Next i
    If done = 0 Then
        Inform "No rounded rectangles in the selection — nothing changed. " & _
               "(Use 'Match adjustments' for other shape types.)"
    End If
    Exit Sub
Oops:
    ReportError OP
End Sub
