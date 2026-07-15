Attribute VB_Name = "modColor"
Option Explicit

' ============================================================================
' modColor — Tier 1 FLAGSHIP: theme-colour application.
' RULE: colours are applied via ObjectThemeColor (msoThemeColorDark1=1 ..
' msoThemeColorAccent6=10), NEVER as hardcoded RGB, so shapes keep tracking
' theme changes. Tints & shades (Tier 3) are applied as ObjectThemeColor +
' ColorFormat.Brightness — still theme-linked. The ONLY calculated-RGB use is
' the gallery swatch bitmaps in modRibbon (pure preview pixels, never applied).
'
' The ribbon galleries (modRibbon) call ApplyThemeVariant. The fixed-button
' fallback (if dynamic galleries ever misbehave) calls ApplyThemeColorTagged
' with control.Tag = "<target>:<slot>[:<brightness>]", e.g. "fill:5" or
' "text:1:-0.25". Both paths end in the same ApplyToShape code.
' ============================================================================

' Brightness for the 6 gallery rows: base, lighter 80/60/40 %, darker 25/50 %.
' (Same ladder PowerPoint's own colour picker uses.)
Public Function VariantBrightness(ByVal variantRow As Long) As Double
    Select Case variantRow
        Case 1: VariantBrightness = 0.8
        Case 2: VariantBrightness = 0.6
        Case 3: VariantBrightness = 0.4
        Case 4: VariantBrightness = -0.25
        Case 5: VariantBrightness = -0.5
        Case Else: VariantBrightness = 0#
    End Select
End Function

Public Function VariantName(ByVal variantRow As Long) As String
    Select Case variantRow
        Case 1: VariantName = ", lighter 80%"
        Case 2: VariantName = ", lighter 60%"
        Case 3: VariantName = ", lighter 40%"
        Case 4: VariantName = ", darker 25%"
        Case 5: VariantName = ", darker 50%"
        Case Else: VariantName = ""
    End Select
End Function

' Apply one theme colour slot (1..10) + brightness to one shape for a target:
' "fill" | "line" | "text" | "both" (fill+line, "paint whole shape").
Private Sub ApplyToShape(ByVal shp As Shape, ByVal target As String, _
                         ByVal slot As Long, ByVal brightness As Double)
    Dim cf As ColorFormat
    If target = "fill" Or target = "both" Then
        shp.Fill.Visible = msoTrue
        Set cf = shp.Fill.ForeColor
        cf.ObjectThemeColor = slot          ' theme-linked, never RGB
        cf.Brightness = brightness
    End If
    If target = "line" Or target = "both" Then
        shp.Line.Visible = msoTrue
        Set cf = shp.Line.ForeColor
        cf.ObjectThemeColor = slot
        cf.Brightness = brightness
    End If
    If target = "text" Then
        If shp.HasTextFrame Then
            Set cf = shp.TextFrame2.TextRange.Font.Fill.ForeColor
            cf.ObjectThemeColor = slot
            cf.Brightness = brightness
        End If
    End If
End Sub

' Main entry used by the ribbon galleries (via modRibbon.OnPaletteAction).
Public Sub ApplyThemeVariant(ByVal target As String, ByVal slot As Long, _
                             ByVal variantRow As Long)
    Const OP As String = "Apply theme colour"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long, shp As Shape, leaf As Collection
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        Set shp = sr.Item(i)
        If shp.Type = msoGroup Or shp.HasTable Then
            ' recurse so table cells / group members get the colour too
            Set leaf = New Collection
            CollectLeafShapes shp, leaf
            Dim child As Shape
            For Each child In leaf
                On Error Resume Next   ' some leaves have no fill/line (e.g. table frame)
                ApplyToShape child, target, slot, VariantBrightness(variantRow)
                On Error GoTo Oops
            Next child
        Else
            ApplyToShape shp, target, slot, VariantBrightness(variantRow)
        End If
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' Fallback path: fixed ribbon buttons with Tag "<target>:<slot>[:<brightness>]".
Public Sub ApplyThemeColorTagged(control As Object)
    Const OP As String = "Apply theme colour"
    On Error GoTo Oops
    Dim parts() As String, target As String, slot As Long, bright As Double
    Dim sr As ShapeRange, i As Long
    parts = Split(control.Tag, ":")
    If UBound(parts) < 1 Then
        MsgBox "Bad colour button tag: '" & control.Tag & "'", vbCritical, APP_NAME
        Exit Sub
    End If
    target = parts(0)
    slot = CLng(parts(1))
    If UBound(parts) >= 2 Then bright = Val(Replace(parts(2), ",", "."))
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        ApplyToShape sr.Item(i), target, slot, bright
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

' Remove the fill / line ("no colour" buttons at the end of the galleries
' would complicate indexing, so these are plain buttons).
Public Sub RemoveFill(control As Object)
    Const OP As String = "No fill"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        sr.Item(i).Fill.Visible = msoFalse
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub RemoveLine(control As Object)
    Const OP As String = "No line"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        sr.Item(i).Line.Visible = msoFalse
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub
