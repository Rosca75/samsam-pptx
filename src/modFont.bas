Attribute VB_Name = "modFont"
Option Explicit

' ============================================================================
' modFont — Tier 1 font tools.
' Theme fonts are applied as "+mj-lt" / "+mn-lt" placeholder names so the text
' keeps TRACKING the theme (same philosophy as ObjectThemeColor for colours).
' Font size changes walk TextRange.Runs so mixed-size selections step evenly.
' ============================================================================

Private Const FONT_MIN As Double = 4
Private Const FONT_MAX As Double = 400

' All selected shapes that can hold text, as a Collection (incl. table cells).
Private Function TextTargets(ByVal sr As ShapeRange) As Collection
    Dim col As New Collection, leaf As Collection, i As Long, shp As Shape
    Set leaf = New Collection
    For i = 1 To sr.Count
        CollectLeafShapes sr.Item(i), leaf
    Next i
    For Each shp In leaf
        If shp.HasTextFrame Then col.Add shp
    Next shp
    Set TextTargets = col
End Function

' Tag-dispatched: "HEAD" (theme heading font) | "BODY" (theme body font).
Public Sub ApplyThemeFont(control As Object)
    Const OP As String = "Apply theme font"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, themeName As String
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    ' "+mj-lt" = major (heading) latin, "+mn-lt" = minor (body) latin —
    ' assigning the placeholder keeps the text linked to the theme.
    themeName = IIf(control.Tag = "HEAD", "+mj-lt", "+mn-lt")
    For Each shp In TextTargets(sr)
        shp.TextFrame2.TextRange.Font.Name = themeName
    Next shp
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag carries the point size ("10", "12", "14", "18", "24" ribbon buttons).
Public Sub SetFontSizeButton(control As Object)
    Const OP As String = "Set font size"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For Each shp In TextTargets(sr)
        If shp.TextFrame.HasText Then shp.TextFrame.TextRange.Font.Size = Val(control.Tag)
    Next shp
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub SetExactFontSize(control As Object)
    Const OP As String = "Set exact font size"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, v As Double
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    v = AskNumber("Font size in points:", "12")
    If v = NUM_CANCELLED Then Exit Sub
    If v < FONT_MIN Or v > FONT_MAX Then
        MsgBox "Size must be between " & FONT_MIN & " and " & FONT_MAX & " pt.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    For Each shp In TextTargets(sr)
        If shp.TextFrame.HasText Then shp.TextFrame.TextRange.Font.Size = v
    Next shp
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "+" | "-" — 2 pt steps per run, so a mixed selection keeps
' its size differences instead of collapsing to one size.
Public Sub StepFontSize(control As Object)
    Const OP As String = "Grow/shrink font"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, run As TextRange
    Dim delta As Double, newSize As Double
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    delta = IIf(control.Tag = "+", 2, -2)
    For Each shp In TextTargets(sr)
        If shp.TextFrame.HasText Then
            For Each run In shp.TextFrame.TextRange.Runs
                newSize = run.Font.Size + delta
                If newSize < FONT_MIN Then newSize = FONT_MIN
                If newSize > FONT_MAX Then newSize = FONT_MAX
                run.Font.Size = newSize
            Next run
        End If
    Next shp
    Exit Sub
Oops:
    ReportError OP
End Sub

' Match the reference shape's font (name, size, bold, italic, colour) onto the
' other selected shapes. Theme-linked colours stay theme-linked.
Public Sub MatchFontToReference(control As Object)
    Const OP As String = "Match font"
    On Error GoTo Oops
    Dim sr As ShapeRange, ref As Shape, shp As Shape, i As Long
    Dim refFont As Object ' TextRange2 font of the reference
    Dim tgtFont As Object
    Set sr = GuardShapes(2, OP)
    If sr Is Nothing Then Exit Sub
    Set ref = RefShape(sr)
    If Not ref.HasTextFrame Then
        MsgBox "The reference (first-selected) shape has no text to copy the font from.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    Set refFont = ref.TextFrame2.TextRange.Font
    For i = 2 To sr.Count
        Set shp = sr.Item(i)
        If shp.HasTextFrame Then
            Set tgtFont = shp.TextFrame2.TextRange.Font
            tgtFont.Name = refFont.Name
            If refFont.Size > 0 Then tgtFont.Size = refFont.Size
            tgtFont.Bold = refFont.Bold
            tgtFont.Italic = refFont.Italic
            ' colour: preserve theme linkage when the reference has it
            If refFont.Fill.ForeColor.ObjectThemeColor <> msoNotThemeColor Then
                tgtFont.Fill.ForeColor.ObjectThemeColor = refFont.Fill.ForeColor.ObjectThemeColor
                tgtFont.Fill.ForeColor.Brightness = refFont.Fill.ForeColor.Brightness
            Else
                tgtFont.Fill.ForeColor.RGB = refFont.Fill.ForeColor.RGB
            End If
        End If
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub
