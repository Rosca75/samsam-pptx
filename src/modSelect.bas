Attribute VB_Name = "modSelect"
Option Explicit

' ============================================================================
' modSelect — Tier 2 selection superpowers.
' Select-same compares against the REFERENCE (first-selected) shape and
' extends the selection to every matching top-level shape on the slide.
' ============================================================================

Private Const SIZE_TOLERANCE_PT As Double = 0.5

' True when both colour formats resolve to the same colour, respecting theme
' linkage (two shapes on the same theme slot match even mid-theme-change).
Private Function SameColor(ByVal a As ColorFormat, ByVal b As ColorFormat) As Boolean
    If a.ObjectThemeColor <> msoNotThemeColor Or b.ObjectThemeColor <> msoNotThemeColor Then
        SameColor = (a.ObjectThemeColor = b.ObjectThemeColor And _
                     Abs(a.Brightness - b.Brightness) < 0.01)
    Else
        SameColor = (a.RGB = b.RGB)
    End If
End Function

Private Function SameFill(ByVal a As Shape, ByVal b As Shape) As Boolean
    On Error Resume Next
    If a.Fill.Visible <> b.Fill.Visible Then Exit Function
    If a.Fill.Visible = msoFalse Then
        SameFill = True
    Else
        SameFill = SameColor(a.Fill.ForeColor, b.Fill.ForeColor)
    End If
End Function

Private Function SameFont(ByVal a As Shape, ByVal b As Shape) As Boolean
    On Error Resume Next
    If Not (a.HasTextFrame And b.HasTextFrame) Then Exit Function
    If Not (a.TextFrame.HasText And b.TextFrame.HasText) Then Exit Function
    SameFont = (a.TextFrame.TextRange.Font.Name = b.TextFrame.TextRange.Font.Name)
End Function

Private Function SameSize(ByVal a As Shape, ByVal b As Shape) As Boolean
    SameSize = (Abs(a.Width - b.Width) < SIZE_TOLERANCE_PT And _
                Abs(a.Height - b.Height) < SIZE_TOLERANCE_PT)
End Function

Private Function SameType(ByVal a As Shape, ByVal b As Shape) As Boolean
    If a.Type <> b.Type Then Exit Function
    If a.Type = msoAutoShape Then
        SameType = (a.AutoShapeType = b.AutoShapeType)
    Else
        SameType = True
    End If
End Function

' Tag-dispatched: "FILL" | "FONT" | "SIZE" | "TYPE".
Public Sub SelectSame(control As Object)
    Const OP As String = "Select same"
    On Error GoTo Oops
    Dim sr As ShapeRange, ref As Shape, shp As Shape, sld As Slide
    Dim names() As String, n As Long, matched As Boolean
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    Set ref = RefShape(sr)
    Set sld = CurrentSlide()
    If sld Is Nothing Then
        MsgBox "Switch to Normal view on a slide first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    ReDim names(1 To sld.Shapes.Count)
    For Each shp In sld.Shapes
        Select Case control.Tag
            Case "FILL": matched = SameFill(ref, shp)
            Case "FONT": matched = SameFont(ref, shp)
            Case "SIZE": matched = SameSize(ref, shp)
            Case "TYPE": matched = SameType(ref, shp)
            Case Else:   matched = False
        End Select
        If matched Then
            n = n + 1
            names(n) = shp.Name
        End If
    Next shp
    If n = 0 Then
        Inform "No other shape on this slide matches the reference shape's " & _
               LCase$(control.Tag) & "."
        Exit Sub
    End If
    ReDim Preserve names(1 To n)
    sld.Shapes.Range(names).Select
    Exit Sub
Oops:
    ReportError OP
End Sub

' True when the shape lies entirely outside the slide area ("gutter junk").
Private Function IsOffSlide(ByVal shp As Shape, ByVal sw As Double, ByVal sh As Double) As Boolean
    IsOffSlide = (shp.Left + shp.Width <= 0 Or shp.Top + shp.Height <= 0 Or _
                  shp.Left >= sw Or shp.Top >= sh)
End Function

' Select every fully off-slide shape on the CURRENT slide.
Public Sub SelectOffSlideShapes(control As Object)
    Const OP As String = "Select off-slide shapes"
    On Error GoTo Oops
    Dim sld As Slide, shp As Shape, names() As String, n As Long
    Dim sw As Double, sh As Double
    Set sld = CurrentSlide()
    If sld Is Nothing Then
        MsgBox "Switch to Normal view on a slide first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    sw = ActivePres.PageSetup.SlideWidth
    sh = ActivePres.PageSetup.SlideHeight
    If sld.Shapes.Count = 0 Then
        Inform "This slide has no shapes."
        Exit Sub
    End If
    ReDim names(1 To sld.Shapes.Count)
    For Each shp In sld.Shapes
        If IsOffSlide(shp, sw, sh) Then
            n = n + 1
            names(n) = shp.Name
        End If
    Next shp
    If n = 0 Then
        Inform "No off-slide shapes on this slide."
        Exit Sub
    End If
    ReDim Preserve names(1 To n)
    sld.Shapes.Range(names).Select
    Inform n & " off-slide shape(s) selected. Press Delete to remove them."
    Exit Sub
Oops:
    ReportError OP
End Sub

' Report off-slide shapes across the WHOLE deck (selection is per-slide only).
Public Sub ReportOffSlideShapes(control As Object)
    Const OP As String = "Report off-slide shapes"
    On Error GoTo Oops
    Dim sld As Slide, shp As Shape, msg As String, total As Long
    Dim sw As Double, sh As Double
    sw = ActivePres.PageSetup.SlideWidth
    sh = ActivePres.PageSetup.SlideHeight
    For Each sld In ActivePres.Slides
        For Each shp In sld.Shapes
            If IsOffSlide(shp, sw, sh) Then
                total = total + 1
                If total <= 30 Then
                    msg = msg & "Slide " & sld.SlideIndex & ": " & shp.Name & vbCrLf
                End If
            End If
        Next shp
    Next sld
    If total = 0 Then
        Inform "No off-slide shapes anywhere in the deck."
    Else
        Inform total & " off-slide shape(s) found:" & vbCrLf & vbCrLf & msg & _
               IIf(total > 30, "… and " & (total - 30) & " more." & vbCrLf, "") & vbCrLf & _
               "Go to each slide and use 'Select off-slide' to grab them."
    End If
    Exit Sub
Oops:
    ReportError OP
End Sub
