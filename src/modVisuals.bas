Attribute VB_Name = "modVisuals"
Option Explicit

' ============================================================================
' modVisuals — Tier 3 generated-geometry micro-visuals: Harvey balls and RAG
' status dots. These are drawn from NATIVE shapes at insert time (parameterised
' geometry) — they are NOT a stored asset library, which stays out of scope.
' Harvey balls use theme colours (ObjectThemeColor) so they follow the theme.
' RAG dots use brand-fixed RGB — the documented exception: status colours must
' stay red/amber/green regardless of the deck theme.
' ============================================================================

Private Const BALL_SIZE_CM As Double = 1#
Private Const MSO_SHAPE_PIE As Long = 142   ' msoShapePie (numeric: the constant
                                            ' is missing from older Office libs)

' Insert location: centre of the current slide.
Private Function InsertOrigin(ByRef sld As Slide, ByRef x As Double, ByRef y As Double) As Boolean
    Set sld = CurrentSlide()
    If sld Is Nothing Then
        MsgBox "Switch to Normal view on a slide first.", vbExclamation, APP_NAME
        Exit Function
    End If
    x = (ActivePres.PageSetup.SlideWidth - CmToPt(BALL_SIZE_CM)) / 2
    y = (ActivePres.PageSetup.SlideHeight - CmToPt(BALL_SIZE_CM)) / 2
    InsertOrigin = True
End Function

' Tag carries the percentage: "0" | "25" | "50" | "75" | "100".
Public Sub InsertHarveyBall(control As Object)
    Const OP As String = "Insert Harvey ball"
    On Error GoTo Oops
    Dim sld As Slide, x As Double, y As Double, size As Double
    Dim ring As Shape, pie As Shape, ball As Shape, pct As Long
    pct = Val(control.Tag)
    If Not InsertOrigin(sld, x, y) Then Exit Sub
    size = CmToPt(BALL_SIZE_CM)

    ' outline ring (always present, theme dark colour)
    Set ring = sld.Shapes.AddShape(msoShapeOval, x, y, size, size)
    ring.Fill.Visible = msoFalse
    ring.Line.Visible = msoTrue
    ring.Line.Weight = 1.25
    ring.Line.ForeColor.ObjectThemeColor = 1      ' msoThemeColorDark1
    ring.Name = "HarveyRing_" & ring.Id

    If pct >= 100 Then
        ring.Fill.Visible = msoTrue
        ring.Fill.ForeColor.ObjectThemeColor = 5  ' msoThemeColorAccent1
        ring.Name = "HarveyBall100_" & ring.Id
        ring.Select
        Exit Sub
    End If

    If pct > 0 Then
        ' filled pie wedge from 12 o'clock, clockwise pct% of the circle
        Set pie = sld.Shapes.AddShape(MSO_SHAPE_PIE, x, y, size, size)
        pie.Adjustments(1) = -90
        pie.Adjustments(2) = -90 + 360# * pct / 100#
        pie.Fill.Visible = msoTrue
        pie.Fill.ForeColor.ObjectThemeColor = 5   ' msoThemeColorAccent1
        pie.Line.Visible = msoFalse
        pie.Name = "HarveyPie_" & pie.Id
        Set ball = sld.Shapes.Range(Array(ring.Name, pie.Name)).Group
        ball.Name = "HarveyBall" & pct & "_" & ball.Id
        ball.Select
    Else
        ring.Name = "HarveyBall0_" & ring.Id
        ring.Select
    End If
    Exit Sub
Oops:
    ReportError OP
End Sub

' Tag-dispatched: "R" | "A" | "G".
Public Sub InsertRAG(control As Object)
    Const OP As String = "Insert RAG status"
    On Error GoTo Oops
    Dim sld As Slide, x As Double, y As Double, dot As Shape
    If Not InsertOrigin(sld, x, y) Then Exit Sub
    Set dot = sld.Shapes.AddShape(msoShapeOval, x, y, _
                                  CmToPt(BALL_SIZE_CM), CmToPt(BALL_SIZE_CM))
    dot.Line.Visible = msoFalse
    dot.Fill.Visible = msoTrue
    ' BRAND-FIXED RGB (documented exception to the ObjectThemeColor rule):
    ' status colours must not drift with the theme.
    Select Case control.Tag
        Case "R": dot.Fill.ForeColor.RGB = RGB(192, 0, 0):    dot.Name = "RAG_Red_" & dot.Id
        Case "A": dot.Fill.ForeColor.RGB = RGB(255, 153, 0):  dot.Name = "RAG_Amber_" & dot.Id
        Case "G": dot.Fill.ForeColor.RGB = RGB(0, 153, 51):   dot.Name = "RAG_Green_" & dot.Id
    End Select
    dot.Select
    Exit Sub
Oops:
    ReportError OP
End Sub
