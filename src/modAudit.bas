Attribute VB_Name = "modAudit"
Option Explicit

' ============================================================================
' modAudit — Tier 3, built LAST, kept DEGRADABLE by design:
'   level 1 (always works): DeckStatistics — pure read-out.
' Nothing here mutates content.
' ============================================================================

' ----------------------------------------------------------------------------
' Deck statistics
' ----------------------------------------------------------------------------

Private Sub AddFontName(ByVal col As Collection, ByVal fontName As String)
    If Len(fontName) = 0 Then Exit Sub
    On Error Resume Next
    col.Add fontName, fontName   ' keyed add — duplicates throw and are ignored
End Sub

Public Sub DeckStatistics(control As Object)
    Const OP As String = "Deck statistics"
    On Error GoTo Oops
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim slideCount As Long, hiddenSlides As Long, shapeCount As Long
    Dim textShapes As Long, wordCount As Long, offSlide As Long
    Dim fonts As New Collection, run As TextRange, fontList As String
    Dim sw As Double, sh As Double, i As Long
    Set pres = ActivePres
    sw = pres.PageSetup.SlideWidth
    sh = pres.PageSetup.SlideHeight
    slideCount = pres.Slides.Count
    For Each sld In pres.Slides
        If sld.SlideShowTransition.Hidden = msoTrue Then hiddenSlides = hiddenSlides + 1
        For Each shp In LeafShapes(sld.Shapes)
            shapeCount = shapeCount + 1
            If shp.Left + shp.Width <= 0 Or shp.Top + shp.Height <= 0 Or _
               shp.Left >= sw Or shp.Top >= sh Then offSlide = offSlide + 1
            If ShapeHasText(shp) Then
                textShapes = textShapes + 1
                On Error Resume Next
                wordCount = wordCount + shp.TextFrame.TextRange.Words.Count
                For Each run In shp.TextFrame.TextRange.Runs
                    AddFontName fonts, run.Font.Name
                Next run
                On Error GoTo Oops
            End If
        Next shp
    Next sld
    For i = 1 To fonts.Count
        fontList = fontList & IIf(i > 1, ", ", "") & fonts(i)
        If i = 15 And fonts.Count > 15 Then
            fontList = fontList & " … +" & (fonts.Count - 15) & " more"
            Exit For
        End If
    Next i
    Inform "Deck statistics — " & pres.Name & vbCrLf & vbCrLf & _
           "Slides:           " & slideCount & _
           IIf(hiddenSlides > 0, "  (" & hiddenSlides & " hidden)", "") & vbCrLf & _
           "Slide masters:    " & pres.Designs.Count & vbCrLf & _
           "Shapes:           " & shapeCount & vbCrLf & _
           "  with text:      " & textShapes & vbCrLf & _
           "  fully off-slide:" & Str$(offSlide) & vbCrLf & _
           "Words:            " & wordCount & vbCrLf & _
           "Fonts in use (" & fonts.Count & "): " & fontList
    Exit Sub
Oops:
    ReportError OP
End Sub
