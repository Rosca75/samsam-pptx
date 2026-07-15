Attribute VB_Name = "modAudit"
Option Explicit

' ============================================================================
' modAudit — Tier 3, built LAST, kept DEGRADABLE by design:
'   level 1 (always works): DeckStatistics — pure read-out.
'   level 2 (report-only):  OffThemeColorReport — lists every hard-coded
'                           (non-theme) fill/line/text colour in the deck.
'                           Long reports land on a tagged report slide.
'   level 3 (quick fix):    SelectOffThemeOnSlide — grabs offenders on the
'                           current slide so the user can recolour them with
'                           the theme galleries.
' Nothing here mutates content except the optional report slide (tagged
' SAMSAM=REPORT, removable via its own button).
' ============================================================================

Private Const TAG_KEY As String = "SAMSAM"
Private Const TAG_REPORT As String = "REPORT"
Private Const MSGBOX_LIMIT As Long = 25      ' report lines shown in a MsgBox
Private Const SLIDE_TEXT_LIMIT As Long = 6000

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

' ----------------------------------------------------------------------------
' Off-theme colour report
' ----------------------------------------------------------------------------

' A colour is "on theme" when it is theme-linked (ObjectThemeColor set) or its
' RGB happens to equal one of the 12 scheme colours (a detached-but-identical
' colour — flagged separately, since it will NOT follow a theme change).
Private Function ThemeSlotOfRGB(ByVal rgbVal As Long) As Long
    Dim i As Long
    For i = 1 To 12
        If ThemeRGB(i) = rgbVal Then
            ThemeSlotOfRGB = i
            Exit Function
        End If
    Next i
    ThemeSlotOfRGB = 0
End Function

' Classifies one ColorFormat. Returns "" when clean, else a short description.
Private Function CheckColor(ByVal cf As ColorFormat) As String
    Dim slot As Long
    If cf.ObjectThemeColor <> msoNotThemeColor Then Exit Function ' theme-linked: clean
    slot = ThemeSlotOfRGB(cf.RGB)
    If slot > 0 Then
        CheckColor = "#" & Right$("000000" & Hex(RGBToHexOrder(cf.RGB)), 6) & _
                     " (matches " & ThemeColorName(slot) & " but is DETACHED)"
    Else
        CheckColor = "#" & Right$("000000" & Hex(RGBToHexOrder(cf.RGB)), 6)
    End If
End Function

' All findings for one shape appended to lines; returns number found.
Private Function AuditShape(ByVal shp As Shape, ByVal slideNo As Long, _
                            ByRef lines As String) As Long
    Dim desc As String, n As Long, run As TextRange2
    On Error Resume Next
    If shp.Fill.Visible = msoTrue And shp.Fill.Type = msoFillSolid Then
        desc = CheckColor(shp.Fill.ForeColor)
        If Len(desc) > 0 Then
            lines = lines & "Slide " & slideNo & " · " & shp.Name & " · fill · " & desc & vbCrLf
            n = n + 1
        End If
    End If
    If shp.Line.Visible = msoTrue Then
        desc = CheckColor(shp.Line.ForeColor)
        If Len(desc) > 0 Then
            lines = lines & "Slide " & slideNo & " · " & shp.Name & " · line · " & desc & vbCrLf
            n = n + 1
        End If
    End If
    If shp.HasTextFrame Then
        If shp.TextFrame.HasText Then
            For Each run In shp.TextFrame2.TextRange.Runs
                If run.Font.Fill.Type = msoFillSolid Then
                    desc = CheckColor(run.Font.Fill.ForeColor)
                    If Len(desc) > 0 Then
                        lines = lines & "Slide " & slideNo & " · " & shp.Name & _
                                " · text · " & desc & vbCrLf
                        n = n + 1
                        Exit For        ' one text finding per shape is enough
                    End If
                End If
            Next run
        End If
    End If
    AuditShape = n
End Function

Public Sub OffThemeColorReport(control As Object)
    Const OP As String = "Off-theme colour report"
    On Error GoTo Oops
    Dim pres As Presentation, sld As Slide, shp As Shape
    Dim lines As String, total As Long
    Set pres = ActivePres
    For Each sld In pres.Slides
        If Not sld.Tags(TAG_KEY) = TAG_REPORT Then
            For Each shp In LeafShapes(sld.Shapes)
                total = total + AuditShape(shp, sld.SlideIndex, lines)
            Next shp
        End If
    Next sld
    If total = 0 Then
        Inform "Clean deck: every fill, line and text colour is theme-linked. " & _
               "Nothing will break when the theme changes."
        Exit Sub
    End If
    If total <= MSGBOX_LIMIT Then
        Inform total & " off-theme colour(s) found:" & vbCrLf & vbCrLf & lines & vbCrLf & _
               "Fix: go to the slide, use 'Select off-theme on slide', then " & _
               "recolour with the theme galleries."
    Else
        WriteReportSlide pres, total, lines
    End If
    Exit Sub
Oops:
    ReportError OP
End Sub

' Long reports go onto a tagged slide at the end of the deck (rebuilt each run).
Private Sub WriteReportSlide(ByVal pres As Presentation, ByVal total As Long, _
                             ByVal lines As String)
    Dim i As Long, sld As Slide, box As Shape, body As String
    For i = pres.Slides.Count To 1 Step -1
        If pres.Slides(i).Tags(TAG_KEY) = TAG_REPORT Then pres.Slides(i).Delete
    Next i
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
    sld.Tags.Add TAG_KEY, TAG_REPORT
    body = "OFF-THEME COLOUR REPORT — " & total & " finding(s)" & vbCrLf & _
           "(Delete this slide when done; rerunning the report replaces it.)" & _
           vbCrLf & vbCrLf & lines
    If Len(body) > SLIDE_TEXT_LIMIT Then
        body = Left$(body, SLIDE_TEXT_LIMIT) & vbCrLf & "… report truncated."
    End If
    Set box = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, CmToPt(1), CmToPt(1), _
              pres.PageSetup.SlideWidth - CmToPt(2), pres.PageSetup.SlideHeight - CmToPt(2))
    box.TextFrame.TextRange.Text = body
    box.TextFrame.TextRange.Font.Size = 10
    box.TextFrame.WordWrap = msoTrue
    Inform total & " off-theme colours found — full report written to a new " & _
           "slide at the end of the deck (slide " & sld.SlideIndex & ")."
    ActiveWindow.View.GotoSlide sld.SlideIndex
End Sub

' Quick fix: select every top-level shape on the CURRENT slide that carries an
' off-theme fill/line/text colour, ready for the theme galleries.
Public Sub SelectOffThemeOnSlide(control As Object)
    Const OP As String = "Select off-theme shapes on slide"
    On Error GoTo Oops
    Dim sld As Slide, shp As Shape, names() As String, n As Long, dummy As String
    Set sld = CurrentSlide()
    If sld Is Nothing Then
        MsgBox "Switch to Normal view on a slide first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    If sld.Shapes.Count = 0 Then
        Inform "This slide has no shapes."
        Exit Sub
    End If
    ReDim names(1 To sld.Shapes.Count)
    For Each shp In sld.Shapes
        dummy = ""
        If AuditShape(shp, sld.SlideIndex, dummy) > 0 Then
            n = n + 1
            names(n) = shp.Name
        End If
    Next shp
    If n = 0 Then
        Inform "All colours on this slide are theme-linked."
        Exit Sub
    End If
    ReDim Preserve names(1 To n)
    sld.Shapes.Range(names).Select
    Inform n & " shape(s) with off-theme colours selected — recolour them " & _
           "with the theme galleries so they follow the theme."
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub RemoveReportSlide(control As Object)
    Const OP As String = "Remove report slide"
    On Error GoTo Oops
    Dim i As Long, removed As Long
    For i = ActivePres.Slides.Count To 1 Step -1
        If ActivePres.Slides(i).Tags(TAG_KEY) = TAG_REPORT Then
            ActivePres.Slides(i).Delete
            removed = removed + 1
        End If
    Next i
    Inform removed & " report slide(s) removed."
    Exit Sub
Oops:
    ReportError OP
End Sub
