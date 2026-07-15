Attribute VB_Name = "modAgenda"
Option Explicit

' ============================================================================
' modAgenda — Tier 3 agenda/section tools.
' Generates one divider slide at the start of each PowerPoint SECTION, showing
' the section name as title plus the full agenda list with the current section
' emphasised. Divider slides are tagged (SAMSAM=AGENDA) so rebuilding replaces
' them cleanly and never touches real content slides.
' ============================================================================

Private Const TAG_KEY As String = "SAMSAM"
Private Const TAG_AGENDA As String = "AGENDA"

Private Sub DeleteTaggedAgendaSlides(ByVal pres As Presentation)
    Dim i As Long
    For i = pres.Slides.Count To 1 Step -1
        If pres.Slides(i).Tags(TAG_KEY) = TAG_AGENDA Then pres.Slides(i).Delete
    Next i
End Sub

' Prefer a layout whose name says "section"; fall back to the layout of the
' slide the divider precedes (keeps the section's own design).
Private Function DividerLayout(ByVal pres As Presentation, ByVal followerSlide As Slide) As CustomLayout
    Dim des As Design, lay As CustomLayout
    For Each des In pres.Designs
        For Each lay In des.SlideMaster.CustomLayouts
            If InStr(1, lay.Name, "section", vbTextCompare) > 0 Then
                Set DividerLayout = lay
                Exit Function
            End If
        Next lay
    Next des
    Set DividerLayout = followerSlide.CustomLayout
End Function

Public Sub BuildAgenda(control As Object)
    Const OP As String = "Build/update agenda slides"
    On Error GoTo Oops
    Dim pres As Presentation, sp As Object ' SectionProperties
    Dim s As Long, firstIdx As Long, sld As Slide, follower As Slide
    Dim titleShape As Shape, listShape As Shape, tr As TextRange
    Dim k As Long, built As Long, para As TextRange
    Set pres = ActivePres
    Set sp = pres.SectionProperties
    If sp.Count = 0 Then
        MsgBox "This deck has no sections. Add sections first (right-click " & _
               "between slides in the thumbnail pane > Add Section).", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    ' rebuild = delete old tagged dividers, then insert fresh ones.
    DeleteTaggedAgendaSlides pres
    ' insert back-to-front so earlier indices stay valid.
    For s = sp.Count To 1 Step -1
        If sp.SlidesCount(s) > 0 Then
            firstIdx = sp.FirstSlide(s)
            Set follower = pres.Slides(firstIdx)
            Set sld = pres.Slides.AddSlide(firstIdx, DividerLayout(pres, follower))
            sld.MoveToSectionStart s
            sld.Tags.Add TAG_KEY, TAG_AGENDA

            ' title = section name (fall back to a text box if no placeholder)
            Set titleShape = Nothing
            On Error Resume Next
            Set titleShape = sld.Shapes.Title
            On Error GoTo Oops
            If titleShape Is Nothing Then
                Set titleShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                    CmToPt(2), CmToPt(2), pres.PageSetup.SlideWidth - CmToPt(4), CmToPt(3))
            End If
            titleShape.TextFrame.TextRange.Text = sp.Name(s)

            ' agenda list, current section emphasised
            Set listShape = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                CmToPt(2), CmToPt(6), pres.PageSetup.SlideWidth - CmToPt(4), _
                pres.PageSetup.SlideHeight - CmToPt(8))
            Set tr = listShape.TextFrame.TextRange
            For k = 1 To sp.Count
                If k > 1 Then tr.InsertAfter vbCr
                tr.InsertAfter sp.Name(k)
            Next k
            listShape.TextFrame2.TextRange.Font.Name = "+mn-lt"  ' theme body font
            For k = 1 To tr.Paragraphs.Count
                Set para = tr.Paragraphs(k)
                If k = s Then
                    para.Font.Bold = msoTrue
                    para.Font.Color.ObjectThemeColor = 5   ' msoThemeColorAccent1
                Else
                    para.Font.Bold = msoFalse
                    para.Font.Color.ObjectThemeColor = 1   ' msoThemeColorDark1
                    para.Font.Color.Brightness = 0.4       ' de-emphasise, theme-linked
                End If
            Next k
            built = built + 1
        End If
    Next s
    Inform built & " agenda/divider slide(s) built. Rerun this button after " & _
           "renaming, adding or reordering sections — dividers are tagged and " & _
           "rebuilt in place."
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub RemoveAgenda(control As Object)
    Const OP As String = "Remove agenda slides"
    On Error GoTo Oops
    Dim pres As Presentation, before As Long
    Set pres = ActivePres
    before = pres.Slides.Count
    DeleteTaggedAgendaSlides pres
    Inform (before - pres.Slides.Count) & " agenda slide(s) removed."
    Exit Sub
Oops:
    ReportError OP
End Sub
