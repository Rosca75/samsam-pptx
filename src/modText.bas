Attribute VB_Name = "modText"
Option Explicit

' ============================================================================
' modText — Tier 2 text hygiene: deck-wide proofing language, whitespace
' cleanup, empty text-box removal, margins and word-wrap toggles.
' Deck-wide operations cover slides, notes pages, masters and layouts.
' Text edits use TextRange.Replace / character deletes — NEVER a wholesale
' .Text assignment, which would destroy run formatting.
' ============================================================================

' PowerPoint's default internal margins (cm) for the "normal" margin preset.
Private Const MARGIN_LR_CM As Double = 0.25
Private Const MARGIN_TB_CM As Double = 0.13

' ----------------------------------------------------------------------------
' Proofing language, whole deck
' ----------------------------------------------------------------------------

Private Function LanguageIdFromCode(ByVal code As String) As Long
    Select Case UCase$(Trim$(code))
        Case "EN", "EN-US":       LanguageIdFromCode = msoLanguageIDEnglishUS
        Case "EN-GB", "UK":       LanguageIdFromCode = msoLanguageIDEnglishUK
        Case "FR", "FR-FR":       LanguageIdFromCode = msoLanguageIDFrench
        Case "FR-LU":             LanguageIdFromCode = msoLanguageIDFrenchLuxembourg
        Case "DE", "DE-DE":       LanguageIdFromCode = msoLanguageIDGerman
        Case "DE-LU":             LanguageIdFromCode = msoLanguageIDGermanLuxembourg
        Case "NL":                LanguageIdFromCode = msoLanguageIDDutch
        Case "IT":                LanguageIdFromCode = msoLanguageIDItalian
        Case "ES":                LanguageIdFromCode = msoLanguageIDSpanishModernSort
        Case "PT":                LanguageIdFromCode = msoLanguageIDPortuguese
        Case Else:                LanguageIdFromCode = 0
    End Select
End Function

Private Sub SetLanguageOnShapes(ByVal shapesColl As Shapes, ByVal langId As Long, _
                                ByRef count As Long)
    Dim shp As Shape
    For Each shp In LeafShapes(shapesColl)
        On Error Resume Next
        If shp.HasTextFrame Then
            shp.TextFrame.TextRange.LanguageID = langId
            count = count + 1
        End If
        On Error GoTo 0
    Next shp
End Sub

Public Sub SetDeckLanguage(control As Object)
    Const OP As String = "Set proofing language (whole deck)"
    On Error GoTo Oops
    Dim code As String, langId As Long, count As Long
    Dim sld As Slide, des As Design, lay As CustomLayout
    code = InputBox("Set the proofing language for the ENTIRE deck " & _
                    "(all slides, notes, masters, layouts)." & vbCrLf & vbCrLf & _
                    "Codes: EN-US, EN-GB, FR, FR-LU, DE, DE-LU, NL, IT, ES, PT", _
                    APP_NAME, "EN-GB")
    If Len(Trim$(code)) = 0 Then Exit Sub
    langId = LanguageIdFromCode(code)
    If langId = 0 Then
        MsgBox "Unknown language code '" & code & "'. Use one of: " & _
               "EN-US, EN-GB, FR, FR-LU, DE, DE-LU, NL, IT, ES, PT", vbExclamation, APP_NAME
        Exit Sub
    End If
    For Each sld In ActivePres.Slides
        SetLanguageOnShapes sld.Shapes, langId, count
        SetLanguageOnShapes sld.NotesPage.Shapes, langId, count
    Next sld
    For Each des In ActivePres.Designs
        SetLanguageOnShapes des.SlideMaster.Shapes, langId, count
        For Each lay In des.SlideMaster.CustomLayouts
            SetLanguageOnShapes lay.Shapes, langId, count
        Next lay
    Next des
    ' also set the presentation default so NEW text boxes follow
    On Error Resume Next
    ActivePres.DefaultLanguageID = langId
    On Error GoTo Oops
    Inform "Proofing language set to " & UCase$(Trim$(code)) & " on " & count & _
           " text frames (slides, notes, masters, layouts)."
    Exit Sub
Oops:
    ReportError OP
End Sub

' ----------------------------------------------------------------------------
' Whitespace cleanup + empty text boxes, deck-wide
' ----------------------------------------------------------------------------

' Removes runs of multiple spaces and trailing spaces without touching run
' formatting. Returns number of shapes changed.
Private Function CleanShapeText(ByVal shp As Shape) As Boolean
    Dim tr As TextRange, para As TextRange, s As String, trail As Long
    Dim i As Long, guard As Long
    If Not shp.HasTextFrame Then Exit Function
    If Not shp.TextFrame.HasText Then Exit Function
    Set tr = shp.TextFrame.TextRange
    ' double spaces -> single (Replace preserves formatting)
    guard = 0
    Do While InStr(tr.Text, "  ") > 0 And guard < 500
        tr.Replace "  ", " ", WholeWords:=False
        guard = guard + 1
        CleanShapeText = True
    Loop
    ' trailing spaces per paragraph
    For i = 1 To tr.Paragraphs.Count
        Set para = tr.Paragraphs(i)
        s = para.Text
        If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        trail = 0
        Do While trail < Len(s) And Mid$(s, Len(s) - trail, 1) = " "
            trail = trail + 1
        Loop
        If trail > 0 Then
            para.Characters(Len(s) - trail + 1, trail).Delete
            CleanShapeText = True
        End If
    Next i
End Function

Public Sub CleanDeckText(control As Object)
    Const OP As String = "Clean text (deck-wide)"
    On Error GoTo Oops
    Dim sld As Slide, shp As Shape, changed As Long
    If MsgBox("Remove double spaces and trailing spaces from ALL text on ALL " & _
              "slides?", vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    For Each sld In ActivePres.Slides
        For Each shp In LeafShapes(sld.Shapes)
            On Error Resume Next
            If CleanShapeText(shp) Then changed = changed + 1
            On Error GoTo Oops
        Next shp
    Next sld
    Inform "Done — cleaned text in " & changed & " shape(s)."
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub DeleteEmptyTextBoxes(control As Object)
    Const OP As String = "Delete empty text boxes"
    On Error GoTo Oops
    Dim sld As Slide, i As Long, shp As Shape, deleted As Long
    If MsgBox("Delete every empty text box on every slide?" & vbCrLf & _
              "(Only plain text boxes are removed — placeholders and shapes stay.)", _
              vbOKCancel + vbQuestion, APP_NAME) <> vbOK Then Exit Sub
    For Each sld In ActivePres.Slides
        For i = sld.Shapes.Count To 1 Step -1
            Set shp = sld.Shapes(i)
            If shp.Type = msoTextBox Then
                If Not ShapeHasText(shp) Or Len(Trim$(Replace(Replace( _
                   shp.TextFrame.TextRange.Text, vbCr, ""), vbLf, ""))) = 0 Then
                    shp.Delete
                    deleted = deleted + 1
                End If
            End If
        Next i
    Next sld
    Inform deleted & " empty text box(es) deleted."
    Exit Sub
Oops:
    ReportError OP
End Sub

' ----------------------------------------------------------------------------
' Margins and word wrap on the selection
' ----------------------------------------------------------------------------

' Tag-dispatched: "NONE" | "NORMAL".
Public Sub SetTextMargins(control As Object)
    Const OP As String = "Set text margins"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long, tf As TextFrame
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        If sr.Item(i).HasTextFrame Then
            Set tf = sr.Item(i).TextFrame
            If control.Tag = "NONE" Then
                tf.MarginLeft = 0: tf.MarginRight = 0
                tf.MarginTop = 0:  tf.MarginBottom = 0
            Else
                tf.MarginLeft = CmToPt(MARGIN_LR_CM): tf.MarginRight = CmToPt(MARGIN_LR_CM)
                tf.MarginTop = CmToPt(MARGIN_TB_CM):  tf.MarginBottom = CmToPt(MARGIN_TB_CM)
            End If
        End If
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub

Public Sub ToggleWordWrap(control As Object)
    Const OP As String = "Toggle word wrap"
    On Error GoTo Oops
    Dim sr As ShapeRange, i As Long, tf As TextFrame
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    For i = 1 To sr.Count
        If sr.Item(i).HasTextFrame Then
            Set tf = sr.Item(i).TextFrame
            tf.WordWrap = Not tf.WordWrap
        End If
    Next i
    Exit Sub
Oops:
    ReportError OP
End Sub
