Attribute VB_Name = "modClean"
Option Explicit

' ============================================================================
' modClean — Tier 2 deck housekeeping: speaker-notes removal, unused
' master/layout deletion, and the shape inspector read-out.
' Destructive operations always confirm first and report what they did.
' ============================================================================

Public Sub RemoveAllNotes(control As Object)
    Const OP As String = "Remove all speaker notes"
    On Error GoTo Oops
    Dim sld As Slide, shp As Shape, cleared As Long
    If MsgBox("Delete the speaker notes of EVERY slide in this presentation?" & _
              vbCrLf & "This cannot be undone.", vbOKCancel + vbExclamation, _
              APP_NAME) <> vbOK Then Exit Sub
    For Each sld In ActivePres.Slides
        For Each shp In sld.NotesPage.Shapes
            On Error Resume Next
            If shp.PlaceholderFormat.Type = ppPlaceholderBody Then
                If shp.HasTextFrame Then
                    If shp.TextFrame.HasText Then
                        shp.TextFrame.TextRange.Text = ""
                        cleared = cleared + 1
                    End If
                End If
            End If
            On Error GoTo Oops
        Next shp
    Next sld
    Inform "Speaker notes removed from " & cleared & " slide(s)."
    Exit Sub
Oops:
    ReportError OP
End Sub

' Delete slide masters (Designs) no slide uses, then unused custom layouts of
' the remaining masters. Reports exactly what was removed.
Public Sub DeleteUnusedMasters(control As Object)
    Const OP As String = "Delete unused masters/layouts"
    On Error GoTo Oops
    Dim pres As Presentation, sld As Slide
    Dim i As Long, j As Long, des As Design, lay As CustomLayout
    Dim usedDesign As Boolean, usedLayout As Boolean
    Dim removedDesigns As String, removedLayouts As Long, keptForError As Long
    Set pres = ActivePres
    If MsgBox("Delete every slide master and layout that no slide currently " & _
              "uses?" & vbCrLf & "A report of what was removed follows.", _
              vbOKCancel + vbExclamation, APP_NAME) <> vbOK Then Exit Sub
    ' 1) unused designs (slide masters), backwards so indices survive deletes
    For i = pres.Designs.Count To 1 Step -1
        Set des = pres.Designs(i)
        usedDesign = False
        For Each sld In pres.Slides
            If sld.Design Is des Then
                usedDesign = True
                Exit For
            End If
        Next sld
        If Not usedDesign And pres.Designs.Count > 1 Then
            removedDesigns = removedDesigns & "  - master: " & des.Name & vbCrLf
            des.Delete
        End If
    Next i
    ' 2) unused layouts inside the remaining designs
    For i = 1 To pres.Designs.Count
        Set des = pres.Designs(i)
        For j = des.SlideMaster.CustomLayouts.Count To 1 Step -1
            Set lay = des.SlideMaster.CustomLayouts(j)
            usedLayout = False
            For Each sld In pres.Slides
                If sld.CustomLayout Is lay Then
                    usedLayout = True
                    Exit For
                End If
            Next sld
            If Not usedLayout And des.SlideMaster.CustomLayouts.Count > 1 Then
                On Error Resume Next   ' preserved/protected layouts refuse deletion
                lay.Delete
                If Err.Number <> 0 Then
                    keptForError = keptForError + 1
                    Err.Clear
                Else
                    removedLayouts = removedLayouts + 1
                End If
                On Error GoTo Oops
            End If
        Next j
    Next i
    Inform "Cleanup done." & vbCrLf & vbCrLf & _
           IIf(Len(removedDesigns) > 0, "Masters removed:" & vbCrLf & removedDesigns, _
               "No unused masters found." & vbCrLf) & vbCrLf & _
           removedLayouts & " unused layout(s) removed." & _
           IIf(keptForError > 0, vbCrLf & keptForError & _
               " layout(s) could not be removed (protected).", "")
    Exit Sub
Oops:
    ReportError OP
End Sub

' Inspector: position/size read-out of the reference shape, in cm.
' Doubles as a debugging aid.
Public Sub InspectShape(control As Object)
    Const OP As String = "Inspect shape"
    On Error GoTo Oops
    Dim sr As ShapeRange, shp As Shape, msg As String, fillDesc As String
    Set sr = GuardShapes(1, OP)
    If sr Is Nothing Then Exit Sub
    Set shp = RefShape(sr)
    On Error Resume Next
    If shp.Fill.Visible = msoTrue Then
        If shp.Fill.ForeColor.ObjectThemeColor <> msoNotThemeColor Then
            fillDesc = ThemeColorName(shp.Fill.ForeColor.ObjectThemeColor) & _
                       " (theme-linked, brightness " & _
                       Format(shp.Fill.ForeColor.Brightness, "0.00") & ")"
        Else
            fillDesc = "fixed RGB #" & Right$("000000" & Hex(RGBToHexOrder( _
                       shp.Fill.ForeColor.RGB)), 6) & "  (does NOT track the theme)"
        End If
    Else
        fillDesc = "none"
    End If
    On Error GoTo Oops
    msg = "Name:      " & shp.Name & vbCrLf & _
          "X (left):  " & FmtCm(shp.Left) & vbCrLf & _
          "Y (top):   " & FmtCm(shp.Top) & vbCrLf & _
          "Width:     " & FmtCm(shp.Width) & vbCrLf & _
          "Height:    " & FmtCm(shp.Height) & vbCrLf & _
          "Rotation:  " & Format(shp.Rotation, "0.0") & Chr$(176) & vbCrLf & _
          "Fill:      " & fillDesc
    If sr.Count > 1 Then msg = msg & vbCrLf & vbCrLf & "(Read-out is for the " & _
                               "reference = first-selected shape.)"
    Inform msg
    Exit Sub
Oops:
    ReportError OP
End Sub

' VBA's RGB() longs are BGR in memory; reorder for human-readable hex.
Public Function RGBToHexOrder(ByVal v As Long) As Long
    RGBToHexOrder = (v And &HFF&) * &H10000 + (v And &HFF00&) + _
                    ((v \ &H10000) And &HFF&)
End Function
