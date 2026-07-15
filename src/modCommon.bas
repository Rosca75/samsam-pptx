Attribute VB_Name = "modCommon"
Option Explicit

' ============================================================================
' modCommon — selection guards and shared helpers. BUILD/IMPORT THIS FIRST.
' Every other module depends on it. Conventions (see CLAUDE.md):
'   * first-selected shape is the reference shape
'   * every entry point guards its selection via GuardShapes — no silent
'     failures, no unhandled errors
'   * user-facing distances are cm; the object model uses points
' NOTE: all module-level declarations sit above the first procedure — VBA
' rejects declarations placed after any Sub/Function.
' ============================================================================

Public Const APP_NAME As String = "SamSam Tools"
Public Const PT_PER_CM As Double = 28.3464566929134
' Sentinel returned by AskNumber when the user cancels the input box.
Public Const NUM_CANCELLED As Double = -1E+99

' ----------------------------------------------------------------------------
' Selection guards
' ----------------------------------------------------------------------------

' Returns the selected ShapeRange, or Nothing after a clear MsgBox when the
' selection has fewer than minCount shapes (or is not a shape selection).
' A text-cursor selection inside a shape counts as that one shape.
Public Function GuardShapes(ByVal minCount As Long, ByVal opName As String) As ShapeRange
    Dim sr As ShapeRange
    On Error GoTo NoSelection
    If ActiveWindow.Selection.Type = ppSelectionNone Then GoTo NoSelection
    Set sr = ActiveWindow.Selection.ShapeRange
    If sr.Count < minCount Then
        MsgBox opName & " needs at least " & minCount & _
               IIf(minCount = 1, " shape", " shapes") & " selected." & vbCrLf & vbCrLf & _
               IIf(minCount >= 2, "Reminder: the FIRST shape you select is the reference.", ""), _
               vbExclamation, APP_NAME
        Set GuardShapes = Nothing
        Exit Function
    End If
    Set GuardShapes = sr
    Exit Function
NoSelection:
    MsgBox opName & ": select " & IIf(minCount = 1, "a shape", minCount & " or more shapes") & _
           " first.", vbExclamation, APP_NAME
    Set GuardShapes = Nothing
End Function

' The reference shape = first shape of the selection.
Public Function RefShape(ByVal sr As ShapeRange) As Shape
    Set RefShape = sr.Item(1)
End Function

' Slide that owns the current selection/view; Nothing outside normal view.
Public Function CurrentSlide() As Slide
    On Error Resume Next
    Set CurrentSlide = ActiveWindow.View.Slide
End Function

Public Function ActivePres() As Presentation
    Set ActivePres = Application.ActivePresentation
End Function

' ----------------------------------------------------------------------------
' Units and numeric input
' ----------------------------------------------------------------------------

Public Function CmToPt(ByVal cm As Double) As Double
    CmToPt = cm * PT_PER_CM
End Function

Public Function PtToCm(ByVal pt As Double) As Double
    PtToCm = pt / PT_PER_CM
End Function

Public Function FmtCm(ByVal pt As Double) As String
    FmtCm = Format(PtToCm(pt), "0.00") & " cm"
End Function

' InputBox for a number. Accepts comma or dot decimals (FR/DE keyboards).
' Returns NUM_CANCELLED when the user cancels or enters nothing.
Public Function AskNumber(ByVal prompt As String, ByVal defaultValue As String) As Double
    Dim raw As String
    raw = InputBox(prompt, APP_NAME, defaultValue)
    If Len(Trim$(raw)) = 0 Then
        AskNumber = NUM_CANCELLED
        Exit Function
    End If
    raw = Replace(Trim$(raw), ",", ".")
    If Not IsNumeric(raw) Then
        MsgBox "'" & raw & "' is not a number.", vbExclamation, APP_NAME
        AskNumber = NUM_CANCELLED
        Exit Function
    End If
    AskNumber = Val(raw)
End Function

' Standard error reporter used by every entry point's error handler.
Public Sub ReportError(ByVal opName As String)
    MsgBox opName & " failed: " & Err.Description & " (error " & Err.Number & ")", _
           vbCritical, APP_NAME
End Sub

Public Sub Inform(ByVal msg As String)
    MsgBox msg, vbInformation, APP_NAME
End Sub

' ----------------------------------------------------------------------------
' Theme access (flagship colour feature support)
' ----------------------------------------------------------------------------

' RGB of theme colour slot idx (1..12: dark1, light1, dark2, light2,
' accent1..6, hyperlink, followed) read from the ACTIVE presentation's first
' slide master — the live scheme the galleries must mirror.
Public Function ThemeRGB(ByVal idx As Long) As Long
    ThemeRGB = ActivePres.Designs(1).SlideMaster.Theme.ThemeColorScheme.Colors(idx).RGB
End Function

' MsoThemeColorIndex for ObjectThemeColor, matching ThemeRGB's slot order.
' (Scheme slots 1..10 map to msoThemeColorDark1=1 .. msoThemeColorAccent6=10.)
Public Function ThemeObjectColor(ByVal idx As Long) As MsoThemeColorIndex
    ThemeObjectColor = idx
End Function

Public Function ThemeColorName(ByVal idx As Long) As String
    Select Case idx
        Case 1: ThemeColorName = "Dark 1 (Text)"
        Case 2: ThemeColorName = "Light 1 (Background)"
        Case 3: ThemeColorName = "Dark 2"
        Case 4: ThemeColorName = "Light 2"
        Case 5 To 10: ThemeColorName = "Accent " & (idx - 4)
        Case 11: ThemeColorName = "Hyperlink"
        Case 12: ThemeColorName = "Followed hyperlink"
        Case Else: ThemeColorName = "Theme colour " & idx
    End Select
End Function

' ----------------------------------------------------------------------------
' Shape iteration helpers
' ----------------------------------------------------------------------------

' Collects shp plus, recursively, group members and table cell shapes into col.
' Used by deck-wide operations (language, cleanup, audit).
Public Sub CollectLeafShapes(ByVal shp As Shape, ByVal col As Collection)
    Dim child As Shape, r As Long, c As Long
    If shp.Type = msoGroup Then
        For Each child In shp.GroupItems
            CollectLeafShapes child, col
        Next child
    ElseIf shp.HasTable Then
        col.Add shp
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                col.Add shp.Table.Cell(r, c).Shape
            Next c
        Next r
    Else
        col.Add shp
    End If
End Sub

' All leaf shapes of a Shapes collection (slide, notes page, master, layout).
Public Function LeafShapes(ByVal shapesColl As Shapes) As Collection
    Dim col As New Collection, shp As Shape
    For Each shp In shapesColl
        CollectLeafShapes shp, col
    Next shp
    Set LeafShapes = col
End Function

Public Function ShapeHasText(ByVal shp As Shape) As Boolean
    On Error Resume Next
    If shp.HasTextFrame Then ShapeHasText = shp.TextFrame.HasText
End Function

' Temporarily clear LockAspectRatio while resizing; returns the previous state.
Public Function UnlockAspect(ByVal shp As Shape) As Boolean
    UnlockAspect = (shp.LockAspectRatio = msoTrue)
    shp.LockAspectRatio = msoFalse
End Function

Public Sub RestoreAspect(ByVal shp As Shape, ByVal wasLocked As Boolean)
    If wasLocked Then shp.LockAspectRatio = msoTrue
End Sub
