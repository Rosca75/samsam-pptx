Attribute VB_Name = "modRibbon"
Option Explicit

' ============================================================================
' modRibbon — ALL ribbon callbacks live here.
' HARD RULE (see CLAUDE.md): every callback parameter is late-bound —
' (control As Object), never (control As IRibbonControl). IRibbonControl needs
' the Office object-library reference, which is unreliable in the target
' environment; Object always works.
'
' Theme-colour galleries: 10 columns (dark1, light1, dark2, light2,
' accent1..6) x 6 rows (base + lighter 80/60/40 + darker 25/50) = 60 items.
' Swatch images are 24-bit BMPs generated on the fly from the LIVE theme and
' loaded with stdole.LoadPicture — no network, no asset files, no pip.
' The swatch RGB is calculated from the theme: that calculation is the
' documented exception to the "no RGB" rule — it only paints PREVIEW pixels;
' the click itself applies ObjectThemeColor + Brightness (modColor).
'
' PowerPoint caches gallery images: after switching to a presentation with a
' different theme, click "Refresh palette" (invalidates the ribbon).
' ============================================================================

Private Const PALETTE_COLS As Long = 10
Private Const PALETTE_ROWS As Long = 6
Private Const SWATCH_PX As Long = 32

Public gRibbon As Object            ' IRibbonUI, late-bound

' Align-to-slide vs align-to-selection toggle state (default: slide).
Private mAlignToSlide As Boolean
Private mAlignInit As Boolean

' ----------------------------------------------------------------------------
' Lifecycle
' ----------------------------------------------------------------------------

Public Sub Ribbon_OnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub RefreshPalette(control As Object)
    On Error GoTo Oops
    If gRibbon Is Nothing Then
        MsgBox "Ribbon reference lost (this can happen after an unhandled VBA " & _
               "error reset the project). Restart PowerPoint to restore it.", _
               vbExclamation, APP_NAME
        Exit Sub
    End If
    gRibbon.Invalidate
    Exit Sub
Oops:
    ReportError "Refresh palette"
End Sub

' ----------------------------------------------------------------------------
' Align mode toggle
' ----------------------------------------------------------------------------

Public Function AlignToSlideMode() As Boolean
    If Not mAlignInit Then
        mAlignToSlide = True
        mAlignInit = True
    End If
    AlignToSlideMode = mAlignToSlide
End Function

Public Sub GetAlignPressed(control As Object, ByRef returnedVal)
    returnedVal = AlignToSlideMode()
End Sub

Public Sub OnAlignToggle(control As Object, pressed As Boolean)
    AlignToSlideMode ' ensure initialised
    mAlignToSlide = pressed
End Sub

' ----------------------------------------------------------------------------
' Sticky painter toggle (delegates to modPainter — fragile code stays there)
' ----------------------------------------------------------------------------

Public Sub GetPainterPressed(control As Object, ByRef returnedVal)
    returnedVal = modPainter.StickyActive()
End Sub

Public Sub OnPainterToggle(control As Object, pressed As Boolean)
    modPainter.SetSticky pressed
End Sub

' ----------------------------------------------------------------------------
' Theme-colour galleries (controls tagged "fill" | "line" | "text" | "both")
' ----------------------------------------------------------------------------

Public Sub GetPaletteCount(control As Object, ByRef returnedVal)
    returnedVal = PALETTE_COLS * PALETTE_ROWS
End Sub

Public Sub GetPaletteItemScreentip(control As Object, index As Integer, ByRef returnedVal)
    On Error Resume Next
    returnedVal = ThemeColorName((index Mod PALETTE_COLS) + 1) & _
                  VariantName(index \ PALETTE_COLS)
End Sub

Public Sub GetPaletteItemImage(control As Object, index As Integer, ByRef returnedVal)
    On Error GoTo Oops
    Dim slot As Long, row As Long, baseRGB As Long
    slot = (index Mod PALETTE_COLS) + 1
    row = index \ PALETTE_COLS
    baseRGB = ThemeRGB(slot)
    Set returnedVal = SwatchPicture(ApplyBrightnessRGB(baseRGB, VariantBrightness(row)))
    Exit Sub
Oops:
    ' Never let a swatch failure kill the whole gallery — return no image.
    Set returnedVal = Nothing
End Sub

Public Sub OnPaletteAction(control As Object, itemId As String, index As Integer)
    modColor.ApplyThemeVariant CStr(control.Tag), _
                               (index Mod PALETTE_COLS) + 1, _
                               index \ PALETTE_COLS
End Sub

' ----------------------------------------------------------------------------
' Swatch bitmap generation (preview pixels only — see header comment)
' ----------------------------------------------------------------------------

' Mirror of ColorFormat.Brightness: >0 blends towards white, <0 towards black.
Public Function ApplyBrightnessRGB(ByVal rgbVal As Long, ByVal brightness As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = rgbVal And &HFF&
    g = (rgbVal \ &H100&) And &HFF&
    b = (rgbVal \ &H10000) And &HFF&
    If brightness > 0 Then
        r = r + (255 - r) * brightness
        g = g + (255 - g) * brightness
        b = b + (255 - b) * brightness
    ElseIf brightness < 0 Then
        r = r * (1 + brightness)
        g = g * (1 + brightness)
        b = b * (1 + brightness)
    End If
    ApplyBrightnessRGB = RGB(ClampByte(r), ClampByte(g), ClampByte(b))
End Function

Private Function ClampByte(ByVal v As Long) As Long
    If v < 0 Then v = 0
    If v > 255 Then v = 255
    ClampByte = v
End Function

' Writes a SWATCH_PX x SWATCH_PX 24-bit BMP (with a grey 1px border so light
' colours stay visible) to %TEMP% and returns it as an IPictureDisp.
Private Function SwatchPicture(ByVal rgbVal As Long) As Object
    Dim path As String
    path = Environ$("TEMP") & "\samsam_swatch.bmp"
    WriteSwatchBmp path, rgbVal
    Set SwatchPicture = stdole.LoadPicture(path)
End Function

Private Sub WriteSwatchBmp(ByVal path As String, ByVal rgbVal As Long)
    Dim rowBytes As Long, dataSize As Long, fileSize As Long
    Dim bytes() As Byte, pos As Long, x As Long, y As Long
    Dim r As Byte, g As Byte, b As Byte, edge As Boolean
    Dim fnum As Integer

    rowBytes = SWATCH_PX * 3            ' 96 — already a multiple of 4, no padding
    dataSize = rowBytes * SWATCH_PX
    fileSize = 54 + dataSize
    ReDim bytes(0 To fileSize - 1)

    ' BITMAPFILEHEADER
    bytes(0) = Asc("B"): bytes(1) = Asc("M")
    PutLong bytes, 2, fileSize
    PutLong bytes, 10, 54
    ' BITMAPINFOHEADER
    PutLong bytes, 14, 40
    PutLong bytes, 18, SWATCH_PX
    PutLong bytes, 22, SWATCH_PX
    bytes(26) = 1                        ' planes
    bytes(28) = 24                       ' bits per pixel
    PutLong bytes, 34, dataSize
    PutLong bytes, 38, 2835              ' 72 dpi
    PutLong bytes, 42, 2835

    r = rgbVal And &HFF&
    g = (rgbVal \ &H100&) And &HFF&
    b = (rgbVal \ &H10000) And &HFF&

    pos = 54
    For y = 0 To SWATCH_PX - 1           ' BMP rows are bottom-up; irrelevant for a flat swatch
        For x = 0 To SWATCH_PX - 1
            edge = (x = 0 Or y = 0 Or x = SWATCH_PX - 1 Or y = SWATCH_PX - 1)
            If edge Then                 ' grey border
                bytes(pos) = 160: bytes(pos + 1) = 160: bytes(pos + 2) = 160
            Else                         ' BGR order
                bytes(pos) = b: bytes(pos + 1) = g: bytes(pos + 2) = r
            End If
            pos = pos + 3
        Next x
    Next y

    fnum = FreeFile
    Open path For Binary Access Write As #fnum
    Put #fnum, 1, bytes
    Close #fnum
End Sub

Private Sub PutLong(ByRef bytes() As Byte, ByVal offset As Long, ByVal v As Long)
    bytes(offset) = v And &HFF&
    bytes(offset + 1) = (v \ &H100&) And &HFF&
    bytes(offset + 2) = (v \ &H10000) And &HFF&
    bytes(offset + 3) = (v \ &H1000000) And &HFF&
End Sub
