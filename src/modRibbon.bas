Attribute VB_Name = "modRibbon"
Option Explicit

' ============================================================================
' modRibbon — shared ribbon plumbing.
' HARD RULE (see CLAUDE.md): every callback parameter is late-bound —
' (control As Object), never (control As IRibbonControl). IRibbonControl needs
' the Office object-library reference, which is unreliable in the target
' environment; Object always works.
'
' Feature callbacks (sizing, arrange, select, text, clean, audit) live in their
' own feature modules; this module holds only the onLoad handler and the shared
' IRibbonUI reference.
' ============================================================================

Public gRibbon As Object            ' IRibbonUI, late-bound

' ----------------------------------------------------------------------------
' Lifecycle
' ----------------------------------------------------------------------------

Public Sub Ribbon_OnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub
