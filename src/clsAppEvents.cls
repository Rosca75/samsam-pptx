VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================================
' clsAppEvents — Application-events plumbing for the Tier 3 sticky painter.
' DELIBERATELY MINIMAL: catch the event, hand off to modPainter, never raise.
' This class (with modPainter) is the quarantined fragile zone — nothing else
' in the add-in may depend on it. Import via VBA editor File > Import File…
' exactly like the .bas modules.
' ============================================================================

Public WithEvents App As Application

Private Sub App_WindowSelectionChange(ByVal Sel As Selection)
    On Error Resume Next   ' event sink must never leak an error into PowerPoint
    modPainter.OnSelectionChanged Sel
End Sub
