VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents App As Application

Private Sub Workbook_Open()
    StartCrosshairAddin
End Sub

Private Sub Workbook_AddinInstall()
    StartCrosshairAddin
End Sub

Private Sub Workbook_AddinUninstall()
    TerminateCrosshair
    Set App = Nothing
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    TerminateCrosshair
    Set App = Nothing
End Sub

Private Sub StartCrosshairAddin()
    Set App = Application
    InitializeCrosshair
End Sub

Private Sub App_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    Crosshair_HandleSelectionChange Sh, Target
End Sub

Private Sub App_SheetActivate(ByVal Sh As Object)
    Crosshair_HandleSheetActivate Sh
End Sub

Private Sub App_SheetDeactivate(ByVal Sh As Object)
    Crosshair_HandleSheetDeactivate Sh
End Sub

Private Sub App_WorkbookBeforeClose(ByVal Wb As Workbook, Cancel As Boolean)
    Crosshair_HandleWorkbookBeforeClose Wb
End Sub

Private Sub App_WorkbookBeforeSave(ByVal Wb As Workbook, ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Crosshair_HandleWorkbookBeforeSave Wb
End Sub
