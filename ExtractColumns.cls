VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExtractColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.CommandBars(2).FindControl(, 30007).Controls(gstrcAPPMENU).Delete
End Sub

Private Sub Workbook_Open()
    Dim ctlToolsMenu As Office.CommandBarButton

    'Add Shortcut to Add-In Menu/Ribbon
    Set ctlToolsMenu = Application.CommandBars(2).FindControl(, 30007).Controls.Add( _
        msoControlButton)
    ctlToolsMenu.Caption = gstrcAPPMENU
    ctlToolsMenu.OnAction = "ShowForm"
End Sub


