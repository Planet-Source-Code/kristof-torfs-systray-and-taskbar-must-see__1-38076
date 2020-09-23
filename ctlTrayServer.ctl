VERSION 5.00
Begin VB.UserControl ctlTrayServer 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
End
Attribute VB_Name = "ctlTrayServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private ServerRunning As Boolean

Public Enum EventTypes      ' all events which may happen, notified by parent
  DoubleClick
  MouseDown
  MouseUp
  MouseMove
End Enum

Public Sub StartServer()        ' listen for incoming commands
    Call LoadTrayIconHandler
    frmTray.Visible = False
    ServerRunning = True
    frmTray.tmrSysTray.Enabled = True
End Sub

Public Sub StopServer()             ' stop listening
    Call UnLoadTrayIconHandler
    frmTray.Visible = False
    ServerRunning = False
    frmTray.tmrSysTray.Enabled = False
End Sub

Public Property Get ItemCount() As Integer      ' Number of system tray icons
    If ServerRunning = False Then Err.Raise 1, , "System Tray Server is not running!!!"
    ItemCount = frmTray.imgTrayIcon.Count
End Property

Public Property Get ItemIcon(Index As Integer) As IPictureDisp      ' The icon for the item
On Error Resume Next
    If ServerRunning = False Then Err.Raise 1, , "System Tray Server is not running!!!"
    Set ItemIcon = frmTray.imgTrayIcon(Index).Picture
End Property

Public Property Get ItemToolTip(Index As Integer) As String         ' The tooltip for the item
On Error Resume Next
    If ServerRunning = False Then Err.Raise 1, , "System Tray Server is not running!!!"
    ItemToolTip = frmTray.imgTrayIcon(Index).ToolTipText
End Property

Public Sub ItemRaiseEvent(Index As Integer, EventType As EventTypes, Button As Integer)     ' Raise an event and make Windows react to it
    If ServerRunning = False Then Err.Raise 1, , "System Tray Server is not running!!!"
    Select Case EventType
        Case DoubleClick: trayClick WM_LBUTTONDBLCLK, Index
        Case MouseMove: trayClick WM_MOUSEMOVE, Index
        Case MouseDown: trayClick IIf(Button = 1, WM_LBUTTONDOWN, WM_RBUTTONDOWN), Index
        Case MouseUp: trayClick IIf(Button = 1, WM_LBUTTONUP, WM_RBUTTONUP), Index
    End Select
End Sub

Public Sub FreeTray()       ' Make Tray lose focus. Do this everytime something else than the tray is focused
    If ServerRunning = False Then Err.Raise 1, , "System Tray Server is not running!!!"
    freeTrayObjects
End Sub

Public Sub About()          ' Display an aboutbox
Attribute About.VB_UserMemId = -552
  frmAbout.Show vbModal
End Sub

Private Sub UserControl_Initialize()        ' Visual effect
  Picture = frmTray.Icon
End Sub

Private Sub UserControl_Resize()        ' Make control have same size at all times
  Width = 480
  Height = 480
End Sub
