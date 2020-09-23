VERSION 5.00
Object = "{EB7F05E0-B2B2-11D6-A90A-00C026F0773F}#3.0#0"; "vbTray.ocx"
Begin VB.Form frmDemo 
   Caption         =   "Form1"
   ClientHeight    =   2790
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin vbTray.ctlTaskBar ctlTaskBar1 
      Height          =   480
      Left            =   120
      TabIndex        =   13
      Top             =   2280
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin vbTray.ctlTrayServer ctlTrayServer1 
      Height          =   480
      Left            =   4080
      TabIndex        =   12
      Top             =   2280
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tray"
      Height          =   2055
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   735
      Begin VB.PictureBox Picture6 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   10
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox Picture5 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   9
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture4 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   8
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.PictureBox Picture3 
      Height          =   495
      Left            =   240
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   1560
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tasks"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.PictureBox Picture2 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox Picture1 
         Height          =   495
         Left            =   120
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Label1"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1560
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Label1"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   960
         Width           =   2775
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   2775
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2400
      Top             =   360
   End
   Begin VB.Label Label4 
      Caption         =   "This demo only lists 3 tasks and 3 tray items. It's only to give you a clue about how it works."
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   2280
      Width           =   3255
   End
   Begin VB.Menu mnuExplorer 
      Caption         =   "&Explore hard disk"
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  ctlTrayServer1.StartServer
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.FreeTray
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.FreeTray       ' this should always be done when tray loses focus...
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ctlTrayServer1.FreeTray
  ctlTrayServer1.StopServer
End Sub

Private Sub mnuExplorer_Click()
  Shell "explorer.exe", vbNormalFocus
End Sub

Private Sub Picture4_DblClick()
  ctlTrayServer1.ItemRaiseEvent 1, DoubleClick, 0
End Sub

Private Sub Picture5_DblClick()
  ctlTrayServer1.ItemRaiseEvent 2, DoubleClick, 0
End Sub

Private Sub Picture6_DblClick()
  ctlTrayServer1.ItemRaiseEvent 3, DoubleClick, 0
End Sub

Private Sub Picture4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 1, MouseUp, Button
End Sub

Private Sub Picture5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 2, MouseUp, Button
End Sub

Private Sub Picture6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 3, MouseUp, Button
End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 1, MouseDown, Button
End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 2, MouseDown, Button
End Sub

Private Sub Picture6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 3, MouseDown, Button
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 1, MouseMove, Button
End Sub

Private Sub Picture5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 2, MouseMove, Button
End Sub

Private Sub Picture6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  ctlTrayServer1.ItemRaiseEvent 3, MouseMove, Button
End Sub

Private Sub Timer1_Timer()
  ' Get current tasks
  ctlTaskBar1.Refresh
  
  ' Draw the first one and set it's label
  ctlTaskBar1.DrawIcon ctlTaskBar1.ItemIcon(0), Picture1.hDC, 0, 0, 16, 16
  Label1.Caption = ctlTaskBar1.ItemText(0)
  
  ' Draw the second one and set it's label
  ctlTaskBar1.DrawIcon ctlTaskBar1.ItemIcon(1), Picture2.hDC, 0, 0, 16, 16
  Label2.Caption = ctlTaskBar1.ItemText(1)
  
  ' Draw the third one and set it's label
  ctlTaskBar1.DrawIcon ctlTaskBar1.ItemIcon(2), Picture3.hDC, 0, 0, 16, 16
  Label3.Caption = ctlTaskBar1.ItemText(2)
  
  ' Draw the tray items
  Picture4.Picture = ctlTrayServer1.ItemIcon(1)
  Picture5.Picture = ctlTrayServer1.ItemIcon(2)
  Picture6.Picture = ctlTrayServer1.ItemIcon(3)
  
  ' Set their tooltips
  Picture4.ToolTipText = ctlTrayServer1.ItemToolTip(1)
  Picture5.ToolTipText = ctlTrayServer1.ItemToolTip(2)
  Picture6.ToolTipText = ctlTrayServer1.ItemToolTip(3)
End Sub
