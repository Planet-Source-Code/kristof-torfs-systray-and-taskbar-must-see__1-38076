VERSION 5.00
Begin VB.Form frmTray 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Windows Sytem Tray"
   ClientHeight    =   945
   ClientLeft      =   150
   ClientTop       =   390
   ClientWidth     =   1845
   Icon            =   "frmTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   1845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox sysTray 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1515
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   480
      Width           =   1575
      Begin VB.Timer tmrSysTray 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   1080
         Top             =   0
      End
      Begin VB.PictureBox picIcon 
         BorderStyle     =   0  'None
         Height          =   375
         Index           =   0
         Left            =   1200
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Image imgTrayIcon 
         Height          =   255
         Index           =   0
         Left            =   -255
         Stretch         =   -1  'True
         Top             =   -50
         Width           =   255
      End
   End
   Begin VB.Label lblTray 
      AutoSize        =   -1  'True
      Caption         =   "Windows System Tray:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1620
   End
End
Attribute VB_Name = "frmTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrSysTray_Timer()
'Update Tray
Dim dtrLeft As Long
Dim iconCount As Long
For iconCount = 1 To imgTrayIcon.Count - 1
    If imgTrayIcon(iconCount).Tag <> "skip" Then dtrLeft = dtrLeft + 300
Next iconCount
End Sub
