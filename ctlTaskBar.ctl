VERSION 5.00
Begin VB.UserControl ctlTaskBar 
   ClientHeight    =   465
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   660
   Picture         =   "ctlTaskBar.ctx":0000
   ScaleHeight     =   465
   ScaleWidth      =   660
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4695
   End
End
Attribute VB_Name = "ctlTaskBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Const IDI_WINLOGO = 32517

Private Const DI_MASK = &H1
Private Const DI_IMAGE = &H2
Private Const DI_NORMAL = DI_MASK Or DI_IMAGE

Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Const GCL_HICON = (-14)
Private Const GCL_HICONSM = (-34)

Public Sub Refresh()
    fEnumWindows List1, UserControl.ContainerHwnd
End Sub

Public Property Get ItemCount() As Integer
    ItemCount = List1.ListCount
End Property

Public Property Get ItemText(Index) As String
    If Index < List1.ListCount Then ItemText = Left(List1.List(Index), InStr(1, List1.List(Index), vbCrLf) - 1)
End Property

Public Property Get ItemHwnd(Index) As Long
    If Index < List1.ListCount Then ItemHwnd = Right(List1.List(Index), Len(List1.List(Index)) - InStr(1, List1.List(Index), vbCrLf))
End Property

Public Property Get ForeGroundWindowIndex() As Integer
Dim tmp As Long
    ForeGroundWindowIndex = -1
    For i = 0 To List1.ListCount - 1
        If GetForegroundWindow = ItemHwnd(i) Then ForeGroundWindowIndex = i
    Next i
End Property

Public Property Let ForeGroundWindow(Index)
    pSetForegroundWindow ItemHwnd(Index)
End Property

Public Property Get ItemIcon(Index) As Long
Dim hIcon As Long
Dim Handle As Long
    Handle = ItemHwnd(Index)
    hIcon = SendMessage(Handle, WM_GETICON, ICON_BIG, 0)
    If hIcon <> 0 Then
        ItemIcon = CopyIcon(hIcon)
        Exit Property
    End If
    hIcon = SendMessage(Handle, WM_GETICON, ICON_SMALL, 0)
    If hIcon <> 0 Then
        ItemIcon = CopyIcon(hIcon)
        Exit Property
    End If
    hIcon = GetClassLong(Handle, GCL_HICON)
    If hIcon <> 0 Then
        ItemIcon = CopyIcon(hIcon)
        Exit Property
    End If
    hIcon = GetClassLong(Handle, GCL_HICONSM)
    If hIcon <> 0 Then
        ItemIcon = CopyIcon(hIcon)
        Exit Property
    End If
    ItemIcon = LoadIcon(0, IDI_WINLOGO)
End Property

Public Sub DrawIcon(IconHandle As Long, hDC As Long, X As Long, y As Long, xWidth As Long, yHeight As Long)
    DrawIconEx hDC, X, y, IconHandle, xWidth, yHeight, 0, 0, DI_NORMAL
    DestroyIcon IconHandle
End Sub

Private Sub UserControl_Resize()
    Width = 480
    Height = 480
End Sub
