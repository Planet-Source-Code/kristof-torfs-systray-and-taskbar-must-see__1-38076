Attribute VB_Name = "modTray"
Option Explicit
' All the code has been based on:
'
'           Softworld tm´s * Softshell Logi Beta 1.3 *
'           abd Mattias Sjögren's code (SysTray).
'           this version: (c)2000 Roeland Kluit
'
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function RegisterClassEx Lib "user32" Alias "RegisterClassExA" (pcWndClassEx As WNDCLASSEX) As Integer
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Integer) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

Private Type COPYDATASTRUCT
  dwData As Long
  cbData As Long
  lpData As Long
End Type

Private Type WNDCLASSEX
  cbSize As Long
  Style As Long
  lpfnWndProc As Long
  cbClsExtra As Long
  cbWndExtra As Long
  hInstance As Long
  hIcon As Long
  hCursor As Long
  hbrBackground As Long
  lpszMenuName As String
  lpszClassName As String
  hIconSm As Long
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const WM_GETICON = &H7F
Private Const WM_QUERYDRAGICON = &H37
Private Const WM_COPYDATA = &H4A
Private Const WS_POPUP = &H80000000
Private Const WS_EX_TOPMOST = &H8&
Private Const HWND_BROADCAST = &HFFFF&
Private Const DI_NORMAL = &H3
Private Const GCL_HICON = (-14)
Private Const GCL_HICONSM = (-34)
Private Const WC_SYSTRAY As String = "Shell_TrayWnd"

Private Pwidth As Long
Private stBool As Boolean
Private m_hTaskBarCreated As Long
Private m_hSysTray As Long
Private sObjLeft As Long
Private IconIndex As Integer
Private lastIndex As Integer

Public m_colTrayIcons As Collection

Public Sub LoadTrayIconHandler()
 Dim wcx As WNDCLASSEX
  Dim lRet As Long
  
  IconIndex = 1
  stBool = False
  m_hTaskBarCreated = RegisterWindowMessage("TaskbarCreated")
  
  With wcx
    .cbSize = Len(wcx)
    .lpfnWndProc = FuncPtr(AddressOf WindowProc)
    .hInstance = App.hInstance
    .lpszClassName = WC_SYSTRAY
  End With
  
  Call RegisterClassEx(wcx)
  
  m_hSysTray = CreateWindowEx(WS_EX_TOPMOST, WC_SYSTRAY, vbNullString, WS_POPUP, 0&, 0&, 0&, 0&, 0&, 0&, App.hInstance, ByVal 0&)

  Set m_colTrayIcons = New Collection
    
    For lRet = 1 To m_colTrayIcons.Count
      m_colTrayIcons.Remove 1
    Next
 
    Call SendMessage(HWND_BROADCAST, m_hTaskBarCreated, 0&, ByVal 0&)
  
End Sub

Public Sub UnLoadTrayIconHandler()

  ' destroy systray window ...
  Call DestroyWindow(m_hSysTray)
  
  ' ... and unregister the window class
  Call UnregisterClass(WC_SYSTRAY, App.hInstance)
  
  ' free icon collection
  Set m_colTrayIcons = Nothing

End Sub

Public Function GetIcon(hwnd As Long) As Long
    Call SendMessageTimeout(hwnd, WM_GETICON, 0, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICONSM)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_GETICON, 1, 0, 0, 1000, GetIcon)
    If Not CBool(GetIcon) Then GetIcon = GetClassLong(hwnd, GCL_HICON)
    If Not CBool(GetIcon) Then Call SendMessageTimeout(hwnd, WM_QUERYDRAGICON, 0, 0, 0, 1000, GetIcon)
End Function

Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  Static cds As COPYDATASTRUCT
  If uMsg = WM_COPYDATA Then
    MoveMemory cds, ByVal lParam, Len(cds)
    If cds.dwData = 1 Then  ' this is probably a tray message
      WindowProc = TrayIconHandler(cds.lpData)
      Exit Function
    End If
  End If
  
  WindowProc = DefWindowProc(hwnd, uMsg, wParam, lParam)
  
End Function

' AddressOf wrapper
Private Function FuncPtr(ByVal pfn As Long) As Long
  FuncPtr = pfn
End Function

Private Function TrayIconHandler(ByVal lpIconData As Long) As Long
  
  Dim nid As NOTIFYICONDATA
  Dim ti As clsTrayIcon
  Dim dwMessage As Long
  Dim sKey As String
  
  ' The NIM_ message starts 4 bytes after lpIconData
  MoveMemory dwMessage, ByVal lpIconData + 4, Len(dwMessage)
  ' The NOTIFYICONDATA struct starts 8 bytes after lpIconData
  MoveMemory nid, ByVal lpIconData + 8, Len(nid)

  sKey = KeyFromIcon(nid.hwnd, nid.uID)
  
  On Error Resume Next
  Dim Ol As Long
  Select Case dwMessage
    Case NIM_ADD
      
      Set ti = New clsTrayIcon
      ti.ModifyFromNID lpIconData + 8
      m_colTrayIcons.Add ti, sKey
      
      With ti
        '//--Softworld Code 2000-08-12
            If stBool = False Then sObjLeft = frmTray.imgTrayIcon(IconIndex - 1).Left + frmTray.imgTrayIcon(IconIndex - 1).Width + 40
            stBool = False
            Load frmTray.imgTrayIcon(IconIndex)
            
            frmTray.imgTrayIcon(IconIndex).Picture = .VBIcon
            frmTray.imgTrayIcon(IconIndex).Top = 40
            frmTray.imgTrayIcon(IconIndex).Left = sObjLeft
            frmTray.imgTrayIcon(IconIndex).Width = frmTray.imgTrayIcon(0).Width
            frmTray.imgTrayIcon(IconIndex).Height = frmTray.imgTrayIcon(0).Height
            frmTray.imgTrayIcon(IconIndex).Visible = True
            frmTray.imgTrayIcon(IconIndex).Tag = sKey
            frmTray.imgTrayIcon(IconIndex).ToolTipText = .ToolTipText
            IconIndex = IconIndex + 1
        '//--
      End With
      
    Case NIM_MODIFY
      
      Set ti = m_colTrayIcons(sKey)
      
      With ti
        .ModifyFromNID lpIconData + 8
        '//--Softworld Code
      
      For Ol = 1 To frmTray.imgTrayIcon.Count - 1
        If frmTray.imgTrayIcon(Ol).Tag = sKey Then
            frmTray.imgTrayIcon(Ol).Picture = .VBIcon
            Exit For
        End If
      Next Ol
    
      '//--
      End With
      
    Case NIM_DELETE
      
      m_colTrayIcons.Remove sKey
      '//--Softworld Code
      
      For Ol = 1 To frmTray.imgTrayIcon.Count - 1
        If frmTray.imgTrayIcon(Ol).Tag = sKey Then
            frmTray.imgTrayIcon(Ol).Tag = "skip"
           
            frmTray.imgTrayIcon(Ol).Visible = False
            Call FixTrayIcons
            
        Exit For
        End If
      Next Ol
    
      '//--
  End Select
  
  Set ti = Nothing

  TrayIconHandler = 1

End Function

Private Function KeyFromIcon(ByVal hOwner As Long, ByVal ID As Long) As String
  KeyFromIcon = "K" & Hex$(hOwner) & "-" & Trim$(Str$(ID))
End Function

Private Sub FixTrayIcons()
'//--Softworld Code
Dim Lo As Long
Dim Asa As Long
For Lo = 1 To frmTray.imgTrayIcon.Count - 1
    If frmTray.imgTrayIcon(Lo).Tag <> "skip" Then
       
        frmTray.imgTrayIcon(Lo).Left = 40 + Asa
        Asa = Asa + frmTray.imgTrayIcon(0).Width + 40
    End If
Next Lo
For Lo = frmTray.imgTrayIcon.Count - 1 To 1 Step -1
    If frmTray.imgTrayIcon(Lo).Tag <> "skip" Then
        sObjLeft = frmTray.imgTrayIcon(Lo).Left + frmTray.imgTrayIcon(Lo).Width + 40
    Exit For
    End If
Next Lo
stBool = True
End Sub

Private Sub DrawIcon(hDC As Long, hwnd As Long, X As Integer, y As Integer)
    Dim ico As Long
    ico = GetIcon(hwnd)
    DrawIconEx hDC, X, y, ico, 16, 16, 0, 0, DI_NORMAL
End Sub

Private Sub UpdateButtonIcon(Index As Long, Sort As Integer)
    DrawIcon frmTray.picIcon(Sort).hDC, Index, 1, 1
End Sub
Public Function freeTrayObjects()
    Dim ti As clsTrayIcon
    On Error Resume Next
    Set ti = m_colTrayIcons(frmTray.imgTrayIcon(lastIndex).Tag)
    If Err = 0 Then
        frmTray.SetFocus
        ti.PostCallbackMessage WM_MOUSEMOVE
        lastIndex = -1
    End If
End Function

Public Function trayClick(ByVal msg As TrayIconMouseMessages, ByVal Index As Integer) As Boolean
On Error Resume Next
    Dim ti As clsTrayIcon
    Dim lRet As Long
           
    Set ti = m_colTrayIcons(frmTray.imgTrayIcon(Index).Tag)
    
    If Err.Number = 0 Then
        ti.PostCallbackMessage msg
        trayClick = True
        lastIndex = Index
    Else
        Err.Clear
    End If
    Set ti = Nothing

End Function
