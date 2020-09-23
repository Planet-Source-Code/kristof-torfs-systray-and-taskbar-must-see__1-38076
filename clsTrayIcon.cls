VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'//--Thank´s to Mattias Sjögren who is the orginal writer of this SysTray      --//
'//--I have made some changes to it so it will fit to my app                   --//

Public Enum TrayIconMouseMessages
  WM_MOUSEMOVE = &H200
  WM_LBUTTONDOWN = &H201
  WM_LBUTTONUP = &H202
  WM_LBUTTONDBLCLK = &H203
  WM_RBUTTONDOWN = &H204
  WM_RBUTTONUP = &H205
  WM_RBUTTONDBLCLK = &H206
  WM_MBUTTONDOWN = &H207
  WM_MBUTTONUP = &H208
  WM_MBUTTONDBLCLK = &H209
End Enum

' NOTIFYICONDATA flags
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Public Enum InfoTipFlags
  NIIF_NONE = &H0
  NIIF_INFO = &H1
  NIIF_WARNING = &H2
  NIIF_ERROR = &H3
End Enum

' OSVERSIONINFO platform flag
Private Const VER_PLATFORM_WIN32_NT = 2

'''''''''''''''''
'''   Types   '''
'''''''''''''''''

Private Type NOTIFYICONDATA
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64
End Type

' extended NOTIFYICONDATA - Implemented in shell32.dll >= v5.0 (Win2000)
Private Type NOTIFYICONDATA_5
  cbSize As Long
  hwnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128
  dwState As Long
  dwStateMask As Long
  szInfo As String * 256
  uTimeout As Long
  szInfoTitle As String * 64
  dwInfoFlags As Long
End Type

Private Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformID As Long
  szCSDVersion(127) As Byte
End Type

Private Type PICTDESC_ICON  ' PICTDESC for PICTYPE_ICON
  cbSizeofStruct As Long
  picType As Long
  hIcon As Long
  padding1 As Long
  padding2 As Long
End Type

Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type


''''''''''''''''''''
'''   Declares   '''
''''''''''''''''''''

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Private Declare Function OleCreateIconIndirect Lib "olepro32.dll" Alias "OleCreatePictureIndirect" (pPictDesc As PICTDESC_ICON, riid As GUID, ByVal fOwn As Long, ppvObj As IPicture) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


'''''''''''''''''''''
'''   Variables   '''
'''''''''''''''''''''

Private m_hOwner As Long
Private m_lID As Long
Private m_sToolTip As String
Private m_hIcon As Long
Private m_lMsg As Long
Private m_dtCreated As Date
Private m_dtModified As Date
Private m_picIcon As Picture
Private m_fSharedIcon As Boolean
Private m_fHidden As Boolean
Private m_sInfoTip As String
Private m_sInfoTitle As String
Private m_lInfoTimeout As Long
Private m_itfInfoIcon As InfoTipFlags

Private m_fIsUnicodeSystem As Boolean
Private m_fIsWindows2000 As Boolean

Public Sub ModifyFromNID(ByVal pNID As Long)

  Dim nid As NOTIFYICONDATA
  Dim nid5 As NOTIFYICONDATA_5
  Dim fIsNid5Struct As Boolean
  
  ' On Unicode systems (WinNT4 and Win2000), pNID will point to a
  ' NOTIFYICONDATAW struct, even if the calling app uses
  ' Shell_NotifyIconA. On Win9x, it's a pointer to a
  ' NOTIFYICONDATAA struct.
  If m_fIsUnicodeSystem Then
    MoveMemory ByVal VarPtr(nid), ByVal pNID, LenB(nid)
    If m_fIsWindows2000 And (nid.cbSize >= LenB(nid5)) Then
      MoveMemory ByVal VarPtr(nid5), ByVal pNID, LenB(nid5)
      fIsNid5Struct = True
    End If
  Else
    MoveMemory nid, ByVal pNID, Len(nid)
  End If
  
  If m_hOwner = 0 Then
    m_hOwner = nid.hwnd
    m_lID = nid.uID
  End If

  ' Update the modified properties
  If nid.uFlags And NIF_MESSAGE Then m_lMsg = nid.uCallbackMessage
  If nid.uFlags And NIF_ICON Then
    m_hIcon = nid.hIcon
    Set m_picIcon = PictureFromhIcon(m_hIcon)
  End If
  
  If fIsNid5Struct Then
    If nid5.uFlags And NIF_TIP Then m_sToolTip = nid5.szTip
    If nid5.uFlags And NIF_STATE Then
      If nid5.dwStateMask And NIS_HIDDEN Then m_fHidden = CBool(nid5.dwState And NIS_HIDDEN)
      If nid5.dwStateMask And NIS_SHAREDICON Then m_fSharedIcon = CBool(nid5.dwState And NIS_SHAREDICON)
    End If
    If nid5.uFlags And NIF_INFO Then
      m_itfInfoIcon = nid5.dwInfoFlags
      m_sInfoTip = nid5.szInfo
      m_sInfoTitle = nid5.szInfoTitle
      m_lInfoTimeout = nid5.uTimeout
      
    End If
  Else
    If nid.uFlags And NIF_TIP Then m_sToolTip = nid.szTip
  End If
  
  ' Update modified date
  m_dtModified = Now
  
End Sub

Public Property Get ToolTipText() As String
  ToolTipText = m_sToolTip
End Property

Public Property Get hIcon() As Long
  hIcon = m_hIcon
End Property

Public Property Get VBIcon() As IPictureDisp
  Set VBIcon = m_picIcon
End Property

Public Property Get CallbackMessage() As Long
  CallbackMessage = m_lMsg
End Property

Public Property Get OwnerWindow() As Long
  OwnerWindow = m_hOwner
End Property

Public Property Get ID() As Long
  ID = m_lID
End Property

Public Property Get Hidden() As Boolean
  Hidden = m_fHidden
End Property

Public Property Get SharedIcon() As Boolean
  SharedIcon = m_fSharedIcon
End Property

Public Property Get InfoTip() As String
  InfoTip = m_sInfoTip
End Property

Public Property Get InfoTitle() As String
  InfoTitle = m_sInfoTitle
End Property

Public Property Get InfoTipIcon() As InfoTipFlags
  InfoTipIcon = m_itfInfoIcon
End Property

Public Property Get InfoTimeout() As String
  InfoTimeout = m_lInfoTimeout
End Property

Public Property Get CreatedDate() As Date
  CreatedDate = m_dtCreated
End Property

Public Property Get ModifiedDate() As Date
  ModifiedDate = m_dtModified
End Property

Public Sub PostCallbackMessage(ByVal Message As TrayIconMouseMessages)

  ' Post message to the message queue of the owner window
  ' wParam = Icon ID
  ' lParam = Mouse message (WM_xBUTTONyyyyy)
  Call PostMessage(m_hOwner, m_lMsg, m_lID, Message)
  
End Sub

' Creates a VB friedly Picture object from a GDI icon object handle
Private Function PictureFromhIcon(ByVal hIcon As Long) As IPicture

  Dim oIcon As Picture
  Dim pdi As PICTDESC_ICON
  Dim IID_IPicture As GUID
    
  
  If hIcon = 0 Then Exit Function

  With pdi
    .cbSizeofStruct = Len(pdi)
    .picType = vbPicTypeIcon    ' PICTYPE_ICON
    .hIcon = hIcon
  End With
  
  ' IID_IPicture = {7BF80980-BF32-101A-8BBB-00AA00300CAB}
  With IID_IPicture
    .Data1 = &H7BF80980
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(2) = &H0
    .Data4(3) = &HAA
    .Data4(4) = &H0
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
  End With
  
  Call OleCreateIconIndirect(pdi, IID_IPicture, 0&, oIcon)

  Set PictureFromhIcon = oIcon

End Function

Private Sub Class_Initialize()

  Dim ovi As OSVERSIONINFO
  
  
  ' check if we are running NT
  ovi.dwOSVersionInfoSize = Len(ovi)
  Call GetVersionEx(ovi)
  m_fIsUnicodeSystem = CBool(ovi.dwPlatformID And VER_PLATFORM_WIN32_NT)
  
  If m_fIsUnicodeSystem And (ovi.dwMajorVersion >= 5) Then m_fIsWindows2000 = True
  
  m_dtCreated = Now
  m_dtModified = Now
  
  m_sInfoTip = "N/A"
  m_sInfoTitle = "N/A"
  
End Sub
