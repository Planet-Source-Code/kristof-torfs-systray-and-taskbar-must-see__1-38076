Attribute VB_Name = "modDevice1"
'*******************************************************************************
'   Thank´s to Microsoft for some of the code..........
'   Thank´s to Brian for some of the code..........
'   Thank's to LogiShell for this module..........
'*******************************************************************************
Public Declare Function Shutdown Lib "user32" Alias "ExitWindowsEx" _
    (ByVal uFlags As Long, _
    ByVal dwReserved As Long) As Long
    
Public Declare Function CopyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal hIcon As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Boolean
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Boolean
Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Public Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl&, ByVal i&, ByVal hDCDest&, ByVal X&, ByVal y&, ByVal flags&) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
Public Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRECT As RECT) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsIconic Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetForeGroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function fCreateShellLink Lib "vb6stkit.DLL" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String, ByVal fPrivate As Integer, ByVal sParent As String) As Long
Public Declare Function SHRunDialog Lib "shell32.dll" Alias "#61" (ByVal hwndOwner As Long, ByVal dwReserved1 As Long, ByVal dwReserved2 As Long, ByVal szTitle As String, ByVal szPrompt As String, ByVal uFlags As Long) As Long
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Public Const MAX_PATH = 260

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type POINTAPI
    X As Long
    y As Long
End Type

Public Const HWND_TOPMOST = -1

Public Const WM_GETICON = &H7F
Public Const WM_CLOSE = &H10

Public Const SW_MINIMIZE = 6
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Const SW_RESTORE = 9

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40

Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_EXETYPE = &H2000
Public Const SHGFI_SYSICONINDEX = &H4000  ' System icon index
Public Const SHGFI_LARGEICON = &H0        ' Large icon
Public Const SHGFI_SMALLICON = &H1        ' Small icon
Public Const ILD_TRANSPARENT = &H1        ' Display transparent
Public Const SHGFI_SHELLICONSIZE = &H4
Public Const SHGFI_TYPENAME = &H400
Public Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Public Const EWX_FORCE = 4
Public Const EWX_LOGOFF = 0
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Const GW_OWNER = 4
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_EXSTYLE = (-20)

Public Const WS_EX_TOOLWINDOW = &H80
Public Const WS_EX_APPWINDOW = &H40000

Public Const LB_ADDSTRING = &H180
Public Const LB_SETITEMDATA = &H19A

Public Const SHRD_NOBROWSE = 1 ' If specified, the "Browse" button won't appear
Public Const SHRD_NOSTRING = 2 ' If specified, there won't be an initial string in the dialog

Public Const VK_ACTION = &H46
Public Const VK_LWIN = &H5B
Public Const KEYEVENTF_KEYUP = &H2

Public shinfo                   As SHFILEINFO
Public OrigForm              As Integer
Public TaskbarOpen        As Boolean
Public SubShown(0 To 3) As Boolean

Dim hwnd As Long
Dim OwnHwnd As Long

'//--Make a shortcut anywere-------------------------------------------
Function MakeShortCut(NewName As String, OrginalPath As String)
    Dim longReturn As Long
    longReturn = fCreateShellLink("..\..\skrivbord", NewName, OrginalPath, "", -1, "$(Programs)")
End Function

'//--Thank´s to someone on the PSC for this function.....
'//--This function displays the Run Dialog (like in Start -> Run...).
'//-- hWndOwner: The hWnd of the owner of the dialog.
'//-- Caption: "Run" is a bad caption, choose your own!
'//-- Prompt: "Type the name of a program, wait, actually don't!" Finally get to choose what you want to write there.
'//-- BrowseButton: Whether or not you want that Browse... button there.
'//-- InitialSelection: Whether or not you want anything to be written in the ComboBox when started (if False,
'//--                   a string is retrieved from the Run MRU list in the registry).
'//-- Returns False on failure (though I could never get it to fail) or True on success.
Function DisplayRunDialog(Optional ByVal hwndOwner As Long = 0, _
    Optional ByVal Caption As String = vbNullString, _
    Optional ByVal Prompt As String = vbNullString, _
    Optional ByVal BrowseButton As Boolean = True, _
    Optional ByVal InitialSelection As Boolean = True) As Boolean
    
    Dim uFlags As Long
    If Not BrowseButton Then uFlags = uFlags Or SHRD_NOBROWSE
    If Not InitialSelection Then uFlags = uFlags Or SHRD_NOSTRING
    DisplayRunDialog = Not CBool(SHRunDialog(hwndOwner, 0, 0, Caption, Prompt, uFlags)) ' No! two Reservedies! "RUN"!

End Function

Public Function ExtractFileName(ByVal strPath As String) As String
    ' StrReverse is only working in VB6
    If strPath = "" Then Exit Function
    strPath = StrReverse(strPath)
    strPath = Left(strPath, InStr(strPath, "\") - 1)
    ExtractFileName = StrReverse(strPath)
End Function

Public Sub HideExplorerTaskBar()
Dim hwnd1 As Long
    hwnd1 = FindWindow("Shell_TrayWnd", vbNullString)
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
End Sub

Public Sub ShowExplorerTaskBar()
Dim hwnd1 As Long
    hwnd1 = FindWindow("Shell_TrayWnd", vbNullString)
    Call SetWindowPos(hwnd1, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
End Sub
            

Public Sub pSetForegroundWindow(ByVal hwnd As Long)
Dim lForeThreadID As Long
Dim lThisThreadID As Long
Dim lReturn       As Long
'
' Make a window, specified by its handle (hwnd)
' the foreground window.
'
' If it is already the foreground window, exit.
'
If hwnd <> GetForegroundWindow() Then
    If IsIconic(hwnd) Then
       Call ShowWindow(hwnd, SW_RESTORE)
    Else
       Call ShowWindow(hwnd, SW_SHOW)
    End If
    '
    ' Get the threads for this window and the foreground window.
    '
    lForeThreadID = GetWindowThreadProcessId(GetForegroundWindow, ByVal 0&)
    lThisThreadID = GetWindowThreadProcessId(hwnd, ByVal 0&)
    '
    ' By sharing input state, threads share their concept of
    ' the active window.
    '
    If lForeThreadID <> lThisThreadID Then
        ' Attach the foreground thread to this window.
        Call AttachThreadInput(lForeThreadID, lThisThreadID, True)
        ' Make this window the foreground window.
        lReturn = SetForeGroundWindow(hwnd)
        ' Detach the foreground window's thread from this window.
        Call AttachThreadInput(lForeThreadID, lThisThreadID, False)
    Else
       lReturn = SetForeGroundWindow(hwnd)
    End If
    '
    ' Restore this window to its normal size.
    '

End If
End Sub

Public Function fEnumWindows(lst As ListBox, OwnHandle As Long) As Long
'
' Clear list, then fill it with the running
' tasks. Return the number of tasks.
'
' The EnumWindows function enumerates all top-level windows
' on the screen by passing the handle of each window, in turn,
' to an application-defined callback function. EnumWindows
' continues until the last top-level window is enumerated or
' the callback function returns FALSE.
'
    OwnHwnd = OwnHandle
With lst
    .Clear
    Call EnumWindows(AddressOf fEnumWindowsCallBack, .hwnd)
    fEnumWindows = .ListCount
End With
End Function

Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lParam As Long) As Long
Dim lReturn     As Long
Dim lExStyle    As Long
Dim bNoOwner    As Boolean
Dim sWindowText As String
'
' This callback function is called by Windows (from
' the EnumWindows API call) for EVERY window that exists.
' It populates the listbox with a list of windows that we
' are interested in.
'
' Windows to display are those that:
'   -   are not this app's
'   -   are visible
'   -   do not have a parent
'   -   have no owner and are not Tool windows OR
'       have an owner and are App windows
'
If hwnd <> OwnHwnd Then
    If IsWindowVisible(hwnd) Then
        If GetParent(hwnd) = 0 Then
            bNoOwner = (GetWindow(hwnd, GW_OWNER) = 0)
            lExStyle = GetWindowLong(hwnd, GWL_EXSTYLE)
            
            If (((lExStyle And WS_EX_TOOLWINDOW) = 0) And bNoOwner) Or _
                ((lExStyle And WS_EX_APPWINDOW) And Not bNoOwner) Then
                '
                ' Get the window's caption.
                '
                sWindowText = Space$(256)
                lReturn = GetWindowText(hwnd, sWindowText, Len(sWindowText))
                If lReturn Then
                   '
                   ' Add it to our list.
                   '
                   sWindowText = Left$(sWindowText, lReturn) & vbCrLf & hwnd
                   lReturn = SendMessage(lParam, LB_ADDSTRING, 0, ByVal sWindowText)
                   Call SendMessage(lParam, LB_SETITEMDATA, lReturn, ByVal hwnd)
                End If
            End If
        End If
    End If
End If
fEnumWindowsCallBack = True
End Function
