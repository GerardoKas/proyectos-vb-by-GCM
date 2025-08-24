VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "rtfGCMClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Const WS_BORDER = &H800000
Private Const WS_CAPTION = &HC00000
Private Const WS_CHILD = &H40000000
Private Const WS_CHILDWINDOW = (WS_CHILD)
Private Const WS_CLIPCHILDREN = &H2000000
Private Const WS_CLIPSIBLINGS = &H4000000
Private Const WS_DISABLED = &H8000000
Private Const WS_DLGFRAME = &H400000
Private Const WS_GROUP = &H20000
Private Const WS_HSCROLL = &H100000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_VISIBLE = &H10000000
Private Const WS_VSCROLL = &H200000

'The following window styles can be specified in the dwStyle parameter:
'
'Style  Meaning
'WS_BORDER  Creates a window that has a thin-line border.
'WS_CAPTION Creates a window that has a title bar (includes the WS_BORDER style).
'WS_CHILD   Creates a child window. This style cannot be used with the WS_POPUP style.
'WS_CHILDWINDOW Same as the WS_CHILD style.
'WS_CLIPCHILDREN    Excludes the area occupied by child windows when drawing occurs within the parent window. This style is used when creating the parent window.
'WS_CLIPSIBLINGS    Clips child windows relative to each other; that is, when a particular child window receives a WM_PAINT message, the WS_CLIPSIBLINGS style clips all other overlapping child windows out of the region of the child window to be updated. If WS_CLIPSIBLINGS is not specified and child windows overlap, it is possible, when drawing within the client area of a child window, to draw within the client area of a neighboring child window.
'WS_DISABLED    Creates a window that is initially disabled. A disabled window cannot receive input from the user.
'WS_DLGFRAME    Creates a window that has a border of a style typically used with dialog boxes. A window with this style cannot have a title bar.
'WS_GROUP   Specifies the first control of a group of controls. The group consists of this first control and all  controls defined after it, up to the next control with the WS_GROUP style. The first control in each group usually has the WS_TABSTOP style so that the user can move from group to group. The user can subsequently change the keyboard focus from one control in the group to the next control in the group by using the direction keys.
'WS_HSCROLL Creates a window that has a horizontal scroll bar.
'WS_ICONIC  Creates a window that is initially minimized. Same as the WS_MINIMIZE style.
'WS_MAXIMIZE    Creates a window that is initially maximized.
'WS_MAXIMIZEBOX Creates a window that has a Maximize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
'WS_MINIMIZE    Creates a window that is initially minimized. Same as the WS_ICONIC style.
'WS_MINIMIZEBOX Creates a window that has a Minimize button. Cannot be combined with the WS_EX_CONTEXTHELP style. The WS_SYSMENU style must also be specified.
'WS_OVERLAPPED  Creates an overlapped window. An overlapped window has a title bar and a border. Same as the WS_TILED style.
'WS_OVERLAPPEDWINDOW    Creates an overlapped window with the WS_OVERLAPPED, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and WS_MAXIMIZEBOX styles. Same as the WS_TILEDWINDOW style.
'WS_POPUP   Creates a pop-up window. This style cannot be used with the WS_CHILD style.
'WS_POPUPWINDOW Creates a pop-up window with WS_BORDER, WS_POPUP, and WS_SYSMENU styles. The WS_CAPTION and WS_POPUPWINDOW styles must be combined to make the window menu visible.
'WS_SIZEBOX Creates a window that has a sizing border. Same as the WS_THICKFRAME style.
'WS_SYSMENU Creates a window that has a window-menu on its title bar. The WS_CAPTION style must also be specified.
'WS_TABSTOP Specifies a control that can receive the keyboard focus when the user presses the TAB key. Pressing the TAB key changes the keyboard focus to the next control with the WS_TABSTOP style.
'WS_THICKFRAME  Creates a window that has a sizing border. Same as the WS_SIZEBOX style.
'WS_TILED   Creates an overlapped window. An overlapped window has a title bar and a border. Same as the WS_OVERLAPPED style.
'WS_TILEDWINDOW Creates an overlapped window with the WS_OVERLAPPED, WS_CAPTION, WS_SYSMENU, WS_THICKFRAME, WS_MINIMIZEBOX, and WS_MAXIMIZEBOX styles. Same as the WS_OVERLAPPEDWINDOW style.
'WS_VISIBLE Creates a window that is initially visible.
'WS_VSCROLL Creates a window that has a vertical scroll bar.
'(...)
''
Private Const WS_EX_ACCEPTFILES = &H10&
Private Const WS_EX_DLGMODALFRAME = &H1&
Private Const WS_EX_NOPARENTNOTIFY = &H4&
Private Const WS_EX_TOPMOST = &H8&
Private Const WS_EX_TRANSPARENT = &H20&
'dwExStyle
'
'Specifies the extended style of the window. This parameter can be one of the following values:
'
'Style Meaning
'WS_EX_ACCEPTFILES   Specifies that a window created with this style accepts drag-drop files.
'WS_EX_APPWINDOW Forces a top-level window onto the taskbar when the window is minimized.
'WS_EX_CLIENTEDGE    Specifies that a window has a border with a sunken edge.
'WS_EX_CONTEXTHELP   Includes a question mark
'in the title bar of the window.
'When the user clicks the question mark, the cursor changes to a question mark with a pointer. If the user then clicks a child window, the child receives a WM_HELP message.
'The child window should pass the message to the parent window procedure, which should call the WinHelp function using the HELP_WM_HELP command. The Help application displays a pop-up window that typically contains help for the child window.WS_EX_CONTEXTHELP cannot be used with the WS_MAXIMIZEBOX or WS_MINIMIZEBOX styles.
'WS_EX_CONTROLPARENT Allows the user to navigate among the child windows of the window by using the TAB key.
'WS_EX_DLGMODALFRAME Creates a window that has a double border; the window can, optionally, be created with a title bar by specifying the WS_CAPTION style in the dwStyle parameter.
'WS_EX_LEFT  Window has generic "left-aligned" properties. This is the default.
'WS_EX_LEFTSCROLLBAR If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the vertical scroll bar (if present) is to the left of the client area. For other languages, the style is ignored and not treated as an error.
'WS_EX_LTRREADING    The window text is displayed using Left to Right reading-order properties. This is the default.
'WS_EX_MDICHILD  Creates an MDI child window.
'WS_EX_NOPARENTNOTIFY    Specifies that a child window created with this style does not send the WM_PARENTNOTIFY message to its parent window when it is created or destroyed.
'WS_EX_OVERLAPPEDWINDOW  Combines the WS_EX_CLIENTEDGE and WS_EX_WINDOWEDGE styles.
'WS_EX_PALETTEWINDOW Combines the WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOW, and WS_EX_TOPMOST styles.
'WS_EX_RIGHT Window has generic "right-aligned" properties. This depends on the window class. This style has an effect only if the shell language is Hebrew, Arabic, or another language that supports reading order alignment; otherwise, the style is ignored and not treated as an error.
'WS_EX_RIGHTSCROLLBAR    Vertical scroll bar (if present) is to the right of the client area. This is the default.
'WS_EX_RTLREADING    If the shell language is Hebrew, Arabic, or another language that supports reading order alignment, the window text is displayed using Right to Left reading-order properties. For other languages, the style is ignored and not treated as an error.
'WS_EX_STATICEDGE    Creates a window with a three-dimensional border style intended to be used for items that do not accept user input.
'WS_EX_TOOLWINDOW    Creates a tool window; that is, a window intended to be used as a floating toolbar. A tool window has a title bar that is shorter than a normal title bar, and the window title is drawn using a smaller font. A tool window does not appear in the taskbar or in the dialog that appears when the user presses ALT+TAB.
'WS_EX_TOPMOST   Specifies that a window created with this style should be placed above all non-topmost windows and should stay above them, even when the window is deactivated. To add or remove this style, use the SetWindowPos function.
'WS_EX_TRANSPARENT   Specifies that a window created with this style is to be transparent. That is, any windows that are beneath the window are not obscured by the window. A window created with this style receives WM_PAINT messages only after all sibling windows beneath it have been updated.
'WS_EX_WINDOWEDGE    Specifies that a window has a border with a raised edge.
Private Const ES_AUTOHSCROLL = &H80&
Private Const ES_AUTOVSCROLL = &H40&
Private Const ES_CENTER = &H1&
Private Const ES_LEFT = &H0&
Private Const ES_LOWERCASE = &H10&
Private Const ES_MULTILINE = &H4&
Private Const ES_NOHIDESEL = &H100&
Private Const ES_OEMCONVERT = &H400&
Private Const ES_PASSWORD = &H20&
Private Const ES_READONLY = &H800&
Private Const ES_RIGHT = &H2&
Private Const ES_UPPERCASE = &H8&
Private Const ES_WANTRETURN = &H1000&
Private Const ES_SUNKEN As Long = &H4000




' /* all character format measurements are in twips */
' Corrected...
Private Const WM_USER As Long = &H400
Private Const EM_SETCHARFORMAT As Long = (WM_USER + 68)
Private Const EM_GETCHARFORMAT As Long = (WM_USER + 58)
Private Const EM_SETBKGNDCOLOR As Long = (WM_USER + 67)
Private Const LF_FACESIZE As Long = 32

Private Type CHARFORMAT
    cbSize As Integer '2
    wPad1 As Integer  '4
    dwMask As Long    '8
    dwEffects As Long '12
    yHeight As Long   '16
    yOffset As Long   '20
    crTextColor As Long '24
    bCharSet As Byte    '25
    bPitchAndFamily As Byte '26
    szFaceName(0 To LF_FACESIZE - 1) As Byte ' 58
    wPad2 As Integer ' 60
End Type
Private Const SCF_ALL As Long = &H4
Private Const SCF_SELECTION As Long = &H1
Private Const SCF_WORD As Long = &H2
'dwmask apply
Private Const CFM_COLOR As Long = &H40000000
'Private Const CFM_ALL As Long = (CFM_EFFECTS Or CFM_SIZE Or CFM_FACE Or CFM_OFFSET Or CFM_CHARSET)
'Private Const CFM_EFFECTS As Long = (CFM_BOLD Or CFM_ITALIC Or CFM_UNDERLINE Or CFM_COLOR Or CFM_STRIKEOUT Or CFE_PROTECTED Or CFM_LINK)
Private Const CFM_SIZE As Long = &H80000000
Private Const CFM_FACE As Long = &H20000000
Private Const CFM_OFFSET As Long = &H10000000
Private Const CFM_CHARSET As Long = &H8000000
Private Const CFM_BACKCOLOR As Long = &H4000000
Private Const CFM_BOLD As Long = &H1
Private Const CFM_ITALIC As Long = &H2

Private Const EM_GETSELTEXT As Long = (WM_USER + 62)
Private Const EM_EXLIMITTEXT As Long = (WM_USER + 53)
'Private Const EM_EXGETSEL As Long = (WM_USER + 52)
Private Const EM_EXSETSEL As Long = (WM_USER + 55)
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type


Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
     ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByRef lParam As Any) As Long


Const ws_Mine = WS_CHILD Or WS_HSCROLL Or WS_VSCROLL Or WS_VISIBLE Or ES_MULTILINE Or ES_SUNKEN

Const ws_ex_Mine = WS_EX_ACCEPTFILES

Dim myHwnd As Long
Dim mWidth As Long
Dim mHeight As Long
Dim mLeft As Long
Dim mTop As Long
Dim mLimitSize As Long

Private Sub Class_Initialize()
myHwnd = CreateWindowEx(ws_ex_Mine, "richedit", "GCMRich", ws_Mine, 0, 0, 50, 50, Form1.hwnd, 0, 0, 0)
setFont
LimitSize = 64000
End Sub

Public Sub move(x, y, w, h)
MoveWindow myHwnd, x, y, w, h, 1
End Sub

Private Sub GetRect()
Dim mRect As RECT
r = GetWindowRect(myHwnd, mRect)
mWidth = mRect.Right - mRect.Left
mHeight = mRect.Bottom - mRect.Top
mLeft = mRect.Left
mTop = mRect.Top
End Sub

Public Sub setFont()
Dim mFmt As CHARFORMAT

'r = SendMessage(myHwnd, EM_GETCHARFORMAT, sfc_all, mFmt)
mFmt.cbSize = Len(mFmt)
mFmt.dwMask = CFM_FACE Or CFM_SIZE Or CFM_BOLD Or CFM_COLOR
mFmt.dwEffects = 0
mFmt.crTextColor = RGB(0, 0, 255)
mFmt.yHeight = 6 * Screen.TwipsPerPixelY
'mFmt.yOffset = -5 * Screen.TwipsPerPixelY
r = SendMessage(myHwnd, EM_SETCHARFORMAT, sfc_all, mFmt)

'EM_SETCHARFORMAT
'wParam = (WPARAM) (UINT) uFlags;
'lParam = (LPARAM) (CHARFORMAT FAR *) lpFmt;

'Parameters

'uFlags
'Character formatting that applies to the control. If this parameter is zero, the default character format is set. Otherwise, it can be one of the following values:
'
'Value Meaning
'SCF_ALL Applies the formatting to all text in the control.
'SCF_SELECTION   Applies the formatting to the current selection. If the selection is empty, the character formatting is applied to the insertion point, and the new character format is in effect only until the insertion point changes.
'SCF_WORD | SCF_SELECTION    Applies the formatting to the selected word or words. If the selection is empty but the insertion point is inside a word, the formatting is applied to the word. The SCF_WORD value must be used in conjunction with the SCF_SELECTION value.
'
'lpFmt
'Pointer to a CHARFORMAT structure specifying the character formatting to use. Only the formatting attributes specified by the dwMask member are changed.

 
End Sub


Public Property Get Text()
Dim mBuffer As String
Dim mCharRange As CHARRANGE
mCharRange.cpMin = 1
mCharRange.cpMax = 5
r = SendMessage(myHwnd, EM_EXSETSEL, 0, mCharRange)
mBuffer = Space(r)
r = SendMessage(myHwnd, EM_GETSELTEXT, 0, mBuffer)
MsgBox mBuffer
End Property
Public Property Set Text(texto)

End Property

Public Property Let LimitSize(size)
mLimitSize = size
r = SendMessage(myHwnd, EM_EXLIMITTEXT, 0, CLng(size))
End Property
