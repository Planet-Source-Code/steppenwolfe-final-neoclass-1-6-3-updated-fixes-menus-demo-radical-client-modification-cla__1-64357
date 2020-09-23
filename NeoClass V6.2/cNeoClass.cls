VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cNeoClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************************
'*  NeoClass! Radical Client Modification SuperClass v. 1.6.1                          *
'*                                                                                     *
'*  Created:     Febuary 13, 2006                                                      *
'*  Updated:     March 22 2006                                                         *
'*  Purpose:     Form Skinning Class                                                   *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.6                                                                   *
'*  Compile:     Native                                                                *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*  Credit:      Based in part on Neocaption written by Steve McMahon                  *
'*                                                                                     *
'***************************************************************************************

'*~ List of exposed functions ~*

'/~ Attatch                 - Attatch the skin

'*~ List of exposed properties ~*

'/~ Image Properties /~
'/~ p_IBorders              - border image
'/~ p_ICaption              - caption bar image
'/~ p_ICBoxMin              - control box minimize
'/~ p_ICBoxMax              - control box maximize
'/~ p_ICBoxRst              - control box restore
'/~ p_ICBoxCls              - control box close
'/~ p_ICommand              - command button image

'/~ Form Properties /~
'/~ p_ActiveRight           - caption right end
'/~ p_BorderHasInactive     - border has inactive image
'/~ p_BottomBorderHeight    - bottom sizing border
'/~ p_ButtonHeight          - control button height
'/~ p_ButtonWidth           - control button width
'/~ p_ButtonOffsetX         - control button offset horz
'/~ p_ButtonOffsetY         - control button offset vert
'/~ p_ControlHasInactive    - caption buttons have inactive
'/~ p_ControlButtonPosition - use custom button offsets
'/~ p_LeftEnd               - left caption bar end
'/~ p_LeftBorderWidth       - left border width
'/~ p_Offset                - title offset
'/~ p_RightEnd              - caption bar right end
'/~ p_RightBorderWidth      - right border width
'/~ p_SkinForm              - skin form switch
'/~ p_TopBorderHeight       - top sizing border height
'/~ p_WindowActive          - window status

'/~ Control Properties /~
'/~ p_SkinCommand           - skin command buttons
'/~ p_CmdFntClr             - command button font color

'/~ Extended Properties (Tags) /~

'/~ IL/IR                   - icon position set w/ control tag
'/~ NO                      - control is ignored


'/~ In my opinion, Steve McMahon is amoung the best vb programmers in the world..
'/~ That being said, his work is also some of the most difficult to understand,
'/~ (and lets face it, he is not big on leaving notes!).
'/~ Neocaption was a massive effort, (8 classes, and a module!) of complex, interdependant
'/~ and often perplexing sprawl.. What I did here, was remove 70% of
'/~ the code, rewrite what was left, created a dynamic instancing engine, and put the
'/~ subclassing inline..

'/~ Update: Febuary 17, 2006 V2
'/~ Made many small revisions and changes to optimize speed and performance
'/~ Removed all client area draw references to eliminate flicker on resize events
'/~ Removed more redundancies and migrated repetitive calls to variables set at instatntiation

'/~ Update: Febuary 18, 2006 V3
'/~ Migrated subclass engine into project. I could have dropped it into this class, maybe even
'/~ gotten rid of the implements interface and kept this project down to one class, but I opted to
'/~ keep the two seperate, in the interest of flexibility. New instances of the subclasser can
'/~ in that way be created for other chores, in est. control subclassing..
'/~ and anyways, neocaption has gone from 2 dlls with 13 modules down to 3 classes, (2 really),
'/~ so I'm ok with that..
'/~ Engine was borrowed from SSubTmr6.dll ASM Thunk version, with a few minor changes.
'/~ Transparency routines have been removed in the interest of backwards compatability,
'/~ another solution is in the works..

'/~ Update: Febuary 21, 2006 V4
'/~ Now employs Dynamic Subclass Instantiation Engine.
'/~ Subclass instances can now be generated or removed dynamically through the MGSubclass() class array.
'/~ A collection maps a control object hwnd to an index number. Index number is pointer into 7 dimensional
'/~ variant array that contains object heirarchy, namely:
'/~ ClassName: ex. "CommandButton", ObjectName: ex. "Command1", Control Array Index: ex. Command1(3)
'/~ Control Hwnd, and Subclass Class Instance: ex. m_GASubclass(2), Controls bitmap DC, and bitmap
'/~ Width and Height paramaters.
'/~ A message store collection object returns message types associated with the object class
'/~ for attatch/detatch events. Client area messages are now handled in a non-linear fashion using parallel
'/~ message queues, with callbacks from subclass interface returned to a global message distribution hub.
'/~ Also, added rollover image change to control box buttons, (original did not have this), and control box
'/~ buttons have been seperated from caption bar image, and shifted vertically. This removes button size limitation
'/~ of original, (caption height / 2). That, and I cleaned up the code some more..
'/~ If you have problems on your os, please be good enough to send me an email, and try to be specific
'/~ ie routine name, call, and where it jumps out/error number. If it does work on your os, chime in on list..
'/~ If someone wants to help with the drawing subs, I will extend it some more. Really, this could be evolved
'/~ into a unified control subclassing interface, anything with an hwnd could be added, drawing routines made
'/~ reusable, and class size kept managable.. Just depends on your imagination now..
'/~ I'll add another update with a better example, (controls on form possibly), then that's all there is..

'/~ Making a skin (Taken from RCM SC)
'/~ Creating a skin for this is really very easy once you get the hang of it. It only takes
'/~ me about an hour to do one, (now that I've had so much practice!) Here is the jist of it.
'/~ I am using Fireworks to create the graphics, but photoshop or any advanced editor will do.
'/~ Start with the caption bar, just a rectangular image, (but could be odd shaped w/ transparencies..).
'/~ I am using caption bars of 250 px long with varying heights, but most sizes (within reason) will
'/~ work fine. Remember though that images are tiled, so a very short bar would be a performance hit.
'/~ The caption bar size properties are set with the following properties:
'/~ p_ICaption - picture reference
'/~ p_CaptionHeight - caption bar image height
'/~ p_LeftEnd - past any text on the beginning of the bar
'/~ p_ActiveRight - this is where the right side of bar begins
'/~ p_RightEnd - the length of the bar
'/~ Button images are tristate - normal|hover|down and must be of the same size and tiled horizontally.
'/~ Button properties are:
'/~ p_ICBoxMin - picture reference to minimize buttons
'/~ p_ICBoxMax - picture reference to mazimize buttons
'/~ p_ICBoxRst - picture reference to restore buttons
'/~ p_ICBoxCls - picture reference to close buttons
'/~ p_ButtonHeight - button height in px
'/~ p_ButtonWidth - button width in px
'/~ p_ControlButtonPosition - use a custom offset (boolean)
'/~ p_ButtonOffsetX - button offset from left
'/~ p_ButtonOffsetY - button offset from right
'/~ Next create the frame image. There are seven rectangular slices to the frame. Each
'/~ must be equal in width and height and must be a perfect square. Ex. 10px by 10px would
'/~ need a frame image 10px high by 70px wide. The grouping is as follows:
'/~ 1) left side 2) left bottom corner 3) left bottom corner inactive 4) bottom
'/~ 5) right side 6) right bottom corner 7) right bottom corner inactive

'/~ Update: March 03, 2006 V5
'/~ Added an example of using the Dynamic Instancing Engine to subclass common controls in the client area.
'/~ Command buttons are now 'skinned', with a 5 state image set. Entire infrastructure is now in place
'/~ to subclass common controls.
'/~ Many small fixes and optimizations were also made, including escape from drag on caption button down,
'/~ and caption button reset when mouse is off the NC.

'/~ Update: March 21, 2006 V6
'/~ Fixed a bug where tracking array was accessed per its initialization which caused application hang
'/~ when compiled. Added owner drawn text and icon, and migrated a couple of the skins from RCM into this
'/~ project. Unless another major bug is found, this will be the last update.

'/~ Update: March 22, 2006 V6.1
'/~ Fixed caption text overlapping buttons with minimum form size constraint. Backwash into form
'/~ appears to have been resolved with an additional paint call. Fine tuned a couple other things..
'/~ Applied fixes to RCM also, including close bug, and added notes for what could be done to
'/~ complete the project.

'/~ Unless another major bug is found, that is -absolutely the last update! (I mean it this time! ;o)
'/~ Killed RCM SC
'/~ The main reason was the poor response.
'/~ I have included the last rev of RCM in this upload, (with notes), if you want to take a crack at
'/~ perfecting and reposting it, do so with my blessing..

'/~ Well, usual spiel, for a comment or a job.. steppenwolfe_2000@yahoo.com
'/~ Cheers
'/~ John


Implements MISubclass

Private Const GWL_STYLE                 As Long = (-16)
Private Const GWL_EXSTYLE               As Long = (-20)
Private Const DI_NORMAL                 As Long = &H1 Or &H2

Private Type LOGBRUSH
    lbStyle                             As Long
    lbColor                             As Long
    lbHatch                             As Long
End Type

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type WINDOWPOS
    hwnd                                As Long
    hWndInsertAfter                     As Long
    X                                   As Long
    Y                                   As Long
    cx                                  As Long
    cy                                  As Long
    flags                               As Long
End Type

Private Type NCCALCSIZE_PARAMS
    rgrc(0 To 2)                        As RECT
    lppos                               As Long
End Type

Private Type BITMAP
    bmType                              As Long
    bmWidth                             As Long
    bmHeight                            As Long
    bmWidthBytes                        As Long
    bmPlanes                            As Integer
    bmBitsPixel                         As Integer
    bmBits                              As Long
End Type

Private Type LOGFONT
    lfHeight                            As Long
    lfWidth                             As Long
    lfEscapement                        As Long
    lfOrientation                       As Long
    lfWeight                            As Long
    lfItalic                            As Byte
    lfUnderline                         As Byte
    lfStrikeOut                         As Byte
    lfCharSet                           As Byte
    lfOutPrecision                      As Byte
    lfClipPrecision                     As Byte
    lfQuality                           As Byte
    lfPitchAndFamily                    As Byte
    lfFaceName                          As String * 32
End Type

Private Type TRACKMOUSEEVENT_STRUCT
    cbSize                              As Long
    dwFlags                             As TRACKMOUSEEVENT_FLAGS
    hwndTrack                           As Long
    dwHoverTime                         As Long
End Type

Private Type tDrawCore
    '/* image params
    ImgRect                             As RECT
    TxtRect                             As RECT
    hdc                                 As Long
    hwnd                                As Long
    Focused                             As Boolean
    State                               As Long
    Offset                              As Long
    '/* icon
    ImageList                           As Long
    IconIdx                             As Long
    IconPos                             As eIcnPos
    '/* font
    Caption                             As String
    Font                                As String
    FontSize                            As Integer
    Bold                                As Boolean
    Italic                              As Boolean
    Underline                           As Boolean
    StrikeThrough                       As Boolean
    FontColor                           As Long
End Type

Private Type BLENDFUNCTION
    BlendOp                             As Byte
    BlendFlags                          As Byte
    SourceConstantAlpha                 As Byte
    AlphaFormat                         As Byte
End Type

Private Type MINMAXINFO
    ptReserved                          As POINTAPI
    ptMaxSize                           As POINTAPI
    ptMaxPosition                       As POINTAPI
    ptMinTrackSize                      As POINTAPI
    ptMaxTrackSize                      As POINTAPI
End Type

Private Enum eObjTrackType
    ObjDc = 0
    OldObj = 1
    NewObj = 2
End Enum

Private Enum TRACKMOUSEEVENT_FLAGS
  TME_HOVER = &H1&
  TME_LEAVE = &H2&
  TME_QUERY = &H40000000
  TME_CANCEL = &H80000000
End Enum

Private Enum ECNCButtonStates
    Up = &H0
    Over = &H1
    Down = &H2
End Enum

Public Enum ECNCSysCommandConstants
    SC_ARRANGE = &HF110&
    SC_CLOSE = &HF060&
    SC_MAXIMIZE = &HF030&
    SC_MINIMIZE = &HF020&
    SC_MOVE = &HF010&
    SC_NEXTWINDOW = &HF040&
    SC_PREVWINDOW = &HF050&
    SC_RESTORE = &HF120&
    SC_SIZE = &HF000&
End Enum

Public Enum eHitConstants
    HTBORDER = 18
    HTBOTTOM = 15
    HTBOTTOMLEFT = 16
    HTBOTTOMRIGHT = 17
    HTCAPTION = 2
    HTCLIENT = 1
    HTGROWBOX = 4
    HTHSCROLL = 6
    HTLEFT = 10
    HTMAXBUTTON = 9
    HTMENU = 5
    HTMINBUTTON = 8
    HTNOWHERE = 0
    HTRIGHT = 11
    HTSYSMENU = 3
    HTTOP = 12
    HTTOPLEFT = 13
    HTTOPRIGHT = 14
    HTVSCROLL = 7
End Enum

Public Enum eDrawMethods
    ECNCModifyNonClientArea = 0
    ECNCUseClientArea = 1
End Enum

Public Enum eVisible
    None = 0
    Mask = 1
    Region = 2
End Enum

Public Enum eIcnPos
    eLeft
    eRight
End Enum

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, _
                                                           ByVal X As Long, _
                                                           ByVal Y As Long, _
                                                           ByVal nWidth As Long, _
                                                           ByVal nHeight As Long, _
                                                           ByVal hSrcDC As Long, _
                                                           ByVal xSrc As Long, _
                                                           ByVal ySrc As Long, _
                                                           ByVal nSrcWidth As Long, _
                                                           ByVal nSrcHeight As Long, _
                                                           ByVal crTransparent As Long) As Boolean

Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, _
                                                       ByVal hwndTo As Long, _
                                                       lppt As Any, _
                                                       ByVal cPoints As Long) As Long

Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, _
                                                                          ByVal hWnd2 As Long, _
                                                                          ByVal lpsz1 As String, _
                                                                          lpsz2 As Any) As Long

Private Declare Function IsZoomed Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, _
                                               ByVal hMenu As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
                                              ByVal dX As Long, _
                                              ByVal dy As Long, _
                                              ByVal cButtons As Long, _
                                              ByVal dwExtraInfo As Long)

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Any) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal X As Long, _
                                                ByVal Y As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, _
                                                    lprcUpdate As Any, _
                                                    ByVal hrgnUpdate As Long, _
                                                    ByVal fuRedraw As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, _
                                                     ByVal X As Long, _
                                                     ByVal Y As Long, _
                                                     ByVal nWidth As Long, _
                                                     ByVal nHeight As Long, _
                                                     ByVal hSrcDC As Long, _
                                                     ByVal xSrc As Long, _
                                                     ByVal ySrc As Long, _
                                                     ByVal nSrcWidth As Long, _
                                                     ByVal nSrcHeight As Long, _
                                                     ByVal dwRop As Long) As Long

Private Declare Function GetDC Lib "user32.dll" (ByVal hwnd As Long) As Long

Private Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function TrackMouseEventComCtl Lib "Comctl32" Alias "_TrackMouseEvent" (lpEventTrack As TRACKMOUSEEVENT_STRUCT) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal lpLibFileName As String) As Long

Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long, _
                                                   ByRef lpPoint As Any) As Long

Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
                                                               ByVal lHPalette As Long, _
                                                               ByRef lColorRef As Long) As Long

Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, _
                                                    ByVal nWidth As Long, _
                                                    ByVal crColor As Long) As Long

Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, _
                                                         ByRef lpRect As RECT) As Long

Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, _
                                                       ByVal X As Long, _
                                                       ByVal Y As Long) As Long

Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, _
                                                    ByVal nBkMode As Long) As Long

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                      ByVal lpStr As String, _
                                                                      ByVal nCount As Long, _
                                                                      ByRef lpRect As RECT, _
                                                                      ByVal wFormat As Long) As Long

Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, _
                                                      ByVal xLeft As Long, _
                                                      ByVal yTop As Long, _
                                                      ByVal hIcon As Long, _
                                                      ByVal cxWidth As Long, _
                                                      ByVal cyWidth As Long, _
                                                      ByVal istepIfAniCur As Long, _
                                                      ByVal hbrFlickerFreeDraw As Long, _
                                                      ByVal diFlags As Long) As Long

Private Declare Function DrawState Lib "user32.dll" Alias "DrawStateA" (ByVal hdc As Long, _
                                                                        ByVal hBrush As Long, _
                                                                        ByVal lpDrawStateProc As Long, _
                                                                        ByVal lParam As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal n1 As Long, _
                                                                        ByVal n2 As Long, _
                                                                        ByVal n3 As Long, _
                                                                        ByVal n4 As Long, _
                                                                        ByVal un As Long) As Long

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, _
                                                       ByVal crColor As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function AlphaBlend Lib "msimg32" (ByVal hdcDest As Long, _
                                                   ByVal nXOriginDest As Long, _
                                                   ByVal nYOriginDest As Long, _
                                                   ByVal nWidthDest As Long, _
                                                   ByVal hHeightDest As Long, _
                                                   ByVal hdcSrc As Long, _
                                                   ByVal nXOriginSrc As Long, _
                                                   ByVal nYOriginSrc As Long, _
                                                   ByVal nWidthSrc As Long, _
                                                   ByVal nHeightSrc As Long, _
                                                   ByVal blendFunc As Long) As Boolean

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, _
                                                      Source As Any, _
                                                      ByVal Length As Long)

Private m_eLastHT                       As eHitConstants
Private m_picCaption                    As IPicture
Private m_picBorders                    As IPicture
Private m_tClientR                      As RECT
Private m_tBtn(0 To 2)                  As RECT
Private m_pCaption                      As StdPicture
Private m_pBorders                      As StdPicture
Private m_pCBoxMin                      As StdPicture
Private m_pCBoxMax                      As StdPicture
Private m_pCBoxRst                      As StdPicture
Private m_pCBoxCls                      As StdPicture
Private m_pCommand                      As StdPicture
Private m_pListvw                       As StdPicture
Private m_bActive                       As Boolean
Private m_bAppActive                    As Boolean
Private m_bBorderHasInactive            As Boolean
Private m_bStyleBit                     As Boolean
Private m_bCanSize                      As Boolean
Private m_bCanClose                     As Boolean
Private m_bControlHasInactive           As Boolean
Private m_bCustomButtonPosition         As Boolean
Private m_bIsMDIChild                   As Boolean
Private m_bMaximise                     As Boolean
Private m_bMinimise                     As Boolean
Private m_bMouseDownMinimise            As Boolean
Private m_bMouseDownMaximise            As Boolean
Private m_bMouseDownClose               As Boolean
Private m_bSkinForm                     As Boolean
Private m_bZoomedMDIChild               As Boolean
Private m_bSkinCommand                  As Boolean
Private m_bSkinLvw                      As Boolean
Private m_bTrackUser32                  As Boolean
Private m_bStyleChange                  As Boolean
Private m_bMouseMve                     As Boolean
Private m_bBtnChnge                     As Boolean
Private m_lBottom                       As Long
Private m_lBottomBorderHeight           As Long
Private m_lButtonWidth                  As Long
Private m_lButtonHeight                 As Long
Private m_lCaptionHeight                As Long
Private m_CmdFontClr                    As Long
Private m_lControlOffsetX               As Long
Private m_lControlOffsetY               As Long
Private m_hDC                           As Long
Private m_hWnd                          As Long
Private m_hWndMDIClient                 As Long
Private m_lLeftBorderWidth              As Long
Private m_lLeft                         As Long
Private m_lLeftEnd                      As Long
Private m_lOffset                       As Long
Private m_lRightStart                   As Long
Private m_lRightEnd                     As Long
Private m_lRight                        As Long
Private m_lRightBorderWidth             As Long
Private m_lTop                          As Long
Private m_lTopBorderHeight              As Long
Private m_hBmp                          As Long
Private m_hBmpOld                       As Long
Private m_lDc()                         As Long
Private m_lWd()                         As Long
Private m_lHt()                         As Long
Private m_lLastPosTop                   As Long
Private m_lLastPosLeft                  As Long
Private m_lCmdLst                       As Long
Private m_CapText                       As String
Private m_lCapIcon                      As Long
Private m_oFrm                          As Form
Private m_oIml                          As ImageList
Private m_SizeInfo                      As MINMAXINFO
Private m_FrameHdc                      As Long
'/* dc tracker
Private vObjArray                       As Variant
Private c_ObjTrack                      As Collection
'/* 2d subclass array
Private m_GASubclass()                  As MGSubclass
Private c_InstTrack                     As Collection
Private vInstArray                      As Variant
Private m_CurrInst                      As Long

'~ Image Properties ~

Public Property Get p_ICaption() As StdPicture
'/* caption bar image
    Set p_ICaption = m_pCaption
End Property

Public Property Set p_ICaption(Propval As StdPicture)
    Set m_pCaption = Propval
End Property

Public Property Get p_IBorders() As StdPicture
'/* border image
    Set p_IBorders = m_pBorders
End Property

Public Property Set p_IBorders(Propval As StdPicture)
    Set m_pBorders = Propval
End Property

Public Property Get p_ICBoxMin() As StdPicture
'/* control box minimize buttons
    Set p_ICBoxMin = m_pCBoxMin
End Property

Public Property Set p_ICBoxMin(Propval As StdPicture)
    Set m_pCBoxMin = Propval
End Property

Public Property Get p_ICBoxMax() As StdPicture
'/* control box maximize buttons
    Set p_ICBoxMax = m_pCBoxMax
End Property

Public Property Set p_ICBoxMax(Propval As StdPicture)
    Set m_pCBoxMax = Propval
End Property

Public Property Get p_ICBoxRst() As StdPicture
'/* control box restore buttons
    Set p_ICBoxRst = m_pCBoxRst
End Property

Public Property Set p_ICBoxRst(Propval As StdPicture)
    Set m_pCBoxRst = Propval
End Property

Public Property Get p_ICBoxCls() As StdPicture
'/* control box close buttons
    Set p_ICBoxCls = m_pCBoxCls
End Property

Public Property Set p_ICBoxCls(Propval As StdPicture)
    Set m_pCBoxCls = Propval
End Property

Public Property Get p_ICommand() As StdPicture
'/* command button image
    Set p_ICommand = m_pCommand
End Property

Public Property Set p_ICommand(Propval As StdPicture)
    Set m_pCommand = Propval
End Property


'~ subclass switches ~

Public Property Get p_SkinCommand() As Boolean
'/* skin command buttons switch
    p_SkinCommand = m_bSkinCommand
End Property

Public Property Let p_SkinCommand(Propval As Boolean)
    m_bSkinCommand = Propval
End Property


'~ form options ~

Public Property Get p_CmdFntClr() As Long
'/* cmd font color
    p_CmdFntClr = m_CmdFontClr
End Property

Public Property Let p_CmdFntClr(ByVal Propval As Long)
    m_CmdFontClr = Propval
End Property

Public Property Get p_ActiveRight() As Long
'/* caption right end
    p_ActiveRight = m_lRightStart
End Property

Public Property Let p_ActiveRight(ByVal Propval As Long)
    m_lRightStart = Propval
End Property

Public Property Get p_BorderHasInactive() As Boolean
'/* border has inactive image
    p_BorderHasInactive = m_bBorderHasInactive
End Property

Public Property Let p_BorderHasInactive(ByVal Propval As Boolean)
    m_bBorderHasInactive = Propval
End Property

Public Property Get p_BottomBorderHeight() As Long
'/* bottom sizing border
    p_BottomBorderHeight = m_lBottomBorderHeight
End Property

Public Property Let p_BottomBorderHeight(ByVal Propval As Long)
    m_lBottomBorderHeight = Propval
End Property

Public Property Get p_ButtonOffsetX() As Long
'/* control button offset horz
    p_ButtonOffsetX = m_lControlOffsetX
End Property

Public Property Let p_ButtonOffsetX(ByVal Propval As Long)
    m_lControlOffsetX = Propval
End Property

Public Property Get p_ButtonOffsetY() As Long
'/* control button offset vert
    p_ButtonOffsetY = m_lControlOffsetY
End Property

Public Property Let p_ButtonOffsetY(ByVal Propval As Long)
    m_lControlOffsetY = Propval
End Property

Public Property Get p_ControlHasInactive() As Boolean
'/* caption buttons have inactive
    p_ControlHasInactive = m_bControlHasInactive
End Property

Public Property Let p_ControlHasInactive(ByVal Propval As Boolean)
'/* control buttons have inactive image
    m_bControlHasInactive = Propval
End Property

Public Property Get p_ControlButtonPosition() As Boolean
'/* use custom button offsets
    p_ControlButtonPosition = m_bCustomButtonPosition
End Property

Public Property Let p_ControlButtonPosition(ByVal Propval As Boolean)
    m_bCustomButtonPosition = Propval
End Property

Public Property Get p_LeftEnd() As Long
'/* left caption bar end
    p_LeftEnd = m_lLeftEnd
End Property

Public Property Let p_LeftEnd(ByVal Propval As Long)
    m_lLeftEnd = Propval
End Property

Public Property Get p_Offset() As Long
'/* title offset
    p_Offset = m_lOffset
End Property

Public Property Let p_Offset(ByVal Propval As Long)
    m_lOffset = Propval
End Property

Public Property Get p_SkinForm() As Boolean
'/* form skin switch
    p_SkinForm = m_bSkinForm
End Property

Public Property Let p_SkinForm(ByVal Propval As Boolean)
    m_bSkinForm = Propval
End Property

Public Property Get p_TopBorderHeight() As Long
'/* top sizing border height
    p_TopBorderHeight = m_lTopBorderHeight
End Property

Public Property Let p_TopBorderHeight(ByVal Propval As Long)
    m_lTopBorderHeight = Propval
End Property



'////////////////////////////
'//// Private Properties ////
'////////////////////////////

Private Property Get p_WindowActive() As Boolean
'/* window status
    p_WindowActive = m_bActive
End Property

Private Property Get p_CanClose() As Boolean
'/* test if window can close

Dim lStyle As Long

    lStyle = GetWindowLong(m_hWnd, (-16))
    p_CanClose = ((lStyle And &H80000) = &H80000)

End Property

Private Property Get p_CanSize() As Boolean
'/* test if window is sizeable

Dim lStyle As Long

    lStyle = GetWindowLong(m_hWnd, (-16))
    p_CanSize = ((lStyle And &H40000) = &H40000)

End Property

Private Property Set p_FrmRef(ByRef Propval As Form)
    Set m_oFrm = Propval
End Property

Public Property Get p_ImlRef() As ImageList
'/* form reference
    p_ImlRef = m_oIml
End Property

Public Property Set p_ImlRef(ByRef Propval As ImageList)
    Set m_oIml = Propval
End Property

'/////////////////////////
'////   Entry Point  /////
'/////////////////////////

Public Sub Attach(ByRef mFrm As Object)
'/* processing hub

Dim tRect   As RECT
Dim tBmp    As BITMAP

On Error Resume Next

    '/* if we are alive, reset
    If mFrm.hwnd = m_hWnd Then Window_Detach
    Set p_FrmRef = mFrm
    '/* store hwnd
    m_hWnd = mFrm.hwnd
    '/* get caption text/icon
    Get_Caption m_hWnd
    Get_Icon m_hWnd
    '/* dc tracker
    PreInstance_ObjTrack
    '/* initialize tracking variables
    Init_Images
    m_bTrackUser32 = Function_Exported("TrackMouseEvent", "User32")
    '~ Phase 1 ~
    '~ Initialize the form ~
    If m_bSkinForm Then
        '/* test for base images
        If (p_IBorders Is Nothing) Or (p_ICaption Is Nothing) _
        Or (p_ICBoxMin Is Nothing) Or (p_ICBoxMax Is Nothing) _
        Or (p_ICBoxRst Is Nothing) Or (p_ICBoxCls Is Nothing) Then Exit Sub
        '/* get system default sizes for caption and frame
        m_lLeft = GetSystemMetrics(32)
        m_lRight = m_lLeft
        m_lCaptionHeight = GetSystemMetrics(4)
        m_lTop = m_lCaptionHeight + GetSystemMetrics(33)
        m_lBottom = GetSystemMetrics(33)
        '/* create caption/border/button image mappings
        Store_Picture p_ICaption, 0
        Store_Picture p_IBorders, Increment_Images("p_IBorders")
        '/* add two instances for borders
        Increment_Images ("Border1")
        Increment_Images ("Border2")
        '/* load them
        Store_Picture p_ICBoxMin, Increment_Images("p_ICBoxMin")
        Store_Picture p_ICBoxMax, Increment_Images("p_ICBoxMax")
        Store_Picture p_ICBoxRst, Increment_Images("p_ICBoxRst")
        Store_Picture p_ICBoxCls, Increment_Images("p_ICBoxCls")
        Set m_picCaption = p_ICaption
        Set m_picBorders = p_IBorders
        Set m_pCBoxMin = p_ICBoxMin
        Set m_pCBoxMax = p_ICBoxMax
        Set m_pCBoxRst = p_ICBoxRst
        Set m_pCBoxCls = p_ICBoxCls
        '/* get button status
        m_bCanClose = p_CanClose
        m_bCanSize = p_CanSize
        'NonClient_Attatch m_hWnd
        '/* set mdi status
        m_bIsMDIChild = Mdi_Check(m_hWnd)
    End If
    
    '/* caption buttons
    GetObjectAPI p_ICBoxMin.Handle, Len(tBmp), tBmp
    With tBmp
        m_lButtonWidth = (.bmWidth / 3)
        m_lButtonHeight = .bmHeight
    End With
    
    '/* caption bar
    GetObjectAPI p_ICaption.Handle, Len(tBmp), tBmp
    With tBmp
        m_lRightEnd = .bmWidth
    End With
    
    '/* borders
    GetObjectAPI p_IBorders.Handle, Len(tBmp), tBmp
    With tBmp
        m_lLeftBorderWidth = (.bmWidth / 7)
        m_lRightBorderWidth = (.bmWidth / 7)
    End With
    
    '~ Phase 2 ~
    '~ Client area processing ~

    '/* command button status
    If m_bSkinCommand Then
        If Not p_ICommand Is Nothing Then
            Set m_pCommand = p_ICommand
        Else
            m_bSkinCommand = False
        End If
    End If
    
    '/* validity check
    If m_bSkinCommand Or m_bSkinForm Then
        CAProcessing_Hub
    End If
    
    '~ launch me ~
    '/* set window properties
    SetForegroundWindow m_hWnd
    SetFocusAPI m_hWnd
    SendMessageLong m_hWnd, &H86, 1, 0
    RedrawWindow m_hWnd, 0, 0, &H4 Or &H1 Or &H400 Or &H80
    Window_Show mFrm
    
On Error GoTo 0

End Sub


'********************************************************************************************************
'*                         DSIE - DYNAMIC SUBCLASSING INSTANTIATION ENGINE                              *
'*                                                                                                      *
'*  This could be transported out of class to give you the ability to subclass components               *
'*  dynamically at runtime, without the need for existing declarations or fixed instancing calls.       *
'*  Objects on form are counted and 7 dimensional variant array is set to upper boundry.                *
'*  A collection object is used to fetch array index by referencing the objects hwnd as a key.          *
'*  Variant array stores object heirarchy: class/name/index/hwnd and subclass instance index.           *
'*  New instances of the MGSubclass class are generated as members of a class array, corresponding      *
'*  to the objects instance index. Messages from the MISubclass interface are forwarded to              *
'*  routines: Global_Action or Local_Action depending on the message type. Global_Action is a           *
'*  global processing hub, that applies an action to all objects in the subclass member queue.          *
'*  Local_Action uses the hwnd paramater to send a message to a specific object member.                 *
'*  The advantage of this approach is that identical message types can be attatched to multiple         *
'*  objects. Object events can then be easily tracked by hwnd to originating object.                    *
'*                                                                                                      *
'********************************************************************************************************


Private Sub CAProcessing_Hub()

'/* collect control data and create a unique
'/* subclassing instance for each control hwnd
'/* using a dynamically sized class array

Dim vItem   As Variant
Dim lCount  As Long

    '/* create our collection/variant array
    PreInstance_Tracker
    '/* initialize out class array
    PreInstance_SCEngine
    '/* enumerate controls
    Object_Allocation
    '/* build our class array
    SCEngine_Build

End Sub

Private Function PreInstance_Tracker() As Long

'/* maybe make this a dynamic allocation lator
'/* but for now, lets keep it simple..
'/* allocate n + 1 to variant array
'/* (total objects + form) and
'/* initialize tracking collection.
'/* instance tracking employs a 8D variant array:
'/* class/name/index/hwnd/instance/dc/xcoord/ycoord

Dim oObj    As Object
Dim X       As Long

    X = 1
    For Each oObj In m_oFrm
        X = X + 1
    Next oObj
    'bitmapdc, ctrldc, xcoord, ycoord
    ReDim vInstArray(0 To 7, 0 To X) As Variant
    PreInstance_Tracker = X
    Set c_InstTrack = New Collection

End Function

Private Sub PreInstance_SCEngine()

'/* instance class

    ReDim m_GASubclass(0 To 0) As MGSubclass
    
End Sub

Private Sub Create_Instance(lInst As Long)

'/* create a new subclass instance

    ReDim Preserve m_GASubclass(0 To lInst) As MGSubclass
    Set m_GASubclass(lInst) = New MGSubclass
    
End Sub

Private Sub Unload_SCEngine()

'/* unload all subclass instances

Dim lInst As Long

    For lInst = 0 To Instance_Count
        Remove_SCInstance lInst
    Next lInst
    
End Sub

Private Sub Object_Allocation()

'/* assign object params to variant array
'/* assign index/hwnd to collection.
'/* array params: class/name/index/hwnd/instance.
'/* essentialy anything with a handle could
'/* be added to the list and subclassed..

Dim lVcnt   As Long
Dim oCtrl   As Control

On Error Resume Next

    lVcnt = 0
    '/* form
    If m_bSkinForm Then
        vInstArray(0, lVcnt) = "Form"
        vInstArray(1, lVcnt) = m_oFrm.Name
        vInstArray(2, lVcnt) = 0
        vInstArray(3, lVcnt) = m_oFrm.hwnd
        vInstArray(4, lVcnt) = lVcnt
        c_InstTrack.Add lVcnt, CStr(m_oFrm.hwnd)
        lVcnt = lVcnt + 1
    End If
    
    For Each oCtrl In m_oFrm
        '/* first filter
        If oCtrl.hwnd = 0 Then GoTo Skip
        If oCtrl.Tag = "NO" Then GoTo Skip
        '/* command button
        If TypeOf oCtrl Is CommandButton Then
            If m_bSkinCommand Then
                With oCtrl
                    vInstArray(0, lVcnt) = "CommandButton"
                    vInstArray(1, lVcnt) = .Name
                    vInstArray(2, lVcnt) = .Index
                    vInstArray(3, lVcnt) = .hwnd
                    vInstArray(4, lVcnt) = lVcnt
                    c_InstTrack.Add lVcnt, CStr(.hwnd)
                    lVcnt = lVcnt + 1
                End With
            End If
        End If
Skip:
    Next oCtrl

On Error GoTo 0

End Sub

Private Sub SCEngine_Build()

'/* generate per object subclass instances and
'/* assign object messages to that sc instance

Dim vObject     As Variant
Dim lHwnd       As Long
Dim lInst       As Long
Dim lCntr       As Long

    lCntr = 0
    For Each vObject In c_InstTrack
        lHwnd = Return_Handle(lCntr)
        lInst = Return_Instance(lHwnd)
        If Not lHwnd = 0 Then
            Create_Instance lInst
            Object_Attatch lHwnd
            lCntr = lCntr + 1
        End If
    Next vObject

End Sub

Private Function Message_Store(ByVal lHwnd As Long) As Collection

'/* message to object association
'/* return message group associated
'/* with that control class

Dim sClass      As String

    Set Message_Store = New Collection
    sClass = Return_ClassName(lHwnd)
    
    Select Case sClass
    Case "Form"
        '/* remove caption bar
        Message_Store.Add &H83
        '/* draw borders and caption
        Message_Store.Add &H85
        '/* caption is redrawn
        Message_Store.Add &H86
        '/* ctrl box events
        Message_Store.Add &HA0
        Message_Store.Add &HA1
        Message_Store.Add &HA2
        Message_Store.Add &HA3
        Message_Store.Add &H84
        '/* message hack
        Message_Store.Add &H20
        '/* sys menu redraw
        Message_Store.Add &H116
        Message_Store.Add &H117
        '/* style change
        Message_Store.Add &H7D
        '/* text redraw
        Message_Store.Add &HC
        '/* active app
        Message_Store.Add &H1C
        '/* sizing
        Message_Store.Add &H214
        '/* mouse move
        Message_Store.Add &H200
        '/* mdi specific
        If m_bIsMDIChild Then
            Message_Store.Add &H5
        End If
        '/* detatch on close
        Message_Store.Add &H2
        '/* size move
        Message_Store.Add &H232&
        '/* min size
        Message_Store.Add &H24

    Case "CommandButton"
        '/* mouse move
        Message_Store.Add &H200
        '/* click down
        Message_Store.Add &H201
        '/* click up
        Message_Store.Add &H202
        '/* mouse leave
        Message_Store.Add &H2A3
        '/* paint
        Message_Store.Add &H85
        '/* focus
        Message_Store.Add &H7
        Message_Store.Add &H8
    End Select
    
End Function

Private Sub Object_Attatch(ByVal lHwnd As Long)

'/* attatch control messages

Dim lInst       As Long
Dim vMessage    As Variant

    lInst = Return_Instance(lHwnd)
    '/* with that instance
    With m_GASubclass(lInst)
        '/* collect messages associated with object
        For Each vMessage In Message_Store(lHwnd)
            '/* attatch message store
            .Attach_Message Me, lHwnd, CLng(vMessage)
        Next vMessage
    End With

End Sub

Private Sub SCEngine_Destroy()

'/* clear all objects from the message
'/* queue, and destroy class instances

Dim vObject     As Variant
Dim lHwnd       As Long
Dim lCntr       As Long
Dim lInst       As Long

    lCntr = 0
    If c_InstTrack Is Nothing Then Exit Sub
    For Each vObject In c_InstTrack
        lHwnd = Return_Handle(lCntr)
        lInst = Return_Instance(lHwnd)
        If Not lHwnd = 0 Then
            Object_Dettatch lHwnd
            Remove_SCInstance lInst
        End If
        lCntr = lCntr + 1
    Next vObject

End Sub

Private Sub Object_Dettatch(ByVal lHwnd As Long)

'/* dettatch control messages

Dim lInst       As Long
Dim vMessage    As Variant

    lInst = Return_Instance(lHwnd)
    With m_GASubclass(lInst)
        For Each vMessage In Message_Store(lHwnd)
            .Detach_Message Me, lHwnd, CLng(vMessage)
        Next vMessage
    End With

End Sub

Private Function MISubclass_WindowProc(ByVal lHwnd As Long, _
                                       ByVal iMsg As Long, _
                                       ByVal wParam As Long, _
                                       ByVal lParam As Long) As Long

'/* message forwarding
'/~ I left some things to be processed
'/~ locally, only because it is simpler
'/~ then doing exotic work arounds..

Dim lStyle          As Long
Dim lNewStyle       As Long
Dim lInst           As Long
Static lDx          As Long
Static lDy          As Long
Static bNoChange    As Boolean
Dim eHitCode        As eHitConstants

    '~ Phase 1 ~
    '~ Non Client Area ~
    m_CurrInst = Return_Instance(lHwnd)
    
    Select Case iMsg
    Case &H2, &H1C, &H85, &HA0, &HA3, &H214, &H200, &HA2
        SCMessage_Hub lHwnd, iMsg, wParam, lParam
        
    '/* min size
    Case &H24
        RtlMoveMemory ByVal lParam, m_SizeInfo, LenB(m_SizeInfo)
 
    '/* hit test
    Case &H84
        eHitCode = Get_HitCode()
        m_eLastHT = eHitCode
        MISubclass_WindowProc = eHitCode

    '/* nc mouse down
    Case &HA1
        If NCMouse_Down Then
            '/* mouse dwn flicker hack
            If lDx = 0 Then lDx = 1
            If lDy = 0 Then lDy = 1
            lDx = -1 * lDx: lDy = -1 * lDy
            mouse_event &H1, lDx, lDy, 0, 0
            Exit Function
        End If
        wParam = Get_HitCode()
        SCMessage_Hub lHwnd, iMsg, wParam, lParam
        MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
        MISubclass_WindowProc lHwnd, &H85, 0, 0
        
    '/* cursor
    Case &H20
        bNoChange = True
        lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
        lNewStyle = lStyle And Not &H10000000
        SetWindowLong m_hWnd, GWL_STYLE, lNewStyle
        MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
        If GetMenu(lHwnd) <> 0 Then SetMenu lHwnd, 0
        SetWindowLong lHwnd, GWL_STYLE, lStyle
        bNoChange = False
    
    '/* remove caption
    Case &H83
        If wParam <> 0 Then
            SCMessage_Hub lHwnd, iMsg, wParam, lParam
            MISubclass_WindowProc lHwnd, &H85, 0, 0
            MISubclass_WindowProc = &H400
        Else
            MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
        End If
        
    '/* caption redrawn
    Case &H86
        m_bActive = Not (wParam = 0)
        MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
        MISubclass_WindowProc lHwnd, &H85, 0, 0
        Window_Redraw lHwnd
        
    '/* active
    Case &H6
        MISubclass_WindowProc lHwnd, &H85, 0, 0
        MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
        
    '/* style change
    Case &H7D
        If Not bNoChange Then
            MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
            MISubclass_WindowProc lHwnd, &H85, 0, 0
        Else
            MISubclass_WindowProc = m_GASubclass(m_CurrInst).Get_OldProc(lHwnd, iMsg, wParam, lParam)
        End If
        
    '~ Phase 2 ~
    '~ Client Area ~
    
    '/* click down
    Case &H201, &H202, &H7, &H8, &H2A3
        SCMessage_Hub lHwnd, iMsg, wParam, lParam
        MISubclass_WindowProc m_hWnd, &H85, 0, 0
        
    End Select
    
End Function

Private Sub SCMessage_Hub(ByVal lHwnd As Long, _
                          ByVal lMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long)

'/* message processing core

Dim lhDC        As Long
Dim lStyle      As Long
Dim sClass      As String
Dim tRect       As RECT
Dim tWP         As WINDOWPOS
Dim tNCR        As NCCALCSIZE_PARAMS

    '/* get control class name
    sClass = Return_ClassName(lHwnd)
    
    Select Case sClass
    '/* form responses
    Case "Form"
        Select Case lMsg
        '/* active check
        Case &H1C
            m_bAppActive = Not (wParam = 0)
        
        '/* button mouse over
        Case &HA0
            NCMouse_Over
            
        '/* form mouse move
        Case &H200
            If m_bBtnChnge Then Reset_Caption
            
        '/* mouse down
        Case &HA1
        '/* display patch
        If m_bActive Then
            If m_eLastHT = HTCAPTION Then
                Window_Move
            End If
        Else
            If m_eLastHT = HTCAPTION Then
                SetForegroundWindow m_hWnd
                Window_Move
            End If
        End If
        
        '/* button up
        Case &HA2
            NCMouse_Down
            NCMouse_Up
            
        '/* caption dbl click
        Case &HA3
            If Not m_bCanSize Then Exit Sub
            lStyle = GetWindowLong(lHwnd, (-16))
            If ((lStyle And &H1000000) = &H1000000) Then
                Window_Post SC_RESTORE
            Else
                Window_Post SC_MAXIMIZE
            End If
            
        '/* destroy window
        Case &H2
            Window_Detach
            
        '/* window resize
        Case &H214
            If m_bSkinCommand Then Command_Spoke 0, 0, True

        '/* paint
        Case &H85
            If IsWindowVisible(m_hWnd) = 0 Then Exit Sub
            If m_bIsMDIChild And IsZoomed(lHwnd) > 0 Then Exit Sub
            lhDC = GetWindowDC(m_hWnd)
            GetWindowRect m_hWnd, tRect
            With tRect
                OffsetRect tRect, -.Left, -.Top
                Call_Paint lhDC, .Left, .Top, .Right, .Bottom
            End With
            '/* draw icon and caption text
            Set_Caption lhDC
            Set_Icon lhDC
            ReleaseDC lHwnd, lhDC
          
        '/* remove caption
        Case &H83
            m_bZoomedMDIChild = (m_bIsMDIChild And (IsZoomed(lHwnd) <> 0))
            CopyMemory tNCR, ByVal lParam, Len(tNCR)
            CopyMemory tWP, ByVal tNCR.lppos, Len(tWP)
            '/* get new size
            With tNCR.rgrc(0)
                .Left = tWP.X
                .Top = tWP.Y
                .Right = tWP.X + tWP.cx
                .Bottom = tWP.Y + tWP.cy
                '/* get image sizes
                If Not (m_bIsMDIChild And IsZoomed(lHwnd) <> 0) Then
                    Get_LeftWidth m_lLeft
                    Get_TopHeight m_lTop
                    Get_RightWidth m_lRight
                    Get_BottomHeight m_lBottom
                End If
                '/* set size of new rectangle
                .Left = .Left + m_lLeft
                .Top = .Top + m_lTop
                .Right = .Right - m_lRight - 1
                .Bottom = .Bottom - m_lBottom - 1
            End With
            LSet tNCR.rgrc(1) = tNCR.rgrc(0)
            CopyMemory ByVal lParam, tNCR, Len(tNCR)
        End Select
        
    '/* cmd btn responses
    '/* 5 states: 0) normal 1) pushed 2) disabled 3) hover 4) focus
    Case "CommandButton"
        Select Case lMsg
        '/* mouse move
        Case &H200
            If Not m_bMouseMve Then
                Track_Mouse lHwnd
                m_bMouseMve = True
                m_lCmdLst = lHwnd
                Command_Spoke lHwnd, 3, False
            End If
            
        '/* click up
        Case &H202
            Command_Spoke lHwnd, 0, False
            
        '/* click down
        Case &H201
            Command_Spoke lHwnd, 1, False

        '/* mouse leave
        Case &H2A3
            Command_Spoke lHwnd, 0, False
            m_bMouseMve = False
            
        Case &H85
            Command_Spoke 0, 0, True

        Case &H7
            Command_Spoke lHwnd, 4, False

        Case &H8
            Command_Spoke lHwnd, 0, False

        End Select
    End Select

End Sub

Private Property Get MISubclass_MsgResponse() As EMsgResponse
'/* message status property

    Select Case m_GASubclass(m_CurrInst).p_CurrentMessage
    Case &H85, &HA1, &HA3, &H117, &H116, &H20, &H7D, &HC, &H84, &H5, &H211, &H212, &H216
        MISubclass_MsgResponse = emrConsume
    Case Else
        MISubclass_MsgResponse = emrPreProcess
    End Select

End Property

Private Property Let MISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'<STUB>
End Property

Private Function Return_ClassName(ByVal lHwnd As Long) As String
'/* fetch classname from array

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_ClassName = vInstArray(0, lIndex)

On Error GoTo 0

End Function

Private Function Return_CtrlName(ByVal lHwnd As Long) As String
'/* fetch control name from array

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_CtrlName = vInstArray(1, lIndex)

On Error GoTo 0

End Function

Private Function Return_CtrlIndex(ByVal lHwnd As Long) As Long
'/* fetch ctrl index from array

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_CtrlIndex = vInstArray(2, lIndex)

On Error GoTo 0

End Function

Private Function Return_Handle(ByVal lInst As Long) As Long
'/* fetch classname from array

Dim lIndex As Long

On Error Resume Next

    Return_Handle = vInstArray(3, lInst)

On Error GoTo 0

End Function

Private Function Return_Instance(ByVal lHwnd As Long) As Long
'/* fetch class instance from array

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_Instance = vInstArray(4, lIndex)

On Error GoTo 0

End Function

Private Function Return_CntDc(ByVal lHwnd As Long) As Long
'/* fetch ctrl dc

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_CntDc = vInstArray(5, lIndex)

On Error GoTo 0

End Function

Private Function Return_XCoord(ByVal lHwnd As Long) As Long
'/* fetch bmp width

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_XCoord = vInstArray(6, lIndex)

On Error GoTo 0

End Function

Private Function Return_YCoord(ByVal lHwnd As Long) As Long
'/* fetch bmp height

Dim lIndex As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    Return_YCoord = vInstArray(7, lIndex)

On Error GoTo 0

End Function

Private Function PostInstance_Tracker(lHwnd As Long) As Long
'/* remove object reference from
'/* tracker collection

    c_InstTrack.Remove (CStr(lHwnd))
    
End Function

Private Function Instance_Count() As Long
'/* get number of concurrent running
'/* subclass instances

    Instance_Count = UBound(m_GASubclass)

End Function

Private Sub Remove_SCInstance(ByVal lInst As Long)
'/* remove selected instance

    Set m_GASubclass(lInst) = Nothing
    
End Sub

Public Sub Window_Detach()
'/* unload message queue

    SCEngine_Destroy
    Set c_InstTrack = Nothing
    ObjTrack_Destroy

End Sub

Private Sub Test_Build()

'/* look for issues..

Dim vObject     As Variant
Dim lHwnd       As Long
Dim lInst       As Long
Dim lCntr       As Long

    lCntr = 0
    If c_InstTrack Is Nothing Then Exit Sub
    For Each vObject In c_InstTrack
        Debug.Print "class: " & vInstArray(0, lCntr)
        Debug.Print "name: " & vInstArray(1, lCntr)
        Debug.Print "indx: " & vInstArray(2, lCntr)
        Debug.Print "hwnd: " & vInstArray(3, lCntr)
        Debug.Print "inst: " & vInstArray(4, lCntr)
        Debug.Print "dc: " & vInstArray(5, lCntr)
        Debug.Print "xc: " & vInstArray(6, lCntr)
        Debug.Print "yc: " & vInstArray(7, lCntr)
        lCntr = lCntr + 1
    Next vObject
    
End Sub


'********************************************************************************************************
'*                                      RENDER PROCESSING SPOKES                                        *
'*                                                                                                      *
'*  Drawing subroutines (spokes), and their supporting call outs. Control spokes are called from        *
'*  the global message processor Message_Hub, out to the object specific spoke. painting routines       *
'*  are reused between spokes for bitmap generation and painting, text out, and icon painting.          *
'*  Bitmap DC, X and Y coordinates are tracked through the vInstTrack array, and images are             *
'*  regenerated only if destination control size changes.                                               *
'*                                                                                                      *
'********************************************************************************************************


Public Sub Command_Spoke(ByVal lHwnd As Long, _
                         ByVal lState As Long, _
                         ByVal bPntAll As Boolean)

'/* overdraw buttons
'/* 5 states: 0) normal 1) pushed 2) disabled 3) hover 4) focus

Dim lIndex      As Long
Dim sObj        As String
Dim tRDst       As RECT
Dim tDCore      As tDrawCore
Dim oCmd        As Control

On Error Resume Next

    For Each oCmd In m_oFrm
    If Not TypeOf oCmd Is CommandButton Then GoTo Skip
        With oCmd
            If .Width = 0 Then GoTo Skip
            sObj = Return_CtrlName(lHwnd)
            lIndex = Return_CtrlIndex(lHwnd)
            If (.Name = sObj) And (Not (.Index = lIndex)) _
            And (Not (bPntAll)) Then GoTo Skip
            If bPntAll Then lState = 0
            If Not oCmd.Enabled Then lState = 2
            '/* image params
            With tRDst
                .Bottom = oCmd.Height / Screen.TwipsPerPixelY
                .Left = 0
                .Right = (oCmd.Width / Screen.TwipsPerPixelX) * 5
                .Top = 0
            End With
            '/* resize test
            Test_Bitmap m_pCommand, tRDst, oCmd.hwnd, 5
            '/* render structure
            With tDCore
                .hwnd = oCmd.hwnd
                .hdc = GetDC(oCmd.hwnd)
                .Offset = 0
                .State = lState
                '/* image coord
                .ImgRect.Left = (tRDst.Right * lState) / 5
                .ImgRect.Top = 0
                .ImgRect.Right = (tRDst.Right * lState) + tRDst.Right
                .ImgRect.Bottom = tRDst.Bottom
                '/* rect for font/icon
                .TxtRect.Bottom = .ImgRect.Bottom
                .TxtRect.Left = 0
                .TxtRect.Right = tRDst.Right / 5
                .TxtRect.Top = 0
                '/* caption
                .Caption = oCmd.Caption
                .Font = oCmd.Font
                .FontSize = (oCmd.Font.Size * -20) / Screen.TwipsPerPixelY
                .Bold = oCmd.Font.Bold
                .Italic = oCmd.Font.Italic
                .StrikeThrough = oCmd.Font.StrikeThrough
                .Underline = oCmd.Font.Underline
                .FontColor = m_CmdFontClr
                '/* icon
                If Not InStrB(oCmd.Tag, "IL") = 0 Then
                    .IconPos = eLeft
                    .IconIdx = CLng(Mid$(oCmd.Tag, 3))
                ElseIf Not InStrB(oCmd.Tag, "IR") = 0 Then
                    .IconPos = eRight
                    .IconIdx = CLng(Mid$(oCmd.Tag, 3))
                Else
                    .IconIdx = -1
                End If
            End With
            '/ send to drawing core
            SCDraw_Core tDCore
            m_oFrm.Refresh
Skip:
        End With
    Next oCmd
            
On Error GoTo 0

End Sub

Private Sub SCDraw_Core(ByRef tDCore As tDrawCore)

'/* central control drawing hub

Dim iconExists  As Boolean
Dim lIndex      As Long
Dim tLogFont    As LOGFONT
Dim lPFont      As Long
Dim lMFont      As Long
Dim lIhandle    As Long
Dim lSrcDc      As Long
Dim lOldFont    As Long

On Error Resume Next

    With tDCore
        lIndex = c_InstTrack.Item(CStr(.hwnd))
        lSrcDc = vInstArray(lIndex, 5)
        Draw_Bitmap .hdc, lSrcDc, .ImgRect, .Offset, lIndex
    End With
    
    '/* draw icon
    If Not ((tDCore.IconIdx = -1) Or (m_oIml Is Nothing)) Then
        With tDCore
            lIhandle = m_oIml.ListImages.Item(.IconIdx).Picture.Handle
            If Not lIhandle = 0 Then
                Select Case .IconPos
                Case eLeft
                    '/* shift text left
                    .TxtRect.Left = .TxtRect.Left + 20
                    Draw_Icon .hdc, 6, (.TxtRect.Bottom - .TxtRect.Top - 16) \ 2&, _
                    lIhandle, (.State < 3)
      
                Case eRight
                    .TxtRect.Left = -20
                    Draw_Icon .hdc, .TxtRect.Right - 20, _
                    (.TxtRect.Bottom - .TxtRect.Top - 16) \ 2&, lIhandle, (.State < 3)
                End Select
            End If
        End With
    End If
    
    '/* create font structure
    With tLogFont
        .lfFaceName = tDCore.Font + Chr$(0)
        .lfHeight = tDCore.FontSize
        .lfWeight = IIf(tDCore.Bold = True, 700, 500)
        .lfItalic = IIf(tDCore.Italic = True, 1, 0)
        .lfUnderline = IIf(tDCore.Underline = True, 1, 0)
        .lfStrikeOut = IIf(tDCore.StrikeThrough = True, 1, 0)
    End With
    
    '/* write text to control
    With tDCore
        '/* create font
        lMFont = CreateFontIndirect(tLogFont)
        '/* place in dc
        lPFont = SelectObject(.hdc, lMFont)
        '/* bg transparent
        SetBkMode .hdc, 1
        '/* set color
        SetTextColor .hdc, tDCore.FontColor
        '/* draw test
        DrawText .hdc, .Caption, -1, .TxtRect, &H1 Or &H20 Or &H4
        '/* delete old font
        DeleteObject lPFont
        '/* delete old font
        lOldFont = CLng(Return_NewObj("Font"))
        If lOldFont <> 0 Then DeleteObject lOldFont
        '/* store new font
        ObjTrack_Add "Font", lMFont, NewObj
    End With

On Error GoTo 0

End Sub

Private Sub Draw_Icon(ByVal lhDC As Long, _
                      ByVal lLeft As Long, _
                      ByVal lTop As Long, _
                      ByVal lIcon As Long, _
                      Optional ByVal bEnable As Boolean = True)
'/* like it says

Dim lBrush As Long

    DrawIconEx lhDC, lLeft, lTop, lIcon, 16, 16, 0, 0, &H3
  
End Sub

Private Sub Draw_Bitmap(ByVal lDstDc As Long, _
                        ByVal lSrcDc As Long, _
                        ByRef tDest As RECT, _
                        ByVal lOffset As Long, _
                        ByVal lIndex As Long)

'/* draw bitmap

On Error Resume Next

    With tDest
        BitBlt lDstDc, 0, 0, .Right, .Bottom, lSrcDc, .Left, 0, &HCC0020
    End With
    
On Error GoTo 0
    
End Sub

Private Sub Draw_Alpha(ByVal lDstDc As Long, _
                       ByVal lDstX As Long, _
                       ByVal lDestY As Long, _
                       ByVal lDestW As Long, _
                       ByVal lDestH As Long, _
                       ByVal lSrcDc As Long, _
                       ByVal lSrcX As Long, _
                       ByVal lSrcY As Long, _
                       ByVal bAlpha As Boolean, _
                       Optional ByVal lTrans As Long)

'/* paint alpha/blit - * not used *

Dim lAlpha  As Long
Dim BF      As BLENDFUNCTION

On Error Resume Next

    With BF
        .BlendOp = &H0
        .BlendFlags = 0
        .SourceConstantAlpha = lTrans
        .AlphaFormat = 0
    End With

    RtlMoveMemory lAlpha, BF, 4
    '/* BitBlt: destdc, dx, dy, dw, dh, srcdc, sx, sy, flags
    '/* AlphaBlend:destdc, dx, dy, dw, dh, srcdc, sx, sy, sw, sh, flags
    If Not bAlpha Then
        BitBlt lDstDc, lDstX, lDestY, lDestW, lDestH, lSrcDc, lSrcX, lSrcY, vbSrcCopy
    Else
        lAlpha = lTrans * 65536
        AlphaBlend lDstDc, lDstX, lDestY, lDestW, lDestH, lSrcDc, lSrcX, lSrcY, lDestW, lDestH, lAlpha
    End If
    
On Error GoTo 0
    
End Sub

Private Sub Test_Bitmap(ByRef sPic As IPicture, _
                        ByRef tCoord As RECT, _
                        ByVal lHwnd As Long, _
                        ByVal lOffset As Long)

'/* test image dimensions

Dim lIndex  As Long

On Error Resume Next

    lIndex = c_InstTrack.Item(CStr(lHwnd))
    If (tCoord.Right <> vInstArray(lIndex, 6)) Or _
    (tCoord.Bottom <> vInstArray(lIndex, 7)) Then
        Destroy_Bitmap lIndex
        Create_Bitmap sPic, tCoord, lIndex, lOffset
    End If

On Error GoTo 0

End Sub

Private Sub Create_Bitmap(ByRef sPic As IPicture, _
                          ByRef tCoord As RECT, _
                          ByVal lIndex As Long, _
                          ByVal lOffset As Long)

'/* create a new region

Dim lhDC        As Long
Dim lhDCC       As Long
Dim lBmpOld     As Long
Dim lmDc        As Long
Dim lXC         As Long
Dim lDD         As Long
Dim lSD         As Long
Dim lSzeDst     As Long
Dim lSzeSrc     As Long
Dim tB          As BITMAP

'/* 5- dc, 6- bmpdc, 7- width, 8- height
On Error Resume Next

    GetObjectAPI sPic.Handle, Len(tB), tB
    '/* create compatable device context
    lhDCC = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
    '/* add dc to array
    lhDC = CreateCompatibleDC(lhDCC)
    '/* place bitmap into dc
    lBmpOld = SelectObject(lhDC, sPic.Handle)
    '/* create our container
    lmDc = Create_Dc(tCoord)
    
    lSzeDst = tCoord.Right / lOffset
    lSzeSrc = tB.bmWidth / lOffset
    
    '/* rebuild source image
    For lXC = 1 To lOffset
        '/* stretchblt: destdc, dx, dy, dw, dh, sourcedc, sx, sy, swidth, sheight, flags
        '/* start
        StretchBlt lmDc, lDD, 0, 3, tCoord.Bottom, lhDC, lSD, 0, _
        3, tB.bmHeight, vbSrcCopy
        '/* middle
        StretchBlt lmDc, lDD + 3, 0, (lSzeDst - 6), tCoord.Bottom, _
        lhDC, lSD + 3, 0, lSzeSrc - 6, tB.bmHeight, vbSrcCopy
        '/* end
        StretchBlt lmDc, lDD + (lSzeDst - 3), 0, 3, tCoord.Bottom, _
        lhDC, lSD + (lSzeSrc - 3), 0, 3, tB.bmHeight, vbSrcCopy
        lSD = lSzeSrc * lXC
        lDD = lSzeDst * lXC
    Next lXC
    
    '/* place in dc
    SelectObject lhDC, lBmpOld
    '/* add dimensions to tracker
    If Not lBmpOld = 0 Then
        vInstArray(lIndex, 5) = lmDc
        vInstArray(lIndex, 6) = tCoord.Right
        vInstArray(lIndex, 7) = tCoord.Bottom
    Else
        '/* invalid region - destroy
        Destroy_Bitmap lIndex
    End If

    '/* cleanup
    DeleteObject lBmpOld
    DeleteDC lhDC
    DeleteDC lhDCC

On Error GoTo 0
    
End Sub

Private Function Create_Dc(ByRef tRect As RECT) As Long
'/* create an image device context

Dim lhDC    As Long
Dim lBmpOld As Long
Dim lmDc    As Long
Dim lBmp    As Long

    lhDC = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
    lmDc = CreateCompatibleDC(lhDC)
    lBmp = CreateCompatibleBitmap(lhDC, tRect.Right, tRect.Bottom)
    lBmpOld = SelectObject(lmDc, lBmp)
    Create_Dc = lmDc
    DeleteObject lBmpOld
    
End Function

Private Sub Destroy_Bitmap(ByVal lIndex As Long)
'/* destroy existing image

Dim lBmpDc  As Long
Dim lhDC    As Long
    
    lhDC = vInstArray(lIndex, 5)
    If Not lhDC = 0 Then DeleteObject lhDC
    vInstArray(lIndex, 5) = 0
    vInstArray(lIndex, 6) = 0
    vInstArray(lIndex, 7) = 0

End Sub


'///////////////////////////
'//// Painting Routines ////
'///////////////////////////

Private Sub Init_Images()
'/* initialize image dimension arrays

    ReDim m_lDc(0 To 0) As Long
    ReDim m_lWd(0 To 0) As Long
    ReDim m_lHt(0 To 0) As Long
    
End Sub

Private Function Increment_Images(ByVal sImageName As String) As Long
'/* shift image array

On Error Resume Next

    ReDim Preserve m_lDc(0 To UBound(m_lDc) + 1) As Long
    ReDim Preserve m_lWd(0 To UBound(m_lWd) + 1) As Long
    ReDim Preserve m_lHt(0 To UBound(m_lHt) + 1) As Long
    Increment_Images = UBound(m_lDc)

On Error GoTo 0

End Function

Private Sub Draw_Window(ByVal lngHdc As Long, _
                        ByVal lHwnd As Long, _
                        ByVal lLeft As Long, _
                        ByVal lTop As Long, _
                        ByVal lRight As Long, _
                        ByVal lBottom As Long, _
                        ByVal bActive As Boolean, _
                        ByVal bAsSample As Boolean)
                         
Dim eState          As ECNCButtonStates
Dim tPos            As POINTAPI
Dim tButtons        As RECT
Dim tButtonR        As RECT
Dim tRect           As RECT
Dim bWindowActive   As Boolean
Dim lHeight         As Long
Dim lRightWidth     As Long
Dim lWidth          As Long
Dim lYpos           As Long
Dim lCap            As Long
Dim lCpEnd          As Long
Dim lBtnLeft        As Long
Dim lBorderSize     As Long
Dim lSrcY           As Long
Dim lOrgX           As Long
Dim hFnt            As Long
Dim hFntOld         As Long
Dim lStyle          As Long
Dim lhDC            As Long

On Error Resume Next

    '/* calculate border sizes
    With tRect
        .Left = lLeft
        .Top = lTop
        .Right = lRight
        .Bottom = lBottom
        Get_Width 2, (.Right - .Left)
    End With
    
    Get_Height 2, m_lHt(0)
    lhDC = Get_Dc(2)
    Get_Width 3, (m_lWd(1) * 2)
    Get_Height 3, (tRect.Bottom - tRect.Top + 1)
    
    '/* get window status
    If bAsSample Then
        bWindowActive = bActive
    Else
        bWindowActive = p_WindowActive
    End If
    
    '/* position offsets
    If Not bWindowActive Then lOrgX = m_lOffset
    
    '/* draw caption bar
    BitBlt lhDC, 0, 0, m_lLeftEnd, Get_Height(0), Get_Dc(0), lOrgX, 0, vbSrcCopy
    
    lRightWidth = (m_lRightEnd - m_lRightStart + 1)
    lCpEnd = (lRight - lLeft) - lRightWidth + 1
    If lCpEnd < lLeft + lRightWidth Then lCpEnd = lLeft + lRightWidth
    BitBlt lhDC, lCpEnd, 0, lRightWidth, Get_Height(0), Get_Dc(0), lOrgX + m_lRightStart, 0, vbSrcCopy
    
    '/* get style context
    lStyle = GetWindowLong(lHwnd, (-16))
    m_bMaximise = ((lStyle And &H10000) = &H10000)
    m_bMinimise = ((lStyle And &H20000) = &H20000)

    '/* fill in caption bar
    lCap = m_lLeftEnd
    Do
        lWidth = m_lRightStart - m_lLeftEnd
        If lWidth <= 0 Then lWidth = 16
        If (lCap + lWidth) > lCpEnd Then lWidth = lCpEnd - lCap
        BitBlt lhDC, lCap, 0, lWidth, Get_Height(0), Get_Dc(0), lOrgX + m_lLeftEnd + 1, 0, vbSrcCopy
        lCap = lCap + lWidth
    Loop While lCap < lCpEnd
    
    '/* get button position offsets
    If m_bCustomButtonPosition Then
        tButtons.Left = lRight - lLeft + m_lControlOffsetX
        tButtons.Top = m_lControlOffsetY
    Else
        tButtons.Left = lCpEnd + lRightWidth - m_lHt(1) + 4
        tButtons.Top = 5
    End If
    
    '/* calculate size/position
    tButtons.Bottom = tButtons.Top + m_lButtonHeight
    lBtnLeft = tButtons.Left
    GetCursorPos tPos
    GetWindowRect m_hWnd, tRect
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
    End With
    
    '/* get close button size/state
    If m_bCanClose Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        tButtonR.Left = lBtnLeft
        tButtonR.Right = tButtonR.Left + m_lButtonWidth + 1
        If bAsSample Then
            NCButton_Draw lhDC, 0, Up, tButtonR, bWindowActive
        Else
            LSet m_tBtn(0) = tButtonR
            eState = Up
            If m_bMouseDownClose And PtInRect(tButtonR, tPos.X, tPos.Y) Then
                eState = Down
            End If
            NCButton_Draw lhDC, 0, eState, tButtonR, bWindowActive
        End If
    End If
    
    '/* get max button size/state
    If m_bMaximise Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        With tButtonR
            .Left = lBtnLeft
            .Right = .Left + m_lButtonWidth + 1
        End With
        If bAsSample Then
            NCButton_Draw lhDC, 1, Up, tButtonR, bWindowActive
        Else
            LSet m_tBtn(1) = tButtonR
            eState = Up
            If m_bMouseDownMinimise And PtInRect(tButtonR, tPos.X, tPos.Y) Then
                eState = Down
            End If
            NCButton_Draw lhDC, 1, eState, tButtonR, bWindowActive
        End If
    End If
    
    '/* get min button size/state
    If m_bMinimise Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        With tButtonR
            .Left = lBtnLeft
            .Right = .Left + m_lButtonWidth + 1
        End With
        If bAsSample Then
            NCButton_Draw lhDC, 2, Up, tButtonR, bWindowActive
        Else
            LSet m_tBtn(2) = tButtonR
            eState = Up
            If m_bMouseDownMaximise And PtInRect(tButtonR, tPos.X, tPos.Y) Then
                eState = Down
            End If
            NCButton_Draw lhDC, 2, eState, tButtonR, bWindowActive
        End If
    End If
    
    '/* draw caption
    BitBlt lngHdc, lLeft, lTop, Get_Width(2), Get_Height(2), lhDC, 0, 0, vbSrcCopy
    lBorderSize = m_lHt(1)
    If m_bBorderHasInactive Then
        lBorderSize = lBorderSize \ 2
        If Not (bWindowActive) Then lSrcY = lSrcY + lBorderSize
    End If
    
    '/* get left border image
    BitBlt Get_Dc(3), 0, 0, lBorderSize, lBorderSize, Get_Dc(1), 0, lSrcY, vbSrcCopy
    '/* get right border image
    BitBlt Get_Dc(3), lBorderSize, 0, lBorderSize, lBorderSize, Get_Dc(1), lBorderSize * 4, lSrcY, vbSrcCopy
    
    '/* size the images
    lYpos = lBorderSize
    lHeight = lBorderSize
    Do While (lYpos < lBottom - lTop)
        BitBlt Get_Dc(3), 0, lYpos, lBorderSize * 2, lHeight, Get_Dc(3), 0, 0, vbSrcCopy
        lYpos = lYpos + lHeight
        lHeight = lHeight * 2
    Loop
    
    '/* draw left and right
    If m_lLeftBorderWidth > 0 Then
        BitBlt lngHdc, lLeft, lTop + Get_Height(0), m_lLeftBorderWidth, _
        lBottom - lTop - Get_Height(0), Get_Dc(3), 0, 0, vbSrcCopy
    Else
        BitBlt lngHdc, lLeft, lTop + Get_Height(0), lBorderSize, _
        lBottom - lTop - Get_Height(0), Get_Dc(3), 0, 0, vbSrcCopy
    End If
    
    If m_lRightBorderWidth > 0 Then
        BitBlt lngHdc, lRight - m_lRightBorderWidth - 1, _
        lTop + Get_Height(0), m_lRightBorderWidth + 1, _
        lBottom - lTop - Get_Height(0), Get_Dc(3), _
        lBorderSize + (lBorderSize - m_lRightBorderWidth - 1), 0, vbSrcCopy
    Else
        BitBlt lngHdc, lRight - lBorderSize, lTop + Get_Height(0), lWidth, _
        lBottom - lTop - Get_Height(0), Get_Dc(3), lBorderSize, 0, vbSrcCopy
    End If
    
    '/* copy bottom border
    BitBlt Get_Dc(2), 0, 0, lBorderSize, lBorderSize, Get_Dc(1), lBorderSize * 3, lSrcY, vbSrcCopy
    '/* resize it
    lCap = lBorderSize
    lWidth = lBorderSize
    Do While (lCap < lRight)
        BitBlt Get_Dc(2), lCap, 0, lWidth, lBorderSize, Get_Dc(2), 0, 0, vbSrcCopy
        lCap = lCap + lWidth
        lWidth = lWidth * 2
    Loop
    
    '/* render bottom
    BitBlt Get_Dc(2), 0, 0, lBorderSize, lBorderSize, Get_Dc(1), lBorderSize * 2, lSrcY, vbSrcCopy
    If m_bCanSize Then
        BitBlt Get_Dc(2), lRight - lLeft - lBorderSize, 0, lBorderSize, _
        lBorderSize, Get_Dc(1), lBorderSize * 6, lSrcY, vbSrcCopy
    Else
        BitBlt Get_Dc(2), lRight - lLeft - lBorderSize, 0, lBorderSize, _
        lBorderSize, Get_Dc(1), lBorderSize * 5, lSrcY, vbSrcCopy
    End If
    
    '/* draw bottom
    BitBlt lngHdc, lLeft, lBottom - lBorderSize, lRight - lLeft + 1, lBorderSize, Get_Dc(2), 0, 0, vbSrcCopy
     
    '/* refresh controls
    Refresh_Controls
    
End Sub

Private Sub Window_Redraw(lHwnd As Long)
'/* redraw window

    RedrawWindow lHwnd, ByVal 0&, 0, &H100 Or &H1 Or &H2 Or &H80

End Sub

Private Sub Window_Show(fWnd As Object)
'/* show window

On Error Resume Next

    fWnd.Refresh
    SetWindowPos m_hWnd, 0, 0, 0, 0, 0, &H2 Or &H1 Or &H4 Or &H200 Or &H20
    
On Error GoTo 0

End Sub

Private Sub Window_Post(ByVal eCmd As ECNCSysCommandConstants)
'/* post change to message queue

    PostMessage m_hWnd, &H112, eCmd, 0

End Sub

Private Sub Refresh_Controls()
'/* refresh controls during resize
'/* may want to add some more to list
'/* just experiment to find out what
'/* needs to be refreshed on resize

Dim oCtrl   As Control

On Error Resume Next

    For Each oCtrl In m_oFrm
        If oCtrl.hdc <> 0 Then
            If TypeOf oCtrl Is PictureBox Then
                oCtrl.Refresh
            ElseIf TypeOf oCtrl Is Frame Then
                oCtrl.Refresh
            ElseIf TypeOf oCtrl Is CommandButton Then
                If Not m_bSkinCommand Then
                    oCtrl.Refresh
                End If
            End If
        End If
    Next oCtrl

On Error GoTo 0

End Sub

Private Sub Call_Paint(ByVal lngHdc As Long, _
                       ByVal lLeft As Long, _
                       ByVal lTop As Long, _
                       ByVal lRight As Long, _
                       ByVal lBottom As Long)

'/* invoke paint routines

    Draw_Window lngHdc, m_hWnd, lLeft, lTop, lRight, lBottom, p_WindowActive, False

End Sub

'///////////////////////////
'//// Caption Support  /////
'///////////////////////////

Private Sub Get_Caption(lHwnd As Long)
'/* get caption text from hosted form

Dim tPnt        As POINTAPI

    m_CapText = m_oFrm.Caption
    If Len(m_CapText) = 0 Then
        m_CapText = "NeoClass V6"
    End If
        
    '/* get text height/width
    GetTextExtentPoint32 m_oFrm.hdc, m_CapText, Len(m_CapText) + 2, tPnt
    
    '/* borrowed from Paul Catons example
    '/* set minimum width
    '/* via caption/button size
    With m_SizeInfo
        '/* minimum size
        With .ptMinTrackSize
            .X = (((m_lButtonWidth + 1) * 3) - m_lControlOffsetX) + tPnt.X + 32
            .Y = 0
        End With
        '/* max pos
        With .ptMaxPosition
            .X = 0
            .Y = 0
        End With
        '/* Maximized size
        With .ptMaxSize
            .X = Screen.Width
            .Y = Screen.Height
        End With
        '/* Maximum size with re-sizing
        With .ptMaxTrackSize
            .X = Screen.Width
            .Y = Screen.Height
        End With
    End With

End Sub

Private Sub Set_Caption(lhDC As Long)
'/* draw caption text

Dim tLogFont    As LOGFONT
Dim lPFont      As Long
Dim lMFont      As Long
Dim lTmpDc      As Long
Dim lTmpBmp     As Long
Dim tRect       As RECT
Dim tPnt        As POINTAPI

    '/* create font structure based on main
    '/* form settings (inheritence provided)
    With tLogFont
        .lfFaceName = m_oFrm.Font + vbNullChar
        .lfHeight = (m_oFrm.Font.Size * -20) / Screen.TwipsPerPixelY
        .lfWeight = 700
        .lfItalic = IIf(m_oFrm.FontItalic, 1, 0)
        .lfUnderline = IIf(m_oFrm.FontUnderline, 1, 0)
        .lfStrikeOut = IIf(m_oFrm.FontStrikethru, 1, 0)
    End With
    
    '/* get text height/width
    GetTextExtentPoint32 lhDC, m_CapText, Len(m_CapText) + 2, tPnt
    '/* create font
    lMFont = CreateFontIndirect(tLogFont)
    '/* place in dc
    lPFont = SelectObject(lhDC, lMFont)
    '/* bg transparent
    SetBkMode lhDC, 1
    
    '/* set text color
    '~ why is this not working??
    SetTextColor lhDC, vbWhite
    
    '/* draw text
    With tRect
        .Right = tPnt.X + 16
        .Top = ((m_pCaption.Height / Screen.TwipsPerPixelY) - tPnt.Y) / 5
        .Bottom = .Top + tPnt.Y
        .Left = (m_lLeftBorderWidth + 18)
    End With
    DrawText lhDC, m_CapText, -1, tRect, &H1 Or &H20 Or &H4
    
    '/* delete fonts
    DeleteObject lPFont
    DeleteObject lMFont

End Sub

Private Sub Get_Icon(ByVal lHwnd As Long)
'/* get form icon handle

    m_lCapIcon = m_oFrm.Icon.Handle

End Sub

Private Sub Set_Icon(ByVal lhDC As Long)
'/* draw icon

Dim lSDc        As Long
Dim lBmp        As Long
Dim tRect       As RECT
Dim lLBrsh      As LOGBRUSH
Dim lTop        As Long
Dim lLeft       As Long

    '/* create a temp image dc
    lSDc = CreateCompatibleDC(m_oFrm.hdc)
    lBmp = CreateCompatibleBitmap(m_oFrm.hdc, 16, 16)
    SelectObject lSDc, lBmp
    DeleteObject lBmp
    tRect.Right = 16
    tRect.Bottom = 16
    
    '/* backpaint
    lLBrsh.lbColor = vbMagenta
    lBmp = CreateBrushIndirect(lLBrsh)
    FillRect lSDc, tRect, lBmp
    DeleteObject lBmp
    
    '/* draw icon
    DrawIconEx lSDc, 0, 0, m_lCapIcon, 16, 16, 0, 0, DI_NORMAL
    
    '/* paint to dc
    lTop = ((m_pCaption.Height / Screen.TwipsPerPixelY) - 24) / 4
    lLeft = (m_lLeftBorderWidth + 4)
    TransparentBlt lhDC, lLeft, lTop, 16, 16, lSDc, 0, 0, 16, 16, vbMagenta
    DeleteObject lBmp
    DeleteDC lSDc

End Sub

Private Function NCMouse_Down() As Boolean
'/* caption button image swap

Dim tPos    As POINTAPI
Dim tRect   As RECT
Dim lhDC    As Long

    GetCursorPos tPos
    GetWindowRect m_hWnd, tRect
    
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
    End With
    
    With tRect
        OffsetRect tRect, -.Left, -.Top
    End With
    
    lhDC = GetWindowDC(m_hWnd)

    If m_bCanClose Then
        If PtInRect(m_tBtn(0), tPos.X, tPos.Y) <> 0 Then
            NCButton_Draw lhDC, 0, Down, m_tBtn(0)
            m_bMouseDownClose = True
            NCMouse_Down = True
        End If
    End If
    
    If m_bMaximise Then
        If PtInRect(m_tBtn(1), tPos.X, tPos.Y) <> 0 Then
            NCButton_Draw lhDC, 1, Down, m_tBtn(1)
            m_bMouseDownMaximise = True
            NCMouse_Down = True
        End If
    End If
    
    If m_bMinimise Then
        If PtInRect(m_tBtn(2), tPos.X, tPos.Y) <> 0 Then
            NCButton_Draw lhDC, 2, Down, m_tBtn(2)
            m_bMouseDownMinimise = True
            NCMouse_Down = True
        End If
    End If

    ReleaseDC m_hWnd, lhDC

End Function

Private Function NCMouse_Over() As Boolean
'/* caption button image swap

Dim tPos            As POINTAPI
Dim tRect           As RECT
Dim lhDC            As Long
Static sPrevStat    As Long
Dim lCurrStat       As Long

    GetCursorPos tPos
    GetWindowRect m_hWnd, tRect
    lhDC = GetWindowDC(m_hWnd)
    
    '/* get cursor pos and offset
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
        If Not PtInRect(m_tBtn(0), .X, .Y) = 0 Then
            '/* swap image
            If Not sPrevStat = 1 Then
                NCButton_Draw lhDC, 0, Over, m_tBtn(0)
            End If
            lCurrStat = 1
        ElseIf Not PtInRect(m_tBtn(1), .X, .Y) = 0 Then
            If Not sPrevStat = 2 Then
                NCButton_Draw lhDC, 1, Over, m_tBtn(1)
            End If
            lCurrStat = 2
        ElseIf Not PtInRect(m_tBtn(2), .X, .Y) = 0 Then
            If Not sPrevStat = 3 Then
                NCButton_Draw lhDC, 2, Over, m_tBtn(2)
            End If
            lCurrStat = 3
        Else
            If sPrevStat > 0 Then
                NCButton_Draw lhDC, (sPrevStat - 1), Up, m_tBtn(sPrevStat - 1)
            End If
        End If
        '/* conditional reset
        If lCurrStat <> sPrevStat And sPrevStat > 0 Then
            NCButton_Draw lhDC, (sPrevStat - 1), Up, m_tBtn(sPrevStat - 1)
        End If
        sPrevStat = lCurrStat
    End With
    
    ReleaseDC m_hWnd, lhDC

End Function

Private Sub NCMouse_Up()

'/* caption button events

Dim lStyle  As Long
Dim tPos    As POINTAPI
Dim tRect   As RECT
Dim lhDC    As Long

    GetCursorPos tPos
    GetWindowRect m_hWnd, tRect
    lhDC = GetWindowDC(m_hWnd)
    
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
    End With
    
    With tRect
        OffsetRect tRect, -.Left, -.Top
    End With

    '/* close window
    If m_bCanClose Then
        If Not PtInRect(m_tBtn(0), tPos.X, tPos.Y) = 0 Then
            If m_bMouseDownClose Then
                Window_Post SC_CLOSE
            End If
        End If
    End If
    
    '/* maximize
    If m_bMaximise Then
        If Not PtInRect(m_tBtn(1), tPos.X, tPos.Y) = 0 Then
            If m_bMouseDownMaximise Then
                lStyle = GetWindowLong(m_hWnd, (-16))
                If ((lStyle And &H1000000) = &H1000000) Then
                    Window_Post SC_RESTORE
                Else
                    Window_Post SC_MAXIMIZE
                End If
            End If
        End If
    End If
    
    '/* minimize
    If m_bMinimise Then
        If Not PtInRect(m_tBtn(2), tPos.X, tPos.Y) = 0 Then
            If m_bMouseDownMinimise Then
                Window_Post SC_MINIMIZE
            End If
        End If
    End If
    
    ReleaseDC m_hWnd, lhDC
    Reset_Caption

End Sub

Private Sub Reset_Caption()

'/* reset caption bar images

Dim lhDC As Long

    lhDC = GetWindowDC(m_hWnd)
    NCButton_Draw lhDC, 0, Up, m_tBtn(0)
    NCButton_Draw lhDC, 1, Up, m_tBtn(1)
    NCButton_Draw lhDC, 2, Up, m_tBtn(2)
    m_bMouseDownMinimise = False
    m_bMouseDownMaximise = False
    m_bMouseDownClose = False
    m_bBtnChnge = False
    
End Sub

Private Sub NCButton_Draw(ByVal lhDC As Long, _
                          ByVal iIndex As Long, _
                          ByVal eState As ECNCButtonStates, _
                          ByRef tRect As RECT, _
                          Optional ByVal bForceActive As Boolean = False)

'/* draw caption bar buttons

Dim lYpos      As Long
Dim lXpos      As Long
Dim lStyle      As Long
    
    m_bBtnChnge = True
    
    '/* get offset
    If Not bForceActive Then
        If m_bControlHasInactive Then
            If Not m_bActive Then
                lXpos = lXpos + m_lOffset
            End If
        End If
    End If
    
    
    Select Case iIndex
    Case 0
        If m_bCanClose Then
            Image_Change eState, lhDC, m_lButtonWidth, _
            m_lButtonHeight, Get_Dc(7), lXpos, lYpos, tRect
        End If
        
    Case 1
        If m_bMaximise Then
            lStyle = GetWindowLong(m_hWnd, (-16))
            If ((lStyle And &H1000000) = &H1000000) Then
                Image_Change eState, lhDC, m_lButtonWidth, _
                m_lButtonHeight, Get_Dc(6), lXpos, lYpos, tRect
            Else
                Image_Change eState, lhDC, m_lButtonWidth, _
                m_lButtonHeight, Get_Dc(5), lXpos, lYpos, tRect
            End If
        End If
        
    Case 2
        If m_bMinimise Then
            Image_Change eState, lhDC, m_lButtonWidth, _
            m_lButtonHeight, Get_Dc(4), lXpos, lYpos, tRect
            m_lLastPosLeft = lXpos + tRect.Left
            m_lLastPosTop = lYpos + tRect.Top
        End If
    End Select

End Sub

Private Sub Image_Change(ByVal lIndex As Long, _
                         ByVal lDstDc As Long, _
                         ByVal lDstWidth As Long, _
                         ByVal lDstHeight As Long, _
                         ByVal lSrcDc As Long, _
                         ByVal lSrcWidth As Long, _
                         ByVal lSrcHeight As Long, _
                         ByRef tRect As RECT)

'/* tristate image swap

    Select Case lIndex
    Case 1
        lSrcWidth = lSrcWidth + m_lButtonWidth
    Case 2
        lSrcWidth = lSrcWidth + (m_lButtonWidth * 2)
    End Select

    With tRect
        BitBlt lDstDc, .Left, .Top, lDstWidth, lDstHeight, _
        lSrcDc, lSrcWidth, lSrcHeight, vbSrcCopy
    End With

End Sub

Private Function Function_Exported(ByVal sFunction As String, _
                                   ByVal sModule As String) As Boolean

'/* test for library support

Dim lModule     As Long
Dim bLoaded     As Boolean

    lModule = GetModuleHandleA(sModule)
    If lModule = 0 Then lModule = LoadLibraryA(sModule)
    If lModule Then bLoaded = True

    If lModule Then
        If GetProcAddress(lModule, sFunction) Then
            Function_Exported = True
        End If
    End If

  If bLoaded Then FreeLibrary lModule
  
End Function

Private Function Mdi_Check(ByVal lHwnd As Long) As Boolean
'/* get mdi status

Dim hWndP   As Long
Dim sBuf    As String
Dim iPos    As Long

    hWndP = GetParent(lHwnd)
    sBuf = String$(260, 0)
    GetClassName hWndP, sBuf, 259
    iPos = InStr(sBuf, vbNullChar)
    
    If iPos > 1 Then
        If Left(sBuf, iPos - 1) = "MDIClient" Then
            Mdi_Check = True
        End If
    End If

End Function

Private Function Get_HitCode() As eHitConstants
'/* retrieve event type

Dim eHitCode    As eHitConstants
Dim tPos        As POINTAPI
Dim tRect       As RECT

    GetCursorPos tPos
    GetWindowRect m_hWnd, tRect
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
    End With
    OffsetRect tRect, -tRect.Left, -tRect.Top
    Get_HitCode = HTCLIENT
    
    With tPos
        If Not (PtInRect(tRect, .X, .Y) = 0) Then
            '/* left
            If .X <= m_lLeft Then
                If .Y <= m_lBottom Then
                    If m_bCanSize Then Get_HitCode = HTTOPLEFT
                ElseIf .Y >= tRect.Bottom - m_lBottom Then
                    If m_bCanSize Then Get_HitCode = HTBOTTOMLEFT
                Else
                    If m_bCanSize Then Get_HitCode = HTLEFT
                End If
            
            '/* right
            ElseIf .X >= tRect.Right - m_lRight Then
                If .Y <= m_lBottom Then
                    If m_bCanSize Then Get_HitCode = HTTOPRIGHT
                ElseIf .Y >= tRect.Bottom - m_lBottom Then
                    If m_bCanSize Then Get_HitCode = HTBOTTOMRIGHT
                Else
                    If m_bCanSize Then Get_HitCode = HTRIGHT
                End If
            
            '/* top
            ElseIf .Y <= m_lTopBorderHeight Then
                If m_bCanSize Then Get_HitCode = HTTOP
            ElseIf .Y >= tRect.Bottom - m_lBottomBorderHeight Then
                If m_bCanSize Then Get_HitCode = HTBOTTOM
                
            '/* caption
            ElseIf .Y <= m_lTop Then
                If .Y <= m_lBottom + m_lCaptionHeight Then
                    Get_HitCode = HTCAPTION
                    If .X <= m_lCaptionHeight Then
                        Get_HitCode = HTSYSMENU
                    End If
                ElseIf .Y > m_lBottom + m_lCaptionHeight Then
                    Get_HitCode = HTCLIENT
                End If
            End If
        End If
    End With

End Function

Private Sub Window_Move()
'/* respond to window resize

Dim tPInit      As POINTAPI
Dim tPLast      As POINTAPI
Dim tPos        As POINTAPI
Dim tRect       As RECT
Dim lStyle      As Long
Dim lParent     As Long
Dim lXpos       As Long
Dim lYpos       As Long

    '/* get current window status
    lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
    '/* if not maximized get size
    If lStyle = &H1000000 Then Exit Sub
    GetWindowRect m_hWnd, tRect
    lParent = GetParent(m_hWnd)
    '/* map position
    If Not lParent = 0 Then MapWindowPoints 0, lParent, tRect, 2
        
    GetCursorPos tPInit
    LSet tPLast = tPInit
    '/* resize images with window
    Do While Not (GetAsyncKeyState(vbLeftButton) = 0) And m_bActive
        GetCursorPos tPos
        With tPos
            If Not .X = tPLast.X Or Not .Y = tPLast.Y Then
                lXpos = .X - tPLast.X
                lYpos = .Y - tPLast.Y
                SetWindowPos m_hWnd, 0, tRect.Left + lXpos, _
                tRect.Top + lYpos, 0, 0, &H1 Or &H4 Or &H200
                LSet tPLast = tPos
                GetWindowRect m_hWnd, tRect
                If Not lParent = 0 Then MapWindowPoints 0, _
                lParent, tRect, 2
            End If
        End With
        DoEvents
        Sleep 1
    Loop

End Sub

Private Function Get_Dc(ByVal lInst As Long) As Long
'/* return image handle

    Get_Dc = m_lDc(lInst)

End Function

Private Function Get_Width(ByVal lInst As Long, _
                           Optional ByVal lCoord As Long) As Long
'/* return width of image

    If lCoord > m_lWd(lInst) Then
        Picture_Create lInst, lCoord, m_lHt(lInst)
    End If
    Get_Width = m_lWd(lInst)

End Function

Private Function Get_Height(ByVal lInst As Long, _
                            Optional ByVal lCoord As Long) As Long
'/* return height of image

    If lCoord > m_lHt(lInst) Then
        Picture_Create lInst, m_lWd(lInst), lCoord
    End If
    Get_Height = m_lHt(lInst)
    
End Function


'/////////////////////////
'//// Object Tracking ////
'/////////////////////////

Private Sub PreInstance_ObjTrack()
'/* instantiate our dc tracker
'/* 4 dimensional variant array
'/* used to track object handles
'/* 0) dc 1) oldobj 2) newobj 3) reserved

    ReDim vObjArray(0 To 4, 0 To 255) As Variant
    Set c_ObjTrack = New Collection
    
    '/* start at 1 index
    c_ObjTrack.Add "0", "0"

End Sub

Private Sub ObjTrack_Add(ByVal sInstName As String, _
                         ByVal lHandle As Long, _
                         ByRef eObjType As eObjTrackType)
                         
'/* add a tracked object

Dim lIndex As Long

On Error Resume Next

    '/* get object instance
    lIndex = c_ObjTrack.Item(sInstName)

    '/* add if new
    If lIndex = 0 Then
        c_ObjTrack.Add (c_ObjTrack.Count + 1), sInstName
        lIndex = c_ObjTrack.Count
    End If
    
    '/* object action
    Select Case eObjType
    '/* dc
    Case 0
        vObjArray(0, lIndex) = lHandle
    
    '/* old object
    Case 1
        vObjArray(1, lIndex) = lHandle
        
    '/* new object
    Case 2
        vObjArray(2, lIndex) = lHandle
        
    End Select
    
On Error GoTo 0

End Sub

Private Sub ObjTrack_Destroy()

'/* destroy all tracked objects

Dim lCount  As Long
Dim vItem   As Variant

On Error Resume Next

    If c_ObjTrack Is Nothing Then Exit Sub
    For Each vItem In c_ObjTrack
    lCount = lCount + 1
        '/* dc
        If Not vObjArray(0, lCount) = 0 Then
            DeleteDC CLng(vObjArray(0, lCount))
        End If
        
        '/* old object
        If Not vObjArray(1, lCount) = 0 Then
            DeleteObject CLng(vObjArray(1, lCount))
        End If
        
        '/* new object
        If Not vObjArray(2, lCount) = 0 Then
            DeleteObject CLng(vObjArray(2, lCount))
        End If
    Next vItem
    
    '/* clean up the rest
    Set m_picCaption = Nothing
    Set m_picBorders = Nothing
    Set m_pCaption = Nothing
    Set m_pBorders = Nothing
    Set m_pCBoxMin = Nothing
    Set m_pCBoxMax = Nothing
    Set m_pCBoxRst = Nothing
    Set m_pCBoxCls = Nothing
    Set c_ObjTrack = Nothing
    Set m_pCommand = Nothing
    Set c_ObjTrack = Nothing
    Set m_oIml = Nothing
    Set m_oFrm = Nothing
    Erase vObjArray

On Error GoTo 0

End Sub

Private Function Return_Dc(ByVal sInstName As String) As Long
'/* fetch new dc

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If lIndex = 0 Then Exit Function
    Return_Dc = vObjArray(0, lIndex)

On Error GoTo 0

End Function

Private Function Return_OldObj(ByVal sInstName As String) As Long
'/* fetch old bmp

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If lIndex = 0 Then Exit Function
    Return_OldObj = vObjArray(1, lIndex)
    
On Error GoTo 0

End Function

Private Function Return_NewObj(ByVal sInstName As String) As Long
'/* fetch new bmp

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If lIndex = 0 Then Exit Function
    Return_NewObj = vObjArray(2, lIndex)
    
On Error GoTo 0

End Function

Private Sub Store_Picture(sPic As IPicture, _
                          ByVal lInst As Long)

'/* copy image to workspace

Dim tB          As BITMAP
Dim lTmpDc      As Long
Dim lNewDc      As Long
Dim lBmpOld     As Long

    '/* get pic size
    GetObjectAPI sPic.Handle, Len(tB), tB
    Get_Width lInst, tB.bmWidth
    Get_Height lInst, tB.bmHeight
    
    '/* add to temp dc
    lTmpDc = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
    lNewDc = CreateCompatibleDC(lTmpDc)
    lBmpOld = SelectObject(lNewDc, sPic.Handle)
    
    '/* copy to global dc
    BitBlt m_lDc(lInst), 0, 0, tB.bmWidth, _
    tB.bmHeight, lNewDc, 0, 0, vbSrcCopy
    
    '/* cleanup
    SelectObject lNewDc, lBmpOld
    DeleteObject lBmpOld
    DeleteDC lNewDc
    DeleteDC lTmpDc

End Sub

Private Sub Picture_Create(ByVal lInst As Long, _
                           ByVal lWidth As Long, _
                           ByVal lHeight As Long)

'/* calculate new image params

Dim lhDC    As Long
Dim lBmp    As Long
Dim lOldBmp As Long

    '/* destroy old image
    Picture_Destroy lInst
    '/* create dc
    lhDC = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
    m_lDc(lInst) = CreateCompatibleDC(lhDC)
    
    '/* add to tracking array
    ObjTrack_Add "framedc" & lInst, m_lDc(lInst), ObjDc
    
    '/* add new bmp
    lBmp = CreateCompatibleBitmap(lhDC, lWidth, lHeight)
    lOldBmp = SelectObject(m_lDc(lInst), lBmp)
    
    '/* add to tracking array
    ObjTrack_Add "framepic" & lInst, lBmp, NewObj
    
    '/* validity check
    If lOldBmp = 0 Then
        Picture_Destroy lInst
    Else
        m_lWd(lInst) = lWidth
        m_lHt(lInst) = lHeight
    End If
    
    '/* cleanup
    DeleteObject lOldBmp
    DeleteDC lhDC

End Sub

Private Sub Picture_Destroy(ByVal lInst As Long)
'/* destroy existing image

Dim lBmpOld     As Long

    lBmpOld = Return_NewObj("framepic" & lInst)
    If Not lBmpOld = 0 Then
        SelectObject m_lDc(lInst), lBmpOld
        DeleteObject lBmpOld
    End If
    
    m_lWd(lInst) = 0
    m_lHt(lInst) = 0
    
    If Not m_lDc(lInst) = 0 Then
        DeleteDC m_lDc(lInst)
        m_lDc(lInst) = 0
    End If

End Sub

Private Sub Get_BottomHeight(lBHeight As Long)
'/* get bottom height

    If m_bBorderHasInactive Then
        lBHeight = Get_Height(1) \ 2 - 1
    Else
        lBHeight = Get_Height(1) - 1
    End If

End Sub

Private Sub Get_LeftWidth(lLWidth As Long)
'/* get left border width

    If m_lLeftBorderWidth > 0 Then
        lLWidth = m_lLeftBorderWidth
    Else
        If m_bBorderHasInactive Then
            lLWidth = Get_Height(1) \ 2
        Else
            lLWidth = Get_Height(1)
        End If
    End If

End Sub

Private Sub Get_RightWidth(lRWidth As Long)
'/* get right border width

    If m_lRightBorderWidth > 0 Then
        lRWidth = m_lRightBorderWidth - 1
    Else
        If m_bBorderHasInactive Then
            lRWidth = Get_Height(1) \ 2 - 1
        Else
            lRWidth = Get_Height(1) - 1
        End If
    End If

End Sub

Private Sub Get_TopHeight(lTHeight As Long)
'/* get top height

    lTHeight = Get_Height(0)

End Sub

Private Sub Track_Mouse(ByVal lHwnd As Long)

Dim tme As TRACKMOUSEEVENT_STRUCT

    With tme
      .cbSize = Len(tme)
      .dwFlags = TME_LEAVE
      .hwndTrack = lHwnd
    End With

    If m_bTrackUser32 Then
        TrackMouseEvent tme
    Else
        TrackMouseEventComCtl tme
    End If
  
End Sub

Private Sub Class_Terminate()
'/* unload me

    Window_Detach

End Sub

