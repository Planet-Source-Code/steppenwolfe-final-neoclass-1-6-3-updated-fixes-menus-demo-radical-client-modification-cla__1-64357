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
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***************************************************************************************
'*  NeoClass!    Radical Client Modification Class Lite v. 1.6.2                       *
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

'/~ Form Properties /~
'/~ p_ActiveRight           - caption right end
'/~ p_BorderHasInactive     - border has inactive image
'/~ p_BottomSizingBorder    - bottom sizing border
'/~ p_ButtonOffsetX         - control button offset horz
'/~ p_ButtonOffsetY         - control button offset vert
'/~ p_CustomCaption         - use owner drawn caption text
'/~ p_CaptionFntClr         - caption text font color
'/~ p_CaptionOffset         - caption title offset
'/~ p_CustomIcon            - owner drawn icon
'/~ p_ControlHasInactive    - caption buttons have inactive
'/~ p_ControlButtonPosition - use custom button offsets
'/~ p_LeftEnd               - left caption bar end
'/~ p_Offset                - inactive frame offset
'/~ p_SkinForm              - skin form switch
'/~ p_TopSizingBorder       - top sizing border height


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
'/~ Button images are tristate - normal|hover|down and must be of the same size and tiled horizontally.
'/~ Button properties are:
'/~ p_ICBoxMin - picture reference to minimize buttons
'/~ p_ICBoxMax - picture reference to mazimize buttons
'/~ p_ICBoxRst - picture reference to restore buttons
'/~ p_ICBoxCls - picture reference to close buttons
'/~ p_ControlButtonPosition - use a custom offset (boolean)
'/~ p_ButtonOffsetX - button offset from left
'/~ p_ButtonOffsetY - button offset from right
'/~ Next create the frame image. There are seven rectangular slices to the frame. Each
'/~ must be equal in width and height and must be a perfect square. Ex. 10px by 10px would
'/~ need a frame image 10px high by 70px wide. The grouping is as follows:
'/~ 1) left side 2) left bottom corner 3) left bottom corner inactive 4) bottom
'/~ 5) right side 6) right bottom corner 7) right bottom corner inactive

'/~ Fixes: March 28, 2006
'/~ Barebones version. Removed DSIE and everything extraneous for speed. Base was NC 6.1.2,
'/~ with all the fixes, and as of date above, this is the latest.
'/~ Fixes: Most important: scrape into form/controls during resize is fixed.
'/~ Also additional property for caption text offset, and image size proiperties
'/~ are calculated by api automatically, and properties removed.
'/~ This is probably the version I will use in my apps..

'/~ Fixes: March 28, 2006
'/~ Flicker when loading/unloading a second form is fixed.
'/~ Scrapes, (backpaint) into form and controls during a resize is fixed.
'/~ Flicker on caption bar down is fixed.
'/~ Redraw speed on sizing/moving has been improved.
'/~ Resource leaks (on exit) are fixed.
'/~ Properties for image size are now done with api calc, and have been removed.
'/~ Properties for caption text color and offset have been added.

'/~ Well, usual spiel, for a comment or a job.. steppenwolfe_2000@yahoo.com
'/~ Cheers
'/~ John

Implements MISubclass

Private Const GWL_STYLE                 As Long = (-16)
Private Const GWL_EXSTYLE               As Long = (-20)
Private Const DI_NORMAL                 As Long = &H1 Or &H2

Private Type LOGBRUSH
    lbStyle                            As Long
    lbColor                            As Long
    lbHatch                            As Long
End Type

Private Type POINTAPI
    X                                  As Long
    Y                                  As Long
End Type

Private Type RECT
    Left                               As Long
    Top                                As Long
    Right                              As Long
    Bottom                             As Long
End Type

Private Type WINDOWPOS
    hWnd                               As Long
    hWndInsertAfter                    As Long
    X                                  As Long
    Y                                  As Long
    cx                                 As Long
    cy                                 As Long
    flags                              As Long
End Type

Private Type NCCALCSIZE_PARAMS
    rgrc(0 To 2)                       As RECT
    lppos                              As Long
End Type

Private Type BITMAP
    bmType                             As Long
    bmWidth                            As Long
    bmHeight                           As Long
    bmWidthBytes                       As Long
    bmPlanes                           As Integer
    bmBitsPixel                        As Integer
    bmBits                             As Long
End Type

Private Type LOGFONT
    lfHeight                           As Long
    lfWidth                            As Long
    lfEscapement                       As Long
    lfOrientation                      As Long
    lfWeight                           As Long
    lfItalic                           As Byte
    lfUnderline                        As Byte
    lfStrikeOut                        As Byte
    lfCharSet                          As Byte
    lfOutPrecision                     As Byte
    lfClipPrecision                    As Byte
    lfQuality                          As Byte
    lfPitchAndFamily                   As Byte
    lfFaceName                         As String * 32
End Type

Private Type BLENDFUNCTION
    BlendOp                            As Byte
    BlendFlags                         As Byte
    SourceConstantAlpha                As Byte
    AlphaFormat                        As Byte
End Type

Private Type MINMAXINFO
    ptReserved                         As POINTAPI
    ptMaxSize                          As POINTAPI
    ptMaxPosition                      As POINTAPI
    ptMinTrackSize                     As POINTAPI
    ptMaxTrackSize                     As POINTAPI
End Type

Private Enum eObjTrackType
    ObjDc = 0
    OldObj = 1
    NewObj = 2
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

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, _
                                                                          ByVal lpClassName As String, _
                                                                          ByVal nMaxCount As Long) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function MapWindowPoints Lib "user32" (ByVal hwndFrom As Long, _
                                                       ByVal hwndTo As Long, _
                                                       lppt As Any, _
                                                       ByVal cPoints As Long) As Long

Private Declare Function IsZoomed Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, _
                                               ByVal hMenu As Long) As Long

Private Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, _
                                                                     pSrc As Any, _
                                                                     ByVal ByteLen As Long)

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long

Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Any) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
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

Private Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, _
                                                    lprcUpdate As Any, _
                                                    ByVal hrgnUpdate As Long, _
                                                    ByVal fuRedraw As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

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

Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hdc As Long, _
                                                       ByVal crColor As Long) As Long

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, _
                                                      Source As Any, _
                                                      ByVal Length As Long)

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
                                              ByVal dX As Long, _
                                              ByVal dy As Long, _
                                              ByVal cButtons As Long, _
                                              ByVal dwExtraInfo As Long)


Private m_eLastHT                  As eHitConstants
Private m_picCaption               As IPicture
Private m_picBorders               As IPicture
Private m_tBtn(0 To 2)             As RECT
Private m_pCaption                 As StdPicture
Private m_pBorders                 As StdPicture
Private m_pCBoxMin                 As StdPicture
Private m_pCBoxMax                 As StdPicture
Private m_pCBoxRst                 As StdPicture
Private m_pCBoxCls                 As StdPicture
Private m_bActive                  As Boolean
Private m_bAppActive               As Boolean
Private m_bBorderHasInactive       As Boolean
Private m_bCanSize                 As Boolean
Private m_bCanClose                As Boolean
Private m_bControlHasInactive      As Boolean
Private m_bCustomButtonPosition    As Boolean
Private m_bIsMDIChild              As Boolean
Private m_bMaximise                As Boolean
Private m_bMinimise                As Boolean
Private m_bMouseDownMinimise       As Boolean
Private m_bMouseDownMaximise       As Boolean
Private m_bMouseDownClose          As Boolean
Private m_bSkinForm                As Boolean
Private m_bZoomedMDIChild          As Boolean
Private m_bBtnChnge                As Boolean
Private m_bCustomCaption           As Boolean
Private m_bCustomIcon              As Boolean
Private m_CaptionFntClr            As Long
Private m_lCaptionOffset           As Long
Private m_lBottom                  As Long
Private m_lBottomBorderHeight      As Long
Private m_lButtonWidth             As Long
Private m_lButtonHeight            As Long
Private m_lCaptionHeight           As Long
Private m_CmdFontClr               As Long
Private m_lControlOffsetX          As Long
Private m_lControlOffsetY          As Long
Private m_hWnd                     As Long
Private m_lLeftBorderWidth         As Long
Private m_lLeft                    As Long
Private m_lLeftEnd                 As Long
Private m_lOffset                  As Long
Private m_lRightStart              As Long
Private m_lRightEnd                As Long
Private m_lRight                   As Long
Private m_lRightBorderWidth        As Long
Private m_lTop                     As Long
Private m_lTopBorderHeight         As Long
Private m_lDc()                    As Long
Private m_lWd()                    As Long
Private m_lHt()                    As Long
Private m_lLastPosTop              As Long
Private m_lLastPosLeft             As Long
Private m_CapText                  As String
Private m_lCapIcon                 As Long
Private m_oFrmMain                 As Object
Private m_SizeInfo                 As MINMAXINFO
Private m_bLoading                  As Boolean
'/* dc tracker
Private vObjArray                  As Variant
Private c_ObjTrack                 As Collection
Private m_GSubclass                As MGSubclass

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

Public Property Get p_oFrmMain() As Object
    Set p_oFrmMain = m_oFrmMain
End Property

Public Property Set p_oFrmMain(ByVal Propval As Object)
    Set m_oFrmMain = Propval
End Property

'~ form options ~
Public Property Get p_CustomCaption() As Boolean
'/* use caption text from form
    p_CustomCaption = m_bCustomCaption
End Property

Public Property Let p_CustomCaption(ByVal Propval As Boolean)
    m_bCustomCaption = Propval
End Property

Public Property Get p_CaptionOffset() As Long
'/* use caption text from form
    p_CaptionOffset = m_lCaptionOffset
End Property

Public Property Let p_CaptionOffset(ByVal Propval As Long)
    m_lCaptionOffset = Propval
End Property

Public Property Get p_CaptionFntClr() As Long
'/* caption font color
    p_CaptionFntClr = m_CaptionFntClr
End Property

Public Property Let p_CaptionFntClr(ByVal Propval As Long)
    m_CaptionFntClr = Propval
End Property

'~ form options ~
Public Property Get p_CustomIcon() As Boolean
'/* use caption icon from form
    p_CustomIcon = m_bCustomIcon
End Property

Public Property Let p_CustomIcon(ByVal Propval As Boolean)
    m_bCustomIcon = Propval
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

Public Property Get p_BottomSizingBorder() As Long
'/* bottom sizing border
    p_BottomSizingBorder = m_lBottomBorderHeight
End Property

Public Property Let p_BottomSizingBorder(ByVal Propval As Long)
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

Public Property Get p_TopSizingBorder() As Long
'/* top sizing border height
    p_TopSizingBorder = m_lTopBorderHeight
End Property

Public Property Let p_TopSizingBorder(ByVal Propval As Long)
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


'/////////////////////////
'////   Entry Point  /////
'/////////////////////////

Public Sub Attach()
'/* initialization core

Dim tBmp    As BITMAP

On Error Resume Next

    '/* if we are alive, reset
    If m_oFrmMain.hWnd = m_hWnd Then
        Window_Detatch m_hWnd
    End If
    
    '/* store hwnd
    m_hWnd = m_oFrmMain.hWnd
    
    '/* caption buttons
    GetObjectAPI p_ICBoxMin.handle, Len(tBmp), tBmp
    With tBmp
        m_lButtonWidth = (.bmWidth / 3)
        m_lButtonHeight = .bmHeight
    End With
    
    '/* caption bar
    GetObjectAPI p_ICaption.handle, Len(tBmp), tBmp
    With tBmp
        m_lRightEnd = .bmWidth
    End With
    
    '/* borders
    GetObjectAPI p_IBorders.handle, Len(tBmp), tBmp
    With tBmp
        m_lLeftBorderWidth = (.bmWidth / 7)
        m_lRightBorderWidth = (.bmWidth / 7)
    End With
    
    '/* get caption text/icon
    If m_lCaptionOffset = 0 Then
        m_lCaptionOffset = 24
    End If
    
    '/* get caption text/icon
    Get_Caption m_hWnd
    Get_Icon m_hWnd
    
    '/* dc tracker
    PreInstance_ObjTrack
    
    '/* initialize tracking variables
    Init_Images
    
    '~ Phase 1 ~
    '~ Initialize the form ~
    '/* test for base images
    If (p_IBorders Is Nothing) Or (p_ICaption Is Nothing) Or _
    (p_ICBoxMin Is Nothing) Or (p_ICBoxMax Is Nothing) Or _
    (p_ICBoxRst Is Nothing) Or (p_ICBoxCls Is Nothing) Then Exit Sub
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
    
    m_bCanClose = p_CanClose
    m_bCanSize = p_CanSize
    '/* set mdi status
    m_bIsMDIChild = Mdi_Check(m_hWnd)
    
    '/* start subclassing
    If Not m_GSubclass Is Nothing Then
        Set m_GSubclass = Nothing
    End If
    Set m_GSubclass = New MGSubclass
    Window_Attatch m_hWnd

    '/* set window properties
    SetForegroundWindow m_hWnd
    SetFocusAPI m_hWnd
    SendMessageLong m_hWnd, &H86, 1, 0
    RedrawWindow m_hWnd, 0, 0, &H4 Or &H1 Or &H400 Or &H80
    Window_Show m_oFrmMain
    
On Error GoTo 0

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
Dim eHitCode        As eHitConstants
Static lDx          As Long
Static lDy          As Long
Static bNoChange    As Boolean

    Select Case iMsg
    Case &H2, &H1C, &HA0, &HA3, &H200, &HA2
        SCMessage_Hub lHwnd, iMsg, wParam, lParam
   
   '/* paint
    Case &H85
        SCMessage_Hub lHwnd, iMsg, wParam, lParam
        Window_Redraw lHwnd

    '/* sizing
    Case &H214
        Window_Redraw lHwnd
        
    '/* exit size/move
    Case &H232
        Window_Redraw lHwnd
        
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
        MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
        MISubclass_WindowProc lHwnd, &H85, 0, 0
        
    '/* init menu
    Case &H116
        '/* menu hack
        MISubclass_WindowProc lHwnd, &H85, 0, 0
    
    '/* cursor
    Case &H20
        bNoChange = True
        lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
        lNewStyle = lStyle And Not &H10000000
        SetWindowLong m_hWnd, GWL_STYLE, lNewStyle
        MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
        If GetMenu(lHwnd) <> 0 Then
            SetMenu lHwnd, 0
        End If
        SetWindowLong lHwnd, GWL_STYLE, lStyle
        bNoChange = False
        
    '/* remove caption
    Case &H83
        If wParam <> 0 Then
            MISubclass_WindowProc lHwnd, &H85, 0, 0
            SCMessage_Hub lHwnd, iMsg, wParam, lParam
            MISubclass_WindowProc lHwnd, &H85, 0, 0
            MISubclass_WindowProc = &H400
        Else
            MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
            MISubclass_WindowProc lHwnd, &H85, 0, 0
        End If
        
    '/* caption redrawn 134
    Case &H86
        MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
        MISubclass_WindowProc lHwnd, &H85, 0, 0
        Window_Redraw lHwnd

    '/* active
    Case &H6
        MISubclass_WindowProc lHwnd, &H85, 0, 0
        MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
        
    '/* style change
    Case &H7D
        If Not bNoChange Then
            MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
            MISubclass_WindowProc lHwnd, &H85, 0, 0
        Else
            MISubclass_WindowProc = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
        End If
            
    End Select

End Function

Private Sub SCMessage_Hub(ByVal lHwnd As Long, _
                          ByVal lMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long)

'/* message processing core

Dim lhDC   As Long
Dim lStyle As Long
Dim tRect  As RECT
Dim tWP    As WINDOWPOS
Dim tNCR   As NCCALCSIZE_PARAMS

    Select Case lMsg
    '/* active check
    Case &H1C
        m_bAppActive = Not (wParam = 0)
        
    '/* button mouse over
    Case &HA0
        NCMouse_Over
        
    '/* form mouse move
    Case &H200
        If m_bBtnChnge Then
            Reset_Caption
        End If
        
    '/* mouse down
    Case &HA1
        '/* display patch
        If m_eLastHT = HTCAPTION Then Window_Move
        
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
        
        
    '/* window resize
    Case &H214
    
    
        '/* paint
    Case &H85
        If IsWindowVisible(lHwnd) = 0 Then Exit Sub
        If m_bIsMDIChild And IsZoomed(lHwnd) > 0 Then Exit Sub
        lhDC = GetWindowDC(lHwnd)
        
        GetWindowRect m_hWnd, tRect
        With tRect
            OffsetRect tRect, -.Left, -.Top
            Call_Paint lhDC, .Left, .Top, .Right, .Bottom
        End With
        
        '/* use caption from form
        If m_bCustomCaption Then Set_Caption lhDC
        '/* use form icon
        If m_bCustomIcon Then Set_Icon lhDC
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

End Sub

Private Property Get MISubclass_MsgResponse() As EMsgResponse
'/* message status property

    Select Case m_GSubclass.p_CurrentMessage
    Case &H85, &HA1, &HA3, &H117, &H116, &H20, &H7D, &HC, &H84, &H5, &H211, &H212, &H216, &H86
        MISubclass_MsgResponse = emrConsume
    Case Else
        MISubclass_MsgResponse = emrPreProcess
    End Select

End Property

Private Property Let MISubclass_MsgResponse(ByVal RHS As EMsgResponse)
'<STUB>
End Property

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

'/* draw the nc caption and frame

Dim eState        As ECNCButtonStates
Dim tPos          As POINTAPI
Dim tButtons      As RECT
Dim tButtonR      As RECT
Dim tRect         As RECT
Dim lHeight       As Long
Dim lRightWidth   As Long
Dim lWidth        As Long
Dim lYpos         As Long
Dim lCap          As Long
Dim lCpEnd        As Long
Dim lBtnLeft      As Long
Dim lBorderSize   As Long
Dim lSrcY         As Long
Dim lOrgX         As Long
Dim lStyle        As Long
Dim lhDC          As Long

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

    '/* draw caption bar
    BitBlt lhDC, 0, 0, m_lLeftEnd, Get_Height(0), Get_Dc(0), lOrgX, 0, vbSrcCopy
    lRightWidth = (m_lRightEnd - m_lRightStart + 1)
    lCpEnd = (lRight - lLeft) - lRightWidth + 1
    If lCpEnd < lLeft + lRightWidth Then
        lCpEnd = lLeft + lRightWidth
    End If
    BitBlt lhDC, lCpEnd, 0, lRightWidth, Get_Height(0), Get_Dc(0), lOrgX + m_lRightStart, 0, vbSrcCopy
    
    '/* get style context
    lStyle = GetWindowLong(lHwnd, (-16))
    m_bMaximise = ((lStyle And &H10000) = &H10000)
    m_bMinimise = ((lStyle And &H20000) = &H20000)
    
    '/* fill in caption bar
    lCap = m_lLeftEnd
    Do
        lWidth = m_lRightStart - m_lLeftEnd
        If lWidth <= 0 Then
            lWidth = 16
        End If
        If (lCap + lWidth) > lCpEnd Then
            lWidth = lCpEnd - lCap
        End If
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
            NCButton_Draw lhDC, 0, Up, tButtonR
        Else
            LSet m_tBtn(0) = tButtonR
            eState = Up
            If m_bMouseDownClose And PtInRect(tButtonR, tPos.X, tPos.Y) Then
                eState = Down
            End If
            NCButton_Draw lhDC, 0, eState, tButtonR
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
            NCButton_Draw lhDC, 1, Up, tButtonR
        Else
            LSet m_tBtn(1) = tButtonR
            eState = Up
            If m_bMouseDownMinimise And PtInRect(tButtonR, tPos.X, tPos.Y) Then
                eState = Down
            End If
            NCButton_Draw lhDC, 1, eState, tButtonR
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
            NCButton_Draw lhDC, 2, Up, tButtonR
        Else
            LSet m_tBtn(2) = tButtonR
            eState = Up
            If m_bMouseDownMaximise And PtInRect(tButtonR, tPos.X, tPos.Y) Then
                eState = Down
            End If
            NCButton_Draw lhDC, 2, eState, tButtonR
        End If
    End If
    
    '/* draw caption
    BitBlt lngHdc, lLeft, lTop, Get_Width(2), Get_Height(2), lhDC, 0, 0, vbSrcCopy
    lBorderSize = m_lHt(1)
    
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
        BitBlt lngHdc, lLeft, lTop + Get_Height(0), m_lLeftBorderWidth, lBottom - lTop - Get_Height(0), Get_Dc(3), 0, 0, vbSrcCopy
    Else
        BitBlt lngHdc, lLeft, lTop + Get_Height(0), lBorderSize, lBottom - lTop - Get_Height(0), Get_Dc(3), 0, 0, vbSrcCopy
    End If
    If m_lRightBorderWidth > 0 Then
        BitBlt lngHdc, lRight - m_lRightBorderWidth - 1, lTop + Get_Height(0), m_lRightBorderWidth + 1, lBottom - lTop - Get_Height(0), Get_Dc(3), lBorderSize + (lBorderSize - m_lRightBorderWidth - 1), 0, vbSrcCopy
    Else
        BitBlt lngHdc, lRight - lBorderSize, lTop + Get_Height(0), lWidth, lBottom - lTop - Get_Height(0), Get_Dc(3), lBorderSize, 0, vbSrcCopy
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
        BitBlt Get_Dc(2), lRight - lLeft - lBorderSize, 0, lBorderSize, lBorderSize, Get_Dc(1), lBorderSize * 6, lSrcY, vbSrcCopy
    Else
        BitBlt Get_Dc(2), lRight - lLeft - lBorderSize, 0, lBorderSize, lBorderSize, Get_Dc(1), lBorderSize * 5, lSrcY, vbSrcCopy
    End If
    
    '/* draw bottom
    BitBlt lngHdc, lLeft, lBottom - lBorderSize, lRight - lLeft + 1, lBorderSize, Get_Dc(2), 0, 0, vbSrcCopy

On Error GoTo 0

End Sub

Private Sub Window_Redraw(lHwnd As Long)
'/* forced repaint

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

Private Sub Get_Caption(ByVal lHwnd As Long)
'/* get caption text from hosted form

Dim tPnt    As POINTAPI

    m_CapText = m_oFrmMain.Caption
    If Len(m_CapText) = 0 Then
        m_CapText = "NeoClass V6"
    End If
    
    '/* get text height/width
    GetTextExtentPoint32 m_oFrmMain.hdc, m_CapText, Len(m_CapText) + 2, tPnt
    '/* borrowed from Paul Catons example
    '/* set minimum width
    '/* via caption/button size
    
    With m_SizeInfo
        '/* minimum size
        With .ptMinTrackSize
            .X = (((m_lButtonWidth + 1) * 3) - m_lControlOffsetX) + tPnt.X + (m_lCaptionOffset + 24)
            .Y = 10
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

Dim tLogFont As LOGFONT
Dim lPFont   As Long
Dim lMFont   As Long
Dim tRect    As RECT
Dim tPnt     As POINTAPI

    '/* create font structure based on main
    '/* form settings (inheritence provided)
    With tLogFont
        .lfFaceName = m_oFrmMain.Font + vbNullChar
        .lfHeight = (m_oFrmMain.Font.Size * -20) / Screen.TwipsPerPixelY
        .lfWeight = 700
        .lfItalic = IIf(m_oFrmMain.FontItalic, 1, 0)
        .lfUnderline = IIf(m_oFrmMain.FontUnderline, 1, 0)
        .lfStrikeOut = IIf(m_oFrmMain.FontStrikethru, 1, 0)
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
    SetTextColor lhDC, m_CaptionFntClr
    
    '/* draw text
    With tRect
        .Right = tPnt.X + m_lCaptionOffset
        .Top = ((m_pCaption.Height / Screen.TwipsPerPixelY) - tPnt.Y) / 5
        .Bottom = .Top + tPnt.Y
        .Left = (m_lLeftBorderWidth + m_lCaptionOffset)
    End With 'TRECT
    DrawText lhDC, m_CapText, -1, tRect, &H1 Or &H20 Or &H4
    
    '/* delete fonts
    DeleteObject lPFont
    DeleteObject lMFont

End Sub

Private Sub Get_Icon(ByVal lHwnd As Long)
'/* get form icon handle

    m_lCapIcon = m_oFrmMain.Icon.handle

End Sub

Private Sub Set_Icon(ByVal lhDC As Long)
'/* draw icon

Dim lSDc   As Long
Dim lBmp   As Long
Dim tRect  As RECT
Dim lLBrsh As LOGBRUSH
Dim lTop   As Long
Dim lLeft  As Long

    '/* create a temp image dc
    lSDc = CreateCompatibleDC(m_oFrmMain.hdc)
    lBmp = CreateCompatibleBitmap(m_oFrmMain.hdc, 16, 16)
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

Dim tPos  As POINTAPI
Dim tRect As RECT
Dim lhDC  As Long

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

Dim tPos         As POINTAPI
Dim tRect        As RECT
Dim lhDC         As Long
Static sPrevStat As Long
Dim lCurrStat    As Long

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
        If lCurrStat <> sPrevStat Then
            If sPrevStat > 0 Then
                NCButton_Draw lhDC, (sPrevStat - 1), Up, m_tBtn(sPrevStat - 1)
            End If
        End If
        sPrevStat = lCurrStat
    End With
    ReleaseDC m_hWnd, lhDC

End Function

Private Sub NCMouse_Up()
'/* caption button events

Dim lStyle As Long
Dim tPos   As POINTAPI
Dim tRect  As RECT
Dim lhDC   As Long

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
                          ByRef tRect As RECT)

'/* draw caption bar buttons

Dim lYpos  As Long
Dim lXpos  As Long
Dim lStyle As Long

    m_bBtnChnge = True
    
    Select Case iIndex
    Case 0
        If m_bCanClose Then
            Image_Change eState, lhDC, m_lButtonWidth, m_lButtonHeight, Get_Dc(7), lXpos, lYpos, tRect
        End If
        
    Case 1
        If m_bMaximise Then
            lStyle = GetWindowLong(m_hWnd, (-16))
            If ((lStyle And &H1000000) = &H1000000) Then
                Image_Change eState, lhDC, m_lButtonWidth, m_lButtonHeight, Get_Dc(6), lXpos, lYpos, tRect
            Else
                Image_Change eState, lhDC, m_lButtonWidth, m_lButtonHeight, Get_Dc(5), lXpos, lYpos, tRect
            End If
        End If
        
    Case 2
        If m_bMinimise Then
            Image_Change eState, lhDC, m_lButtonWidth, m_lButtonHeight, Get_Dc(4), lXpos, lYpos, tRect
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
        BitBlt lDstDc, .Left, .Top, lDstWidth, lDstHeight, lSrcDc, lSrcWidth, lSrcHeight, vbSrcCopy
    End With

End Sub

Private Function Mdi_Check(ByVal lHwnd As Long) As Boolean
'/* get mdi status

Dim hWndP As Long
Dim sBuf  As String
Dim iPos  As Long

    hWndP = GetParent(lHwnd)
    sBuf = String$(260, 0)
    GetClassName hWndP, sBuf, 259
    iPos = InStr(sBuf, vbNullChar)
    If iPos > 1 Then
        If Left$(sBuf, iPos - 1) = "MDIClient" Then
            Mdi_Check = True
        End If
    End If

End Function

Private Function Get_HitCode() As eHitConstants

Dim tPos  As POINTAPI
Dim tRect As RECT

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
                    If m_bCanSize Then
                        Get_HitCode = HTTOPLEFT
                    End If
                ElseIf .Y >= tRect.Bottom - m_lBottom Then
                    If m_bCanSize Then
                        Get_HitCode = HTBOTTOMLEFT
                    End If
                Else
                    If m_bCanSize Then
                        Get_HitCode = HTLEFT
                    End If
                End If
            '/* right
            ElseIf .X >= tRect.Right - m_lRight Then
                If .Y <= m_lBottom Then
                    If m_bCanSize Then
                        Get_HitCode = HTTOPRIGHT
                    End If
                ElseIf .Y >= tRect.Bottom - m_lBottom Then
                    If m_bCanSize Then
                        Get_HitCode = HTBOTTOMRIGHT
                    End If
                Else
                    If m_bCanSize Then
                        Get_HitCode = HTRIGHT
                    End If
                End If
            '/* top
            ElseIf .Y <= m_lTopBorderHeight Then
                If m_bCanSize Then
                    Get_HitCode = HTTOP
                End If
            ElseIf .Y >= tRect.Bottom - m_lBottomBorderHeight Then
                If m_bCanSize Then
                    Get_HitCode = HTBOTTOM
                End If
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

Dim tPInit  As POINTAPI
Dim tPLast  As POINTAPI
Dim tPos    As POINTAPI
Dim tRect   As RECT
Dim lStyle  As Long
Dim lParent As Long
Dim lXpos   As Long
Dim lYpos   As Long

    '/* get current window status
    lStyle = GetWindowLong(m_hWnd, GWL_STYLE)
    '/* if not maximized get size
    If Not lStyle = &H1000000 Then
        GetWindowRect m_hWnd, tRect
        lParent = GetParent(m_hWnd)
        '/* map position
        If Not lParent = 0 Then
            MapWindowPoints 0, lParent, tRect, 2
        End If
        GetCursorPos tPInit
        LSet tPLast = tPInit
        '/* resize images with window
        Do While Not (GetAsyncKeyState(vbLeftButton) = 0) 'And m_bActive
            GetCursorPos tPos
            With tPos
                If Not .X = tPLast.X Or Not .Y = tPLast.Y Then
                    lXpos = .X - tPLast.X
                    lYpos = .Y - tPLast.Y
                    SetWindowPos m_hWnd, 0, tRect.Left + lXpos, tRect.Top + lYpos, 0, 0, &H1 Or &H4 Or &H200
                    LSet tPLast = tPos
                    GetWindowRect m_hWnd, tRect
                    If Not lParent = 0 Then
                        MapWindowPoints 0, lParent, tRect, 2
                    End If
                End If
            End With
            DoEvents
            Sleep 1
        Loop
    End If

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

Dim lCount As Long
Dim vItem  As Variant

On Error Resume Next

    If Not c_ObjTrack Is Nothing Then
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
    End If
        
    '/* clean up the rest
    Set m_picCaption = Nothing
    Set m_picBorders = Nothing
    Set m_pCaption = Nothing
    Set m_pBorders = Nothing
    Set m_pCBoxMin = Nothing
    Set m_pCBoxMax = Nothing
    Set m_pCBoxRst = Nothing
    Set m_pCBoxCls = Nothing
    Set m_oFrmMain = Nothing
    Set c_ObjTrack = Nothing
    Erase vObjArray
    Set m_oFrmMain = Nothing
    
On Error GoTo 0

End Sub

Private Function Return_Dc(ByVal sInstName As String) As Long
'/* fetch new dc

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If Not lIndex = 0 Then
        Return_Dc = vObjArray(0, lIndex)
    End If

On Error GoTo 0

End Function

Private Function Return_OldObj(ByVal sInstName As String) As Long
'/* fetch old bmp

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If Not lIndex = 0 Then
        Return_OldObj = vObjArray(1, lIndex)
    End If

On Error GoTo 0

End Function

Private Function Return_NewObj(ByVal sInstName As String) As Long

'/* fetch new bmp

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If Not lIndex = 0 Then
        Return_NewObj = vObjArray(2, lIndex)
    End If

On Error GoTo 0

End Function

Private Sub Store_Picture(sPic As IPicture, _
                          ByVal lInst As Long)

'/* copy image to workspace

Dim tB      As BITMAP
Dim lTmpDc  As Long
Dim lNewDc  As Long
Dim lBmpOld As Long

    '/* get pic size
    GetObjectAPI sPic.handle, Len(tB), tB
    Get_Width lInst, tB.bmWidth
    Get_Height lInst, tB.bmHeight
    
    '/* add to temp dc
    lTmpDc = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
    lNewDc = CreateCompatibleDC(lTmpDc)
    lBmpOld = SelectObject(lNewDc, sPic.handle)
    
    '/* copy to global dc
    BitBlt m_lDc(lInst), 0, 0, tB.bmWidth, tB.bmHeight, lNewDc, 0, 0, vbSrcCopy
    
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

Private Sub Window_Detatch(ByVal lHwnd As Long)
'/* detatch messages

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        .Detach_Message Me, lHwnd, &H83
        .Detach_Message Me, lHwnd, &H85
        .Detach_Message Me, lHwnd, &H86
        .Detach_Message Me, lHwnd, &HA0
        .Detach_Message Me, lHwnd, &HA1
        .Detach_Message Me, lHwnd, &HA2
        .Detach_Message Me, lHwnd, &HA3
        .Detach_Message Me, lHwnd, &H84
        .Detach_Message Me, lHwnd, &H20
        .Detach_Message Me, lHwnd, &H7D
        .Detach_Message Me, lHwnd, &H1C
        .Detach_Message Me, lHwnd, &H214
        .Detach_Message Me, lHwnd, &H200
        If m_bIsMDIChild Then
            .Detach_Message Me, lHwnd, &H5
        End If
        .Detach_Message Me, lHwnd, &H2
        .Detach_Message Me, lHwnd, &H232&
        .Detach_Message Me, lHwnd, &H24
        .Detach_Message Me, lHwnd, &H6
        .Detach_Message Me, lHwnd, &H116
    End With

End Sub

Private Sub Window_Attatch(ByVal lHwnd As Long)

'/* attatch messages

    If lHwnd = 0 Then Exit Sub
    With m_GSubclass
        '/* remove caption bar
        .Attach_Message Me, lHwnd, &H83 '*
        '/* draw borders and caption
        .Attach_Message Me, lHwnd, &H85 '*
        '/* caption is redrawn
        .Attach_Message Me, lHwnd, &H86 '*
        '/* ctrl box events
        .Attach_Message Me, lHwnd, &HA0 '*
        .Attach_Message Me, lHwnd, &HA1 '*
        .Attach_Message Me, lHwnd, &HA2 '*
        .Attach_Message Me, lHwnd, &HA3 '*
        .Attach_Message Me, lHwnd, &H84 '*
        '/* message hack
        .Attach_Message Me, lHwnd, &H20 '*
        '/* style change
        .Attach_Message Me, lHwnd, &H7D '*
        '/* active app
        .Attach_Message Me, lHwnd, &H1C '*
        '/* sizing
        .Attach_Message Me, lHwnd, &H214 '*
        '/* mouse move
        .Attach_Message Me, lHwnd, &H200 '*
        '/* mdi specific
        If m_bIsMDIChild Then
            .Attach_Message Me, lHwnd, &H5
        End If
        '/* detatch on close
        .Attach_Message Me, lHwnd, &H2 '*
        '/* size move
        .Attach_Message Me, lHwnd, &H232& '*
        '/* min size
        .Attach_Message Me, lHwnd, &H24 '*
        .Attach_Message Me, lHwnd, &H6 '*
        .Attach_Message Me, lHwnd, &H116
    End With

End Sub

Private Sub Class_Terminate()
'/* unload me

On Error Resume Next

    If m_hWnd <> 0 Then
        ObjTrack_Destroy
        Window_Detatch m_hWnd
        Set m_GSubclass = Nothing
    End If

On Error GoTo 0

End Sub


