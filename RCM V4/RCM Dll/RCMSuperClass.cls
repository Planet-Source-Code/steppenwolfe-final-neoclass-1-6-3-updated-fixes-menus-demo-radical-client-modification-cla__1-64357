VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RCMSuperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'***************************************************************************************
'*  RCM Dll!     Radical Client Modification SuperClass v. 1.4                         *
'*                                                                                     *
'*  Started:     March 03, 2006                                                        *
'*  Updated:     March 26 2006                                                         *
'*  Purpose:     Form Skinning Super Class with Transparent Styles                     *
'*  Functions:   (listed)                                                              *
'*  Revision:    1.4                                                                   *
'*  Compile:     Native                                                                *
'*  Author:      John Underhill (Steppenwolfe)                                         *
'*  Credit:      Inspired by Neocaption written by Steve McMahon (aka -the dude-)      *
'*                                                                                     *
'***************************************************************************************


'*~ List of exposed functions ~*

'/~ Set_Host                - attatch the skin
'/~ Create_Shadow           - instantiate shadow effect
'/~ Class_Unload            - unload app

'*~ List of exposed properties ~*

'/~ Image Properties /~
'/~ p_IBorders              - border image
'/~ p_ICaption              - caption bar image
'/~ p_ICBoxMin              - control box minimize
'/~ p_ICBoxMax              - control box maximize
'/~ p_ICBoxRst              - control box restore
'/~ p_ICBoxCls              - control box close

'/~ Form Properties /~
'/~ p_FrmMain               - main form reference
'/~ p_FrmHost               - host frame form reference
'/~ p_ActiveRight           - caption right end
'/~ p_BorderHasInactive     - border has inactive image
'/~ p_BottomSizingBorder    - bottom sizing border
'/~ p_ButtonHeight          - control button height
'/~ p_ButtonWidth           - control button width
'/~ p_ButtonOffsetX         - control button offset horz
'/~ p_ButtonOffsetY         - control button offset vert
'/~ p_ControlButtonPosition - use custom button offsets
'/~ p_LeftEnd               - left caption bar end
'/~ p_LeftBorderWidth       - left border width
'/~ p_Offset                - title offset
'/~ p_RightEnd              - caption bar right end
'/~ p_RightBorderWidth      - right border width
'/~ p_SkinForm              - skin form switch
'/~ p_TopSizingBorder       - top sizing border height

'/~ Extended Properties /~
'/~ p_TransIdx              - transparency index (100 - 255)
'/~ p_SkinApp               - skin an application (simple example)
'/~ p_SkinAppPath           - path to skinned application
'/~ p_ShadowForm            - enable form shadow
'/~ p_ShadowInset           - hosted form shadow
'/~ p_ShadowDepth           - form shadow depth (2 - 15)
'/~ p_ShadowInDepth         - inset shadow depth (2 - 15)


'/*
'/* Preamble
'/~ Win98/ME Support
'/~ There is none, get over it. The fact is, that there are some things that just
'/~ can not be done on these old platforms, I tried, but it proved to be impossible.
'/~ I was recently looking at a number of AV clients, and they all had something
'/~ in common, for support on the older platforms, you had to download last years version.
'/~ With Vista in pre-release, companies, (and developers), are soon going to stop supporting
'/~ these older platforms entirely, so if you are still using it, time to switch.

'/~ First Release
'/~ I killed it, so you really want to take another look at this version, as many issues
'/~ have been addressed, and this version is less rubbery/flickery by far. It is
'/~ however an ongoing work, with a few miles still to travel. When I first had the idea for
'/~ creating a skinner with translucent caption bar and frame, I had little idea how hard it
'/~ was really going to be. I searched the internet for several days, but could not find any (stable)
'/~ example, in any language, with trans caption and opaque client area, so what you are looking at
'/~ is new, and will take some time to get right.

'/~ Usage
'/~ I would use this only on small fixed forms with no resize property. The larger the form, the
'/~ more the flicker as api uses up cycles rendering alpha DIB. Compile it. Works faster compiled.
'/~ And of course, the faster your box, the better it will work.
'/~ Use darker frames. Not sure why, but they seem to hide/work better. Also, keep frames small,
'/~ under 10 px is best.

'/~ Making a skin
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
'/~ Once you have the images, save them as 32bit PNG files. Then you need to convert them to
'/~ 32bit bitmaps. You could save them as 24bit bitmaps, but 32 allows you to preserve transparencies
'/~ in the image which will be rendered to the form.
'/~ I am using the AlphaConv tool (excellent tool), which converts images with a simple
'/~ drag and drop. Link: http://www.najmanowicz.com/temp/AlphaConv.zip

'/*
'/~ And now for something completely different..
'/~ After (rewriting Steve McMahon's) Neocaption last month, I decided to take another look,
'/~ (and grudgingly reinstall..) Windows Blinds. I then took a stroll through
'/~ the wincustomize.com website, browsing for a skin that wasn't too dark/pale/oddball
'/~ for my pc. It was at that time that I noticed some of the newer skins utilized
'/~ something I had never seen before, that is; translucent caption bar and frames
'/~ --with-- the forms client area rendered entirely opaque.
'/~ This has been somewhat of a holy grail to form skinners for sometime,
'/~ it is not supported by api directly, (!not at all!), and the api that
'/~ is out there (alphablend/gdi), is just too slow to be practical. But somehow..
'/~ WB v5.0 managed to pull it off.. "So how the hell did they do they do that?!"
'/~ --I was wondering, and what I came to realize is that they are using multiple forms with
'/~ per-pixel alphablending and Device Independant Bitmaps..
'/~ Now, having some idea how something is done, and realizing it from scratch, are
'/~ two very different things! I have really never focused on the subclassing/drawing
'/~ end of things, so I pretty much had to learn it all as I went, (and there was a lot to learn!)
'/~ What I ended up doing was back-buffering a skin capture, using this to create a 32 bit DIB
'/~ and rendering this to the form using updatelayeredwindows api.
'/~ There were a lot of problems to deal with, like interactive caption buttons, syncing forms
'/~ in realtime, and rendering speed/alphablending being the most daunting.
'/~ The base images were converted from 32 bit PNG files to BMP32 with a handy tool I found
'/~ (somewhere around?) here: http://www.wincustomize.com/Articles.aspx
'/~ This gives a better quality image, and allows you to use images with existing levels of
'/~ transparency.
'/~ There are a number of real advantages, (as well as a few caveats), to using ULW/Alphablending.
'/~ The biggest advantage is that the need to subclass paint messages to constantly
'/~ repaint the form becomes completely unnecessary. Because the DIB requires no stationary redraw
'/~ you can keep the cpu cycles down, (paint message is not even hooked in this example..).
'/~ Also, some very cool effects are now possible using this technique, most noticibly the
'/~ transparent caption/frame a la Vista stylings.. Also, you are not stuck with a single
'/~ set of styles, but with a little sampling and some modest drawing skills, you can
'/~ make your forms look any way you want.. (even like Aero!)
'/~ You might want to elicit the help of some of the skinners on wincustomize, for a mere mention
'/~ in your about page, most of them would be happy to lend/help you create a new style..
'/~ The only real downside is that it is 2K/XP only, and the ULW call is kind of slow,
'/~ so syncing the forms (messages are out of sync with actual size state because of api lag),
'/~ was a bit of a problem. I think it worked out well with this example, forms are a
'/~ little rubbery, (but, --note-- they are this way in win blinds app as well..), and some
'/~ improvement could be made, (perhaps using an api form ex. Paul Catons window class), might
'/~ be the next step, also moving from ulw to gdi for backwards compliance, (I tried to make it 98/ME
'/~ compatable, but was proving to be too difficult..)..  [Paul/Carles, you reading this? Would be challenging..])
'/~ still a few little things to work out, like caption text color, button down image swap, icon dissapearing trick,
'/~ and a few [other] options I may add..
                                                                                                                '/*
'/~ Update V1.1
'/~ I released a bit prematurely, syncing issues have now been (more or less) conquered.
'/~ Next step will be with Transition_Gate routines, where form is hidden while in motion and
'/~ image of form is in its place, has proven to be almost flicker free at an size, no matter how
'/~ much form is abused, (but has a has a bug I am working on), resize will have a similar solution.
'/~ I will look at tightening it up some more and putting version 2 out there sometime next week.
'/~ Be patient, and remember, this is complicated code, and takes time to write..

'/~ Update V1.2
'/~ Fixed the sizing thing- left and top size messages are ignored, as the combination size/move
'/~ causes too much flicker. Fixed the image swap on mouse down for caption buttons. Bunch a other
'/~ little stuff.. Also, posted the last rev of NeoClass..

'/~ Update V1.3
'/~ This is as far as I got with the project, so I thought I would 'donate' my work for
'/~ someone ambitious to follow.. Bunch a fixes/improvements.

'/~ Where do we go from here..
'/~ I am beginning to think that two forms can not be perfectly synced, no matter what method is used..
'/~ This works well with small to medium sized forms, but becomes too difficult to sync as forms are
'/~ loaded up with many controls, or sized too large.
'/~ I do, however, still think that updatelayed windows api and device independant bitmaps, are involved
'/~ in the solution to translucent frames. I'll take another look at it down the road.. If you have a
'/~ brainstorm, remember to fill me in ;o)

'/~ For a comment or a job.. steppenwolfe_2000@yahoo.com
'/~ Cheers
'/~ John


'/* subclasser
Implements MISubclass

Private Const GWL_EXSTYLE               As Long = (-20)
Private Const GWL_STYLE                 As Long = (-16)
Private Const ULW_COLORKEY              As Long = &H1&
Private Const ULW_ALPHA                 As Long = &H2&
Private Const ULW_OPAQUE                As Long = &H4&

Private Const WS_EX_LAYERED             As Long = &H80000
Private Const AC_SRC_OVER               As Long = &H0
Private Const AC_SRC_ALPHA              As Long = &H1
Private Const DI_NORMAL                 As Long = &H1 Or &H2

Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type

Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
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

Private Type BLENDFUNCTION
    BlendOp                             As Byte
    BlendFlags                          As Byte
    SourceConstantAlpha                 As Byte
    AlphaFormat                         As Byte
End Type

Private Type Size
    cx                                  As Long
    cy                                  As Long
End Type

Private Type RGBQUAD
    rgbBlue                             As Byte
    rgbGreen                            As Byte
    rgbRed                              As Byte
    rgbReserved                         As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize                              As Long
    biWidth                             As Long
    biHeight                            As Long
    biPlanes                            As Integer
    biBitCount                          As Integer
    biCompression                       As Long
    biSizeImage                         As Long
    biXPelsPerMeter                     As Long
    biYPelsPerMeter                     As Long
    biClrUsed                           As Long
    biClrImportant                      As Long
End Type

Private Type BITMAPINFO
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors                           As RGBQUAD
End Type

Private Type LOGBRUSH
    lbStyle                             As Long
    lbColor                             As Long
    lbHatch                             As Long
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

Private Type MINMAXINFO
    ptReserved                          As POINTAPI
    ptMaxSize                           As POINTAPI
    ptMaxPosition                       As POINTAPI
    ptMinTrackSize                      As POINTAPI
    ptMaxTrackSize                      As POINTAPI
End Type

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

   
Private Enum eObjTrackType
    ObjDc = 0
    OldBmp = 1
    NewBmp = 2
End Enum

Private Enum eDrawMethods
    ECNCModifyNonClientArea = 0
    ECNCUseClientArea = 1
End Enum

Private Enum ECNCSysCommandConstants
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

Private Enum eHitConstants
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

Private Enum ECNCButtonStates
    Up = &H0
    Over = &H1
    Down = &H2
End Enum

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

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

Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long

Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal hdc As Long) As Long

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

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

Private Declare Function UpdateLayeredWindow Lib "user32.dll" (ByVal hWnd As Long, _
                                                               ByVal hdcDst As Long, _
                                                               pptDst As Any, _
                                                               psize As Any, _
                                                               ByVal hdcSrc As Long, _
                                                               pptSrc As Any, _
                                                               ByVal crKey As Long, _
                                                               ByRef pblend As BLENDFUNCTION, _
                                                               ByVal dwFlags As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, _
                                                           pBitmapInfo As BITMAPINFO, _
                                                           ByVal un As Long, _
                                                           ByRef lplpVoid As Any, _
                                                           ByVal handle As Long, _
                                                           ByVal dw As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        ByVal lParam As Long) As Long

Private Declare Function PtInRect Lib "user32" (lpRect As RECT, _
                                                ByVal X As Long, _
                                                ByVal Y As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long

Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal nBkMode As Long) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal crColor As Long) As Long

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

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, _
                                                                      ByVal lpStr As String, _
                                                                      ByVal nCount As Long, _
                                                                      ByRef lpRect As RECT, _
                                                                      ByVal wFormat As Long) As Long

Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long

Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long

Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, _
                                                                                      ByVal Msg As Long, _
                                                                                      ByVal wParam As Long, _
                                                                                      ByVal lParam As Long, _
                                                                                      ByVal fuFlags As Long, _
                                                                                      ByVal uTimeout As Long, _
                                                                                      lpdwResult As Long) As Long

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, _
                                                                          ByVal nIndex As Integer) As Long

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long


Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, _
                                                          ByVal uExitCode As Long) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As Long, _
                                                                      ByVal lpWindowName As Long) As Long

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, _
                                                 ByVal hWndNewParent As Long) As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, _
                                                                lpdwProcessId As Long) As Long

Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, _
                                                 ByVal wCmd As Long) As Long

Private Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long

Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, _
                                                      Source As Any, _
                                                      ByVal Length As Long)

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (lppictDesc As PictDesc, _
                                                                      riid As Guid, _
                                                                      ByVal fown As Long, _
                                                                      ipic As IPicture) As Long


Public Event eTerminate()

Private m_bBorderHasInactive            As Boolean
Private m_bCanSize                      As Boolean
Private m_bCanClose                     As Boolean
Private m_bCustomButtonPosition         As Boolean
Private m_bMaximise                     As Boolean
Private m_bMinimise                     As Boolean
Private m_bSkinForm                     As Boolean
Private m_bTransFrame                   As Boolean
Private m_bSkinApp                      As Boolean
Private m_bInSizeMove                   As Boolean
Private m_bMouseDown                    As Boolean
Private m_bBtnChnge                     As Boolean
Private m_bInMove                       As Boolean
Private m_bShadowForm                   As Boolean
Private m_bShadowInset                  As Boolean
Private m_bForceUnload                  As Boolean
Private m_bCustomCaption                As Boolean
Private m_bCustomIcon                   As Boolean
Private m_lCaptionOffset                As Long
Private m_CaptionFntClr                 As Long
Private m_lLastPosTop                   As Long
Private m_lLastPosLeft                  As Long
Private m_lBottomBorderHeight           As Long
Private m_lButtonWidth                  As Long
Private m_lButtonHeight                 As Long
Private m_lControlOffsetX               As Long
Private m_lControlOffsetY               As Long
Private m_lMainHnd                      As Long
Private m_lLeftBorderWidth              As Long
Private m_lLeftEnd                      As Long
Private m_lRightStart                   As Long
Private m_lRightEnd                     As Long
Private m_lRightBorderWidth             As Long
Private m_lTopBorderHeight              As Long
Private m_lDc()                         As Long
Private m_lWd()                         As Long
Private m_lHt()                         As Long
Private m_AlphaTmpDc                    As Long
Private m_AlphaBtnBmp                   As Long
Private m_AlphaBtnDc                    As Long
Private m_AlphaBmp                      As Long
Private m_lLeft                         As Long
Private m_lRight                        As Long
Private m_lCaptionHeight                As Long
Private m_lCaptionImgHeight             As Long
Private m_lTop                          As Long
Private m_lBottom                       As Long
Private m_lHostHnd                      As Long
Private m_lTransIdx                     As Long
Private m_TransBtnIdx                   As Long
Private m_mainSurfaceDC                 As Long
Private m_mainSurfaceBitmap             As Long
Private m_backSurfaceDC                 As Long
Private m_lBlendFunc                    As Long
Private m_backBlendFUnc                 As Long
Private m_lCapIcon                      As Long
Private m_lShadowDepth                  As Long
Private m_lShadowInDepth                As Long
Private m_lAppHwnd                      As Long
Private m_sSkinAppPath                  As String
Private m_CapText                       As String
Private vObjArray                       As Variant
Private c_ObjTrack                      As Collection
Private m_blendFunc                     As BLENDFUNCTION
Private m_eLastHT                       As eHitConstants
Public m_oFrmMain                       As Form
Private m_ShadowRight                   As frmShadow
Private m_ShadowBottom                  As frmShadow
Private m_ShadowInRight                 As frmShadow
Private m_ShadowInBottom                As frmShadow
Private m_srcPoint                      As POINTAPI
Private m_tBtn(0 To 2)                  As RECT
Private m_windowSize                    As Size
Private m_pCaption                      As StdPicture
Private m_pBorders                      As StdPicture
Private m_pCBoxMin                      As StdPicture
Private m_pCBoxMax                      As StdPicture
Private m_pCBoxRst                      As StdPicture
Private m_pCBoxCls                      As StdPicture
Private m_pCommand                      As StdPicture
Private m_lTopOffset                    As Long
Private m_lBorderOffset                 As Long
Private m_lCurrentR                     As Long
Private m_lCurrentB                     As Long
Private m_lCurrentX                     As Long
Private m_lCurrentY                     As Long
Private m_SizeInfo                      As MINMAXINFO
Private m_lIconDc                       As Long
Private m_lCapTextDc                    As Long
Private m_bMaxSize                      As Boolean
Private m_lMaxHeight                    As Long
Private m_lMaxWidth                     As Long
Private m_GSubclass                     As MGSubclass

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

Public Property Get p_ButtonHeight() As Long
'/* control button height
    p_ButtonHeight = m_lButtonHeight
End Property

Public Property Let p_ButtonHeight(ByVal Propval As Long)
    m_lButtonHeight = Propval
End Property

Public Property Get p_ButtonWidth() As Long
'/* control button width
    p_ButtonWidth = m_lButtonWidth
End Property

Public Property Let p_ButtonWidth(ByVal Propval As Long)
    m_lButtonWidth = Propval
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

Public Property Get p_LeftBorderWidth() As Long
'/* left border width
    p_LeftBorderWidth = m_lLeftBorderWidth
End Property

Public Property Let p_LeftBorderWidth(ByVal Propval As Long)
    m_lLeftBorderWidth = Propval
End Property

Public Property Get p_CaptionOffset() As Long
'/* use caption text from form
    p_CaptionOffset = m_lCaptionOffset
End Property

Public Property Let p_CaptionOffset(Propval As Long)
    m_lCaptionOffset = Propval
End Property

Public Property Get p_CaptionHeight() As Long
'/* left border width
    p_CaptionHeight = m_lCaptionImgHeight
End Property

Public Property Let p_CaptionHeight(ByVal Propval As Long)
    m_lCaptionImgHeight = Propval
End Property

Public Property Get p_RightEnd() As Long
'/* caption bar right end
    p_RightEnd = m_lRightEnd
End Property

Public Property Let p_RightEnd(ByVal Propval As Long)
    m_lRightEnd = Propval
End Property

Public Property Get p_RightBorderWidth() As Long
'/* right border width
    p_RightBorderWidth = m_lRightBorderWidth
End Property

Public Property Let p_RightBorderWidth(ByVal Propval As Long)
    m_lRightBorderWidth = Propval
End Property

Public Property Get p_SkinForm() As Boolean
'/* form skin switch
    p_SkinForm = m_bSkinForm
End Property

Public Property Let p_SkinForm(ByVal Propval As Boolean)
    m_bSkinForm = Propval
End Property

Public Property Get p_TopSizingBorder() As Long
'/* top sizing border height
    p_TopSizingBorder = m_lTopBorderHeight
End Property

Public Property Let p_TopSizingBorder(ByVal Propval As Long)
    m_lTopBorderHeight = Propval
End Property

Public Property Get p_TransIdx() As Long
'/* main form reference
    p_TransIdx = m_lTransIdx
End Property

Public Property Let p_TransIdx(ByVal Propval As Long)
    m_lTransIdx = Propval
End Property

Public Property Get p_TransBtnIdx() As Long
'/* main form reference
    p_TransBtnIdx = m_TransBtnIdx
End Property

Public Property Let p_TransBtnIdx(ByVal Propval As Long)
    m_TransBtnIdx = Propval
End Property

Public Property Get p_SkinApp() As Boolean
'/* skin app switch
    p_SkinApp = m_bSkinApp
End Property

Public Property Let p_SkinApp(ByVal Propval As Boolean)
    m_bSkinApp = Propval
End Property

Public Property Get p_SkinAppPath() As String
'/* skin app path
    p_SkinAppPath = m_sSkinAppPath
End Property

Public Property Let p_SkinAppPath(ByVal Propval As String)
    m_sSkinAppPath = Propval
End Property

Public Property Get p_ShadowForm() As Boolean
'/* skin app switch
    p_ShadowForm = m_bShadowForm
End Property

Public Property Let p_ShadowForm(ByVal Propval As Boolean)
    m_bShadowForm = Propval
End Property

Public Property Get p_ShadowInset() As Boolean
'/* skin app switch
    p_ShadowInset = m_bShadowInset
End Property

Public Property Let p_ShadowInset(ByVal Propval As Boolean)
    m_bShadowInset = Propval
End Property

Public Property Get p_ShadowDepth() As Long
'/* main form reference
    p_ShadowDepth = m_lShadowDepth
End Property

Public Property Let p_ShadowDepth(ByVal Propval As Long)
    m_lShadowDepth = Propval
End Property

Public Property Get p_ShadowInDepth() As Long
'/* main form reference
    p_ShadowInDepth = m_lShadowInDepth
End Property

Public Property Let p_ShadowInDepth(ByVal Propval As Long)
    m_lShadowInDepth = Propval
End Property

Public Property Get p_ForceUnload() As Boolean
'/* main form reference
    p_ForceUnload = m_bForceUnload
End Property

Public Property Let p_ForceUnload(ByVal Propval As Boolean)
    m_bForceUnload = Propval
End Property

Public Property Get p_CustomCaption() As Boolean
'/* use caption text from form
    p_CustomCaption = m_bCustomCaption
End Property

Public Property Let p_CustomCaption(Propval As Boolean)
    m_bCustomCaption = Propval
End Property

'~ form options ~
Public Property Get p_CustomIcon() As Boolean
'/* use caption icon from form
    p_CustomIcon = m_bCustomIcon
End Property

Public Property Let p_CustomIcon(Propval As Boolean)
    m_bCustomIcon = Propval
End Property

Public Property Get p_CaptionFntClr() As Long
'/* caption font color
    p_CaptionFntClr = m_CaptionFntClr
End Property

Public Property Let p_CaptionFntClr(ByVal Propval As Long)
    m_CaptionFntClr = Propval
End Property

Public Property Get p_MaxSize() As Boolean
'/* use caption icon from form
    p_MaxSize = m_bMaxSize
End Property

Public Property Let p_MaxSize(Propval As Boolean)
    m_bMaxSize = Propval
End Property

Public Property Get p_MaxHeight() As Long
'/* use caption icon from form
    p_MaxHeight = m_lMaxHeight
End Property

Public Property Let p_MaxHeight(Propval As Long)
    m_lMaxHeight = Propval 'm_lMaxWidth
End Property

Public Property Get p_MaxWidth() As Long
'/* use caption icon from form
    p_MaxWidth = m_lMaxWidth
End Property

Public Property Let p_MaxWidth(Propval As Long)
    m_lMaxWidth = Propval
End Property


'////////////////////////////
'//// Private Properties ////
'////////////////////////////

Private Property Get p_CanClose() As Boolean
'/* test if window can close

Dim lStyle  As Long

    lStyle = GetWindowLong(m_lHostHnd, (-16))
    p_CanClose = ((lStyle And &H80000) = &H80000)

End Property

Private Property Get p_CanSize() As Boolean
'/* test if window is sizeable

Dim lStyle  As Long

    lStyle = GetWindowLong(m_lHostHnd, (-16))
    p_CanSize = ((lStyle And &H40000) = &H40000)

End Property

Private Sub Class_Initialize()

    Set m_GSubclass = New MGSubclass
    m_lLeft = GetSystemMetrics(32)
    m_lRight = m_lLeft
    m_lCaptionHeight = GetSystemMetrics(4)
    m_lTop = m_lCaptionHeight + GetSystemMetrics(33)
    m_lBottom = GetSystemMetrics(33)

End Sub


'/////////////////////
'//// Entry Point ////
'/////////////////////

Public Sub Set_Host()
'/* get options and instantiate
'/* attatch the host

Dim tRect   As RECT
Dim lX      As Long
Dim lY      As Long

    '/* get caption text/icon
    If m_lCaptionOffset = 0 Then
        m_lCaptionOffset = 24
    End If
    
    m_lHostHnd = frmHost.hWnd
    m_lMainHnd = m_oFrmMain.hWnd
    Attach m_lHostHnd

    '/* instance object tracking array
    PreInstance_ObjTrack

    '/* initialize tracking variables
    Init_Images
    
    '/* test for base images
    If (p_IBorders Is Nothing) Or (p_ICaption Is Nothing) Or _
    (p_ICBoxMin Is Nothing) Or (p_ICBoxMax Is Nothing) Or _
    (p_ICBoxRst Is Nothing) Or (p_ICBoxCls Is Nothing) Then Exit Sub

    '/* create caption/border/button image mappings
    Store_Picture p_ICaption, 0
    Store_Picture p_IBorders, Increment_Images
    
    '/* add two instances for borders
    Increment_Images
    Increment_Images
    
    '/* load them
    Store_Picture p_ICBoxMin, Increment_Images
    Store_Picture p_ICBoxMax, Increment_Images
    Store_Picture p_ICBoxRst, Increment_Images
    Store_Picture p_ICBoxCls, Increment_Images
    
    '/* image offsets
    m_lTopOffset = m_lCaptionImgHeight + m_lLeftBorderWidth
    m_lBorderOffset = m_lLeftBorderWidth + m_lRightBorderWidth
    
    '/* get button status
    m_bCanClose = p_CanClose
    m_bCanSize = p_CanSize
    
    '/* create image dc
    With frmHost
        m_mainSurfaceDC = CreateCompatibleDC(.hdc)
        m_backSurfaceDC = CreateCompatibleDC(.hdc)
    End With
    
    '/* get caption and icon
    Get_Caption m_lMainHnd
    Get_Icon m_lMainHnd
    
    '/* test app variables
    If m_bSkinApp Then
        If Not Len(m_sSkinAppPath) = 0 Then
            Application_Skin m_sSkinAppPath
        End If
    End If

    '/* swap style bits
    Change_Frame
    
    '/* form shadow
    If m_bShadowForm Or m_bShadowInset Then
        Create_Shadow
    End If
    

End Sub


'///////////////////////////
'//// Subclass Routines ////
'///////////////////////////

Private Sub Attach(ByVal lHwnd As Long)
'/* start subclassing

    '/* host
    Detach lHwnd
    With m_GSubclass
        .Attach_Message Me, lHwnd, &H2
        .Attach_Message Me, lHwnd, &H5
        .Attach_Message Me, lHwnd, &H216
        .Attach_Message Me, lHwnd, &H214
        .Attach_Message Me, lHwnd, &H232&
        .Attach_Message Me, lHwnd, &H84
        .Attach_Message Me, lHwnd, &HA0
        .Attach_Message Me, lHwnd, &HA1
        .Attach_Message Me, lHwnd, &HA2
        .Attach_Message Me, lHwnd, &HA3
        .Attach_Message Me, lHwnd, &H20
        .Attach_Message Me, lHwnd, &H24
    End With
    
End Sub

Private Sub Detach(ByVal lHwnd As Long)
'/* stop subclassing

    With m_GSubclass
        .Detach_Message Me, lHwnd, &H2
        .Detach_Message Me, lHwnd, &H5
        .Detach_Message Me, lHwnd, &H216
        .Detach_Message Me, lHwnd, &H214
        .Detach_Message Me, lHwnd, &H232&
        .Detach_Message Me, lHwnd, &H84
        .Detach_Message Me, lHwnd, &HA0
        .Detach_Message Me, lHwnd, &HA1
        .Detach_Message Me, lHwnd, &HA2
        .Detach_Message Me, lHwnd, &HA3
        .Detach_Message Me, lHwnd, &H20
        .Detach_Message Me, lHwnd, &H24
    End With

End Sub

Private Function MISubclass_WindowProc(ByVal lHwnd As Long, _
                                       ByVal iMsg As Long, _
                                       ByVal wParam As Long, _
                                       ByVal lParam As Long) As Long

'/* subclass message queue

Dim lHitCode        As Long
Static lLastCode    As Long
Dim lStyle          As Long
Dim eHitCode        As eHitConstants
Dim tShadow         As RECT
Dim tRect           As RECT
Dim tRect2          As RECT
Static lCurrX       As Long
Static lCurrY       As Long
Static lCurrR       As Long
Static lCurrB       As Long
Static lSizeHndX    As Long
Static lSizeHndY    As Long
Static bHandled     As Boolean
Static bMoving      As Boolean

    Select Case iMsg
    '/* destroy
    Case &H2
        m_oFrmMain.Hide
        Detach lHwnd
        Class_Unload
        RaiseEvent eTerminate
        
    '/* min size
    Case &H24
        RtlMoveMemory ByVal lParam, m_SizeInfo, LenB(m_SizeInfo)
    
    '/* size
    Case &H5
        If bMoving Then Exit Function
        '/* get current size
        GetWindowRect lHwnd, tRect
        With tRect
            lCurrR = .Right
            lCurrB = .Bottom
            lCurrX = .Left
            lCurrY = .Top
        End With
        '/* control sizing to factor of 2 pixels
        If (Abs(Abs(lCurrR) - Abs(m_lCurrentR)) > 0) Or _
            (Abs(Abs(lCurrB) - Abs(m_lCurrentB)) > 0) Then
            Paint_Window lHwnd
            'Sleep 5
            With tRect
                GetWindowRect lHwnd, tRect
                MoveWindow m_lMainHnd, .Left + m_lLeftBorderWidth, .Top + m_lCaptionImgHeight, _
                (lCurrR - .Left) - m_lBorderOffset, (lCurrB - .Top) - m_lTopOffset, True
            End With
            Sleep 2
            Size_Gate
        End If
       
    '/* mouse over: min 8/max 9/close 20
    Case &HA0
        lHitCode = NCMouse_Over
        If lLastCode = lHitCode Then
            If lLastCode = 0 And bHandled Then
                Paint_Window lHwnd
                bHandled = False
            End If
            Exit Function
        End If
        
        lLastCode = lHitCode
        m_bMouseDown = False
        bHandled = True
        Select Case lHitCode
        Case 8
            Paint_Window lHwnd, lHitCode
        Case 9
            Paint_Window lHwnd, lHitCode
        Case 20
            Paint_Window lHwnd, lHitCode
        Case Else
            Paint_Window lHwnd
            lLastCode = 0
        End Select
        
    '/* nc mouse down
    Case &HA1
        '/* get current coords
        Size_Gate
        '/* turn off shadows for move/resize
        Shadow_State False
        '/* button image state flag
        m_bMouseDown = True
        '/* test for button press
        Select Case lLastCode
        '/* minimize form
        Case 8
            Paint_Window lHwnd, 8
            Sleep 40
            Window_Post SC_MINIMIZE
            
        '/* max/restore
        Case 9
            Paint_Window lHwnd, 9
            Sleep 40
            lStyle = GetWindowLong(m_lHostHnd, (-16))
            If ((lStyle And &H1000000) = &H1000000) Then
                Window_Post SC_RESTORE
            Else
                Window_Post SC_MAXIMIZE
            End If
            
        '/* close
        Case 20
            Paint_Window lHwnd, 20
            Sleep 40
            frmHost.Hide
            m_oFrmMain.Hide
            Window_Post SC_CLOSE
            m_bForceUnload = True

        '/* move event
        Case Else
            m_bMouseDown = False
            '/* consume left and top moves (too flickery)
            If (m_eLastHT = 10) Or (m_eLastHT = 12) Or _
               (m_eLastHT = 13) Or (m_eLastHT = 16) Then
                MISubclass_WindowProc = &H85
            Else
                MISubclass_MsgResponse = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
            End If
        End Select
    
    '/* cursor change
    Case &H20
        If (m_eLastHT = 10) Or (m_eLastHT = 12) Or _
           (m_eLastHT = 13) Or (m_eLastHT = 16) Then
            MISubclass_WindowProc = &H85
        Else
            MISubclass_MsgResponse = m_GSubclass.Get_OldProc(lHwnd, iMsg, wParam, lParam)
        End If
        
    '/* nc mouse up
    Case &HA2
        '/* hide shadow during resize
        m_bMouseDown = False
        If m_bShadowForm Then
            m_ShadowRight.Show
            m_ShadowBottom.Show
        End If

    '/* sizing
    Case &H214
        If bMoving Then Exit Function
        If m_bShadowForm Or m_bShadowInset Then
            Calc_Rect tRect
        End If
    
    '/* move
    Case &H216
        bMoving = True
        '/*** gets rid of flicker, but slow!
        Sleep 100
        CopyMemory tRect, ByVal lParam, Len(tRect)
        With tRect
            SetWindowPos m_lMainHnd, 0, .Left + m_lLeftBorderWidth, _
            .Top + m_lCaptionImgHeight, 0, 0, &H1
        End With

    '/* exit size change
    Case &H232&
        '/* reset shadows
        Shadow_State True
        SetWindowPos m_lMainHnd, 0, 0, 0, 0, 0, &H1 Or &H2 Or &H4 Or &H10 Or &H20
        bMoving = False
        
    '/* hit test
    Case &H84
        '/* reset hitcode
        eHitCode = Get_HitCode()
        m_eLastHT = eHitCode
        MISubclass_WindowProc = eHitCode
        
    '/* caption dbl click
    Case &HA3
        Paint_Window lHwnd
    End Select

End Function

Private Function Transition_Gate(tTrst As RECT) As Long

'/* here is a strange problem..
'/* this is best solution to flicker
'/* on form move. Instead of syncing forms
'/* create an image of frmMain dc, and blt
'/* it into the skin. when mouse
'/* is released from form, main gets
'/* resized w/ top postion and image is hidden..
'/* but here is problem, for some reason,
'/* main dc, or cut from screen dc comes out
'/* monochrome with button face color
'/* invisible.. strange huh?
'/* you can unrem gate in misubclass and
'/* draw_window to see what I mean
'/* if you spot the issue, give me a holler..

Dim lDc         As Long
Dim lNDc        As Long
Dim lBmp        As Long
Dim lBmpOld     As Long
Dim lScrHnd     As Long
Dim lScrDc      As Long
Dim tRect       As RECT
Dim tempBlend   As BLENDFUNCTION

    GetWindowRect m_oFrmMain.hWnd, tRect
    lScrDc = GetDC(0)
    lNDc = CreateCompatibleDC(lScrDc)
    
    lScrHnd = GetDesktopWindow()
    lScrDc = GetDC(lScrHnd)
    BitBlt lNDc, 0&, 0&, Screen.Width, Screen.Height, lScrDc, 0&, 0&, vbSrcCopy
    
    '/* alphablend struct
    With tempBlend
        .AlphaFormat = 0
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    CopyMemory m_backBlendFUnc, tempBlend, 4
    
    With tRect
        lBmp = CreateCompatibleBitmap(lNDc, .Right, .Bottom)
        lBmpOld = SelectObject(lNDc, lBmp)
        BitBlt lNDc, 0, 0, .Right, .Bottom, lScrDc, .Left, .Top, vbSrcCopy
        'AlphaBlend lNDc, 0, 0, .Right, .Bottom, lScrDc, .Left, .Top, .Right, .Bottom, m_backBlendFUnc
    End With
    
    With tTrst
        BitBlt m_AlphaTmpDc, .Left, .Top, .Right, .Bottom, lNDc, 0, 0, vbSrcCopy
    End With
    
    m_oFrmMain.Hide
    DeleteObject lBmpOld
    DeleteObject lBmp
    DeleteDC lDc

End Function

Private Property Get MISubclass_MsgResponse() As EMsgResponse
'/* message status property

    Select Case m_GSubclass.p_CurrentMessage
    Case &HA1, &H5, &H20
        MISubclass_MsgResponse = emrConsume
    Case &H84
        MISubclass_MsgResponse = emrPostProcess
    Case Else
        MISubclass_MsgResponse = emrPreProcess
    End Select

End Property

Private Property Let MISubclass_MsgResponse(ByVal RHS As EMsgResponse)

    '<STUB>

End Property


'///////////////////////////
'//// Instance Tracking ////
'///////////////////////////

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

On Error Resume Next

    For lCount = 1 To UBound(vObjArray)
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
    Next lCount

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

Private Function Return_OldBmp(ByVal sInstName As String) As Long
'/* fetch old bmp

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If lIndex = 0 Then Exit Function
    Return_OldBmp = vObjArray(1, lIndex)
    
On Error GoTo 0

End Function

Private Function Return_NewBmp(ByVal sInstName As String) As Long
'/* fetch new bmp

Dim lIndex As Long

On Error Resume Next

    lIndex = c_ObjTrack.Item(sInstName)
    If lIndex = 0 Then Exit Function
    Return_NewBmp = vObjArray(2, lIndex)
    
On Error GoTo 0

End Function


'//////////////////////////
'//// Sizing Routines  ////
'//////////////////////////

Private Sub Change_Frame()

'/* swap style bits between
'/* host and main forms

Dim lReturn     As Long
Dim tRect       As RECT

    '/* get base styles
    '/* host: extended styles + layered
    lReturn = GetWindowLong(m_lHostHnd, GWL_EXSTYLE)
    SetWindowLong m_lHostHnd, GWL_EXSTYLE, lReturn Or WS_EX_LAYERED
    m_bInSizeMove = True
    '/* main: borderless
    SetWindowLong m_lMainHnd, GWL_STYLE, &H10000000
    m_oFrmMain.Hide
    m_bInSizeMove = False
    '/* paint host form
    With m_oFrmMain
        MoveWindow m_lHostHnd, .Left, .Top, .Width / Screen.TwipsPerPixelX, _
        .Height / Screen.TwipsPerPixelY, True
        .Show
    End With
    frmHost.Show

End Sub

Private Sub Reset_Frame(Optional bReset As Boolean)

'/* swap style bits between
'/* host and main forms

Dim lReturn     As Long

    '/* get base styles
    '/* host: extended styles + layered
    lReturn = GetWindowLong(m_lHostHnd, GWL_STYLE)
    SetWindowLong m_lMainHnd, GWL_STYLE, lReturn
    m_oFrmMain.Hide
    With frmHost
        MoveWindow m_lMainHnd, .Left, .Top, .Width / Screen.TwipsPerPixelX, _
        .Height / Screen.TwipsPerPixelY, False
    End With
    frmHost.Hide

End Sub

Private Function Size_Gate() As Boolean
'/* sizing minimum check

Dim tRect As RECT
    
    GetWindowRect m_lHostHnd, tRect
    m_lCurrentX = tRect.Left
    m_lCurrentY = tRect.Top
    m_lCurrentR = tRect.Right
    m_lCurrentB = tRect.Bottom
    
End Function

Private Function Calc_Rect(ByRef tRect As RECT) As Long
'/* resize hosted form

Dim lBottom     As Long
Dim lTop        As Long
Dim lLeft       As Long
Dim lRight      As Long

    '/* sync hosted form
    With tRect
        If m_bShadowForm Then
            m_ShadowRight.Render_Shadow .Left, .Right, .Top, .Bottom, True
            m_ShadowBottom.Render_Shadow .Left, .Right, .Top, .Bottom, True
        End If
        '/* sync interior shadow
        If m_bShadowInset Then
            m_ShadowInRight.Render_Shadow .Left, .Right - m_lRightBorderWidth, _
            .Top + m_lLeftBorderWidth, .Bottom - m_lRightBorderWidth, True
            m_ShadowInBottom.Render_Shadow .Left, .Right - m_lLeftBorderWidth, _
            .Top, .Bottom - m_lLeftBorderWidth, True
        End If
    End With
    
End Function


'//////////////////////////////
'//// Application Skinning ////
'//////////////////////////////

Private Sub Application_Skin(ByVal sAppName As String)

'/* launch app as child in main form
'/* and reset style

Dim lPid    As Long

    LockWindowUpdate GetDesktopWindow
    lPid = Shell(sAppName, vbNormalFocus)
    If lPid = 0 Then Exit Sub
    m_lAppHwnd = Application_Instance(lPid)
    SetParent m_lAppHwnd, m_lMainHnd
    SetWindowLong m_lAppHwnd, GWL_STYLE, &H10000000
    SetWindowLong m_lAppHwnd, GWL_EXSTYLE, &H40
    SetWindowPos m_lAppHwnd, 0, 0, 0, 0, 0, &H1 Or &H20 Or &H10 Or &H4
    Putfocus m_lAppHwnd
    LockWindowUpdate False

End Sub

Private Function Application_Instance(ByVal lTarget As Long) As Long
'/* if you want to hook apps on the fly,
'/* this is not the way to do it.. there
'/* are many ways.. easy way is to run timer that
'/* enumerates windows (bad), better way (but hard in vb)
'/* is to hook shell, and store skinned app refs in registry.
'/* hooks in shell address table forward app launch to
'/* dll, which either owns app or forwards call..
'/* either way, rcm classes would have to reside in dll

Dim lHwnd As Long
Dim lPid  As Long

    lHwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While lHwnd <> 0
        If GetParent(lHwnd) = 0 Then
            GetWindowThreadProcessId lHwnd, lPid
            If lPid = lTarget Then
                Application_Instance = lHwnd
                Exit Do
            End If
        End If
        lHwnd = GetWindow(lHwnd, &H2)
    Loop

End Function

Private Sub App_Unload()
'/* destroy app instance

    DestroyWindow m_lAppHwnd

End Sub


'/////////////////////////
'//// Shadow Creation ////
'/////////////////////////

Public Sub Create_Shadow()
'/* instantiate form shadow

On Error Resume Next

    '/* test defaults
    If m_lShadowDepth < 2 Then
        m_lShadowDepth = 2
    ElseIf m_lShadowDepth > 15 Then
        m_lShadowDepth = 15
    End If
    
    If (m_lShadowInDepth < 2) Or _
    (m_lShadowInDepth > 15) Then
        m_lShadowInDepth = m_lShadowDepth
    End If
    
    '/* form shadow
    
    '/* right shadow
    If m_bShadowForm Then
        Set m_ShadowRight = New frmShadow
        With m_ShadowRight
            .p_ShadowSize = m_lShadowDepth
            .p_ShadowType = RightShadow
            .p_Shadow = frmHost
            .Show , frmHost
        End With
        
        '/* bottom shadow
        Set m_ShadowBottom = New frmShadow
        With m_ShadowBottom
            .p_ShadowSize = m_lShadowDepth
            .p_ShadowType = BottomShadow
            .p_Shadow = frmHost
            .Show , frmHost
        End With
    End If
    
    '/* inner shadow
    
    '/* right shadow
    If m_bShadowInset Then
        Set m_ShadowInRight = New frmShadow
        With m_ShadowInRight
            .p_ShadowSize = m_lShadowInDepth
            .p_ShadowType = RightShadow
            .p_Shadow = m_oFrmMain
            .Show , m_oFrmMain
        End With
        
        '/* bottom shadow
        Set m_ShadowInBottom = New frmShadow
        With m_ShadowInBottom
            .p_ShadowSize = m_lShadowInDepth
            .p_ShadowType = BottomShadow
            .p_Shadow = m_oFrmMain
            .Show , m_oFrmMain
        End With
    End If
        
On Error GoTo 0

End Sub

Private Sub Shadow_State(ByVal bShowState As Boolean)
'/* toggle shadow show state

Dim tRect As RECT

    If bShowState Then
        GetWindowRect m_lHostHnd, tRect
        Calc_Rect tRect
        If m_bShadowInset Then
            m_ShadowInRight.Show
            m_ShadowInBottom.Show
        End If

        If m_bShadowForm Then
            m_ShadowRight.Show
            m_ShadowBottom.Show
        End If
    Else
        If m_bShadowInset Then
            m_ShadowInRight.Hide
            m_ShadowInBottom.Hide
        End If

        If m_bShadowForm Then
            m_ShadowRight.Hide
            m_ShadowBottom.Hide
        End If
    End If
    
End Sub

'///////////////////////////
'//// Caption Support  ////
'///////////////////////////

Private Sub Get_Caption(lHwnd As Long)
'/* get caption text from hosted form

Dim tPnt        As POINTAPI

    m_CapText = m_oFrmMain.Caption
    If Len(m_CapText) = 0 Then
        m_CapText = "NeoClass V6"
        tPnt.X = 210
    Else
        '/* get text height/width
        GetTextExtentPoint32 m_oFrmMain.hdc, m_CapText, Len(m_CapText) + 2, tPnt
    End If

    '/* borrowed from Paul Catons example
    '/* set minimum width
    '/* via caption/button size
    With m_SizeInfo
        '/* minimum size
        With .ptMinTrackSize
            .X = (((m_lButtonWidth + 1) * 3) - m_lControlOffsetX) + tPnt.X + (m_lCaptionOffset + 32)
            .Y = 25
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
        If Not m_bMaxSize Then
            With .ptMaxTrackSize
                .X = Screen.Width
                .Y = Screen.Height
            End With
        Else
            With .ptMaxTrackSize
                .X = m_lMaxWidth
                .Y = m_lMaxHeight
            End With
        End If
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
    End With
    DrawText lhDC, m_CapText, -1, tRect, &H1 Or &H20 Or &H4
    
    '/* delete fonts
    DeleteObject lPFont
    DeleteObject lMFont

End Sub

Private Sub Get_Icon(ByVal lHwnd As Long)
'/* get form icon handle

    m_lCapIcon = frmHost.Icon.handle
    Make_Icon

End Sub

Private Sub Set_Icon(ByVal lhDC As Long)
'/* draw icon

Dim lTop        As Long
Dim lLeft       As Long

    '/* paint to dc
    lTop = (m_lCaptionHeight) / 4
    lLeft = (m_lLeftBorderWidth + 4)
    TransparentBlt lhDC, lLeft, lTop, 16, 16, m_lIconDc, 0, 0, 16, 16, vbMagenta

End Sub

Private Sub Make_Icon()

Dim lSDc        As Long
Dim lBmp        As Long
Dim tRect       As RECT
Dim lLBrsh      As LOGBRUSH
Dim lTop        As Long
Dim lLeft       As Long

    '/* create a temp image dc
    m_lIconDc = CreateCompatibleDC(frmHost.hdc)
    lBmp = CreateCompatibleBitmap(frmHost.hdc, 16, 16)
    SelectObject m_lIconDc, lBmp
    DeleteObject lBmp
    tRect.Right = 16
    tRect.Bottom = 16
    
    '/* backpaint
    lLBrsh.lbColor = vbMagenta
    lSDc = CreateBrushIndirect(lLBrsh)
    FillRect m_lIconDc, tRect, lSDc
    DeleteObject lSDc
    DrawIconEx m_lIconDc, 0, 0, m_lCapIcon, 16, 16, 0, 0, DI_NORMAL
    DeleteObject lBmp
    
End Sub


'////////////////////////
'//// Form Paintning ////
'////////////////////////

Private Sub Paint_Window(ByVal lHwnd As Long, _
                         Optional lHitState As Long)

'/* window painting callout

Dim tRect   As RECT
Dim lhDC    As Long

    lhDC = GetWindowDC(lHwnd)
    GetWindowRect lHwnd, tRect
    With tRect
        OffsetRect tRect, -.Left, -.Top
        Draw_Window lhDC, lHwnd, .Left, .Top, .Right, .Bottom, lHitState
    End With
    ReleaseDC lHwnd, lhDC

End Sub

Private Sub Window_Post(ByVal eCmd As ECNCSysCommandConstants)
'/* post change to message queue

    PostMessage m_lHostHnd, &H112, eCmd, 0

End Sub

Private Sub Refresh_Windows()
    
    SetWindowPos m_lHostHnd, 0, 0, 0, 0, 0, &H1 Or &H2 Or &H20 Or &H10 Or &H4
    SetWindowPos m_lMainHnd, 0, 0, 0, 0, 0, &H1 Or &H2 Or &H20 Or &H10 Or &H4

End Sub

Private Sub Refresh_Controls()
'/* refresh controls on main form

Dim oCtrl   As Control

On Error Resume Next

    For Each oCtrl In m_oFrmMain
        If TypeOf oCtrl Is PictureBox Then
            oCtrl.Refresh
        ElseIf TypeOf oCtrl Is Frame Then
            oCtrl.Refresh
        ElseIf TypeOf oCtrl Is TextBox Then
            oCtrl.Refresh
        End If
    Next oCtrl

On Error GoTo 0

End Sub

Private Sub Draw_Window(ByVal lDstHdc As Long, _
                        ByVal lHwnd As Long, _
                        ByVal lLeft As Long, _
                        ByVal lTop As Long, _
                        ByVal lRight As Long, _
                        ByVal lBottom As Long, _
                        Optional ByVal lHitCode As Long)

'/* draw borders and frame into temp dc
'/* uses bitblt for fast rendering into
'/* back buffer. completed image sent
'/* to transparency subs for re-rendering

Dim bWindowActive       As Boolean
Dim lHeight             As Long
Dim lRightWidth         As Long
Dim lWidth              As Long
Dim lYpos               As Long
Dim lCap                As Long
Dim lCpEnd              As Long
Dim lBorderSize         As Long
Dim lSrcY               As Long
Dim lOrgX               As Long
Dim lStyle              As Long
Dim lhDC                As Long
Dim lBtnLeft            As Long
Dim eState              As ECNCButtonStates
Dim tPos                As POINTAPI
Dim tRect               As RECT
Dim tTrst               As RECT
Dim tButtons            As RECT
Dim tButtonR            As RECT

On Error Resume Next

    '/* create back buffer
    Alpha_Buffer lRight, lBottom, True
    
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
    GetWindowRect m_lHostHnd, tRect
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
        LSet m_tBtn(0) = tButtonR
        eState = Up
        If m_bMouseDown And lHitCode = 20 Then
            eState = Down
        ElseIf lHitCode = 20 Then
            eState = Over
        Else
            eState = Up
        End If
        NCButton_Draw lhDC, 0, eState, tButtonR
    End If

    '/* get max button size/state
    If m_bMaximise Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        With tButtonR
            .Left = lBtnLeft
            .Right = .Left + m_lButtonWidth + 1
        End With
        LSet m_tBtn(1) = tButtonR
        eState = Up
        If m_bMouseDown And lHitCode = 9 Then
            eState = Down
        ElseIf lHitCode = 9 Then
            eState = Over
        Else
            eState = Up
        End If
        NCButton_Draw lhDC, 1, eState, tButtonR
    End If

    '/* get min button size/state
    If m_bMinimise Then
        LSet tButtonR = tButtons
        lBtnLeft = lBtnLeft - (m_lButtonWidth + 1)
        With tButtonR
            .Left = lBtnLeft
            .Right = .Left + m_lButtonWidth + 1
        End With
        LSet m_tBtn(2) = tButtonR
        eState = Up
        If m_bMouseDown And lHitCode = 8 Then
            eState = Down
        ElseIf lHitCode = 8 Then
            eState = Over
        Else
            eState = Up
        End If
        NCButton_Draw lhDC, 2, eState, tButtonR
    End If

    '/* copy caption bar
    BitBlt m_AlphaTmpDc, lLeft, lTop, m_lWd(2), m_lHt(2), lhDC, 0, 0, vbSrcCopy
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
    
    '/* draw left and right to screen
    BitBlt m_AlphaTmpDc, lLeft, lTop + m_lHt(0), m_lLeftBorderWidth, _
    lBottom - lTop - m_lHt(0), m_lDc(3), 0, 0, vbSrcCopy
    
    BitBlt m_AlphaTmpDc, lRight - m_lRightBorderWidth - 1, lTop + m_lHt(0), _
    m_lRightBorderWidth + 1, lBottom - lTop - m_lHt(0), m_lDc(3), _
    lBorderSize + (lBorderSize - m_lRightBorderWidth - 1), 0, vbSrcCopy
    
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
    
    '/* create bottom image
    BitBlt m_lDc(2), 0, 0, lBorderSize, lBorderSize, m_lDc(1), _
    lBorderSize * 2, lSrcY, vbSrcCopy
    
    If m_bCanSize Then
        BitBlt m_lDc(2), lRight - lLeft - lBorderSize, 0, _
        lBorderSize, lBorderSize, m_lDc(1), lBorderSize * 6, lSrcY, vbSrcCopy
    Else
        BitBlt m_lDc(2), lRight - lLeft - lBorderSize, 0, _
        lBorderSize, lBorderSize, m_lDc(1), lBorderSize * 5, lSrcY, vbSrcCopy
    End If
    
    '/* draw caption text
    'Set_Caption m_AlphaTmpDc
    
    '/* draw bottom to screen
    BitBlt m_AlphaTmpDc, lLeft, lBottom - lBorderSize, _
    lRight - lLeft + 1, lBorderSize, m_lDc(2), 0, 0, vbSrcCopy

    '/* send to trans
    Set_Transparent lDstHdc, lLeft, lTop, lRight, lBottom, m_AlphaTmpDc, _
    tButtons.Left, tButtons.Top, (m_lButtonWidth + 1) * 3, m_lButtonHeight
    
    '/* flush backbuffers
    Alpha_Buffer 0, 0, False
    
On Error GoTo 0

End Sub

Private Sub Alpha_Buffer(ByVal lDstWidth As Long, _
                         ByVal lDstHgt As Long, _
                         ByVal bCreate As Boolean)

'/* create|destroy alpha backbuffer

Dim lhDC    As Long
Dim lBmpOld As Long

    If bCreate Then
        lhDC = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
        m_AlphaTmpDc = CreateCompatibleDC(lhDC)
        m_AlphaBmp = CreateCompatibleBitmap(lhDC, lDstWidth, lDstHgt)
        lBmpOld = SelectObject(m_AlphaTmpDc, m_AlphaBmp)
        DeleteObject lBmpOld
        DeleteDC lhDC
    Else
        If Not m_AlphaBmp = 0 Then
            DeleteObject m_AlphaBmp
            DeleteDC m_AlphaTmpDc
            m_AlphaTmpDc = 0
        End If
    End If

End Sub

Private Sub Set_Transparent(ByVal lDestDc As Long, _
                            ByVal lDestX As Long, _
                            ByVal lDestY As Long, _
                            ByVal lDstWidth As Long, _
                            ByVal lDstHeight As Long, _
                            ByVal lSrcDc As Long, _
                            ByVal lCapX As Long, _
                            ByVal lCapY As Long, _
                            ByVal lCapWidth As Long, _
                            ByVal lCapHeight As Long)

'/* create DIB and blend to form

Dim tempBI      As BITMAPINFO
Dim tempBlend   As BLENDFUNCTION
Dim lhDC        As Long
Dim tRect       As RECT
    
    '/* 32b DIB structure
    With tempBI.bmiHeader
        .biSize = Len(tempBI.bmiHeader)
        .biBitCount = 32
        .biHeight = lDstHeight
        .biWidth = lDstWidth
        .biPlanes = 1
        .biSizeImage = .biWidth * .biHeight * (.biBitCount / 8)
    End With
    
    '/* test defaults
    If m_lTransIdx = 0 Then m_lTransIdx = 225
    If m_TransBtnIdx = 0 Then m_TransBtnIdx = m_lTransIdx
    
    '/* remove old image
    If Not m_mainSurfaceBitmap = 0 Then
        SelectObject m_mainSurfaceDC, m_mainSurfaceBitmap
        DeleteObject m_mainSurfaceBitmap
        m_mainSurfaceBitmap = 0
    End If
    
    '/* remove old dc
    If Not m_mainSurfaceDC = 0 Then
        DeleteDC m_mainSurfaceDC
        m_mainSurfaceDC = 0
        m_mainSurfaceDC = CreateCompatibleDC(frmHost.hdc)
    End If
    
    '/* copy structure
    m_mainSurfaceBitmap = CreateDIBSection(m_mainSurfaceDC, tempBI, &H0, ByVal 0, 0, 0)
    SelectObject m_mainSurfaceDC, m_mainSurfaceBitmap
    
    '/* add to tracking array
    ObjTrack_Add "mainpic", m_mainSurfaceBitmap, NewBmp
    ObjTrack_Add "maindc", m_mainSurfaceDC, ObjDc
    
    '/* alphablend struct
    With tempBlend
        .AlphaFormat = 0
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = m_lTransIdx
    End With

    CopyMemory m_backBlendFUnc, tempBlend, 4
    AlphaBlend m_mainSurfaceDC, lDestX, lDestY, lDstWidth, _
    lDstHeight, lSrcDc, 0, 0, lDstWidth, lDstHeight, m_backBlendFUnc
    DeleteDC lSrcDc
    
    '/* window size struct
    m_srcPoint.X = lDestX
    m_srcPoint.Y = lDestY
    m_windowSize.cx = lDstWidth
    m_windowSize.cy = lDstHeight
    
    '/* ULW struct
    With m_blendFunc
        .AlphaFormat = AC_SRC_ALPHA
        .BlendFlags = 0
        .BlendOp = AC_SRC_OVER
        .SourceConstantAlpha = 255
    End With
    
    '/* draw icon/text onto temp dc
    If m_bCustomIcon Then Set_Icon m_mainSurfaceDC
    If m_bCustomCaption Then Set_Caption m_mainSurfaceDC
    
    '/* copy DIB into form
    UpdateLayeredWindow m_lHostHnd, frmHost.hdc, ByVal 0, _
    m_windowSize, m_mainSurfaceDC, m_srcPoint, vbMagenta, m_blendFunc, ULW_ALPHA
    
End Sub

Private Function NCMouse_Over() As Long
'/* caption button image swap

Dim lhDC            As Long
Static sPrevStat    As Long
Dim tPos            As POINTAPI
Dim tRect           As RECT

    '/* get rect
    GetCursorPos tPos
    GetWindowRect m_lHostHnd, tRect
    lhDC = GetWindowDC(m_lHostHnd)
    
    '/* get cursor pos and offset
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
        If Not PtInRect(m_tBtn(0), .X, .Y) = 0 Then
            If Not sPrevStat = 1 Then NCMouse_Over = 20
        ElseIf Not PtInRect(m_tBtn(1), .X, .Y) = 0 Then
            If Not sPrevStat = 2 Then NCMouse_Over = 9
        ElseIf Not PtInRect(m_tBtn(2), .X, .Y) = 0 Then
            If Not sPrevStat = 3 Then NCMouse_Over = 8
        End If
    End With
    
    ReleaseDC m_lHostHnd, lhDC

End Function

Private Function Get_HitCode() As eHitConstants
'/* retrieve event type from current position

Dim tPos    As POINTAPI
Dim tRect   As RECT

    '/* get rect
    GetCursorPos tPos
    GetWindowRect m_lHostHnd, tRect
    With tPos
        .X = .X - tRect.Left
        .Y = .Y - tRect.Top
    End With
    OffsetRect tRect, -tRect.Left, -tRect.Top
    
    '/* default code
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
            ElseIf .Y <= m_lBottom Then
                If m_bCanSize Then Get_HitCode = HTTOP
            ElseIf .Y >= tRect.Bottom - m_lBottom Then
                If m_bCanSize Then Get_HitCode = HTBOTTOM
            
            '/* caption
            ElseIf .Y <= m_lTop Then
                If .Y <= m_lBottom + m_lCaptionHeight Then
                    Get_HitCode = HTCAPTION
                    If .X <= m_lCaptionHeight Then Get_HitCode = HTSYSMENU
                ElseIf .Y > m_lBottom + m_lCaptionHeight Then
                    Get_HitCode = HTCLIENT
                End If
            End If
        End If
    End With

End Function

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
            lStyle = GetWindowLong(m_lHostHnd, (-16))
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

    '/* get pos index
    Select Case lIndex
    Case 1
        lSrcWidth = lSrcWidth + m_lButtonWidth
    Case 2
        lSrcWidth = lSrcWidth + (m_lButtonWidth * 2)
    End Select
    
    '/* draw image
    With tRect
        BitBlt lDstDc, .Left, .Top, lDstWidth, _
        lDstHeight, lSrcDc, lSrcWidth, lSrcHeight, vbSrcCopy
    End With

End Sub

Private Sub Init_Images()
'/* initialize image dimension arrays

    ReDim m_lDc(0 To 0) As Long
    ReDim m_lWd(0 To 0) As Long
    ReDim m_lHt(0 To 0) As Long

End Sub

Private Function Increment_Images() As Long
'/* add to image array

On Error Resume Next

    ReDim Preserve m_lDc(0 To UBound(m_lDc) + 1) As Long
    ReDim Preserve m_lWd(0 To UBound(m_lWd) + 1) As Long
    ReDim Preserve m_lHt(0 To UBound(m_lHt) + 1) As Long
    Increment_Images = UBound(m_lDc)
    
On Error GoTo 0

End Function

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

Private Sub Store_Picture(sPic As IPicture, _
                          ByVal lInst As Long)

'/* copy image to workspace

Dim tB          As BITMAP
Dim lTmpDc      As Long
Dim lNewDc      As Long
Dim lBmpOld     As Long

    '/* get pic size
    GetObjectAPI sPic.handle, Len(tB), tB
    Get_Width lInst, tB.bmWidth
    Get_Height lInst, tB.bmHeight
    
    '/* add to temp dc
    lTmpDc = CreateDCAsNull("Display", ByVal 0&, ByVal 0&, ByVal 0&)
    lNewDc = CreateCompatibleDC(lTmpDc)
    lBmpOld = SelectObject(lNewDc, sPic.handle)
    
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
    ObjTrack_Add "framepic" & lInst, lBmp, NewBmp
    'Debug.Print "added " & lBmp
    
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

    lBmpOld = Return_NewBmp("framepic" & lInst)
    If Not lBmpOld = 0 Then
        SelectObject m_lDc(lInst), lBmpOld
        DeleteObject lBmpOld
        'Debug.Print "removed " & lBmpOld
    End If
    
    m_lWd(lInst) = 0
    m_lHt(lInst) = 0
    
    If Not m_lDc(lInst) = 0 Then
        DeleteDC m_lDc(lInst)
        m_lDc(lInst) = 0
    End If

End Sub

Private Sub Unload_Shadow()
'/* unload shadow forms

    If m_bShadowForm Then
        Unload m_ShadowRight
        Set m_ShadowRight = Nothing
        Unload m_ShadowBottom
        Set m_ShadowBottom = Nothing
    End If
    
    If m_bShadowInset Then
        Unload m_ShadowInRight
        Set m_ShadowInRight = Nothing
        Unload m_ShadowInBottom
        Set m_ShadowInBottom = Nothing
    End If

End Sub

Private Sub Object_Cleanup()
'/* cleanup resources

    '/* release resources
    ObjTrack_Destroy
    DestroyIcon m_lCapIcon
    '/* unload shadows
    Unload_Shadow
    '/* destroy app instance
    If m_bSkinApp Then App_Unload

End Sub

Public Sub Unload_Host()

On Error Resume Next

    Reset_Frame
    Unload frmHost
    Set frmHost = Nothing
    
On Error GoTo 0

End Sub

Public Sub Class_Unload()
'/* deallocation and cleanup

On Error Resume Next

    Object_Cleanup
    '/* reset hosted form
    Unload_Host

On Error GoTo 0

End Sub

Private Sub Class_Terminate()
'/* unload me

On Error Resume Next

    If Not m_GSubclass Is Nothing Then
        Set m_GSubclass = Nothing
    End If

On Error GoTo 0

End Sub



'                                                  FIN
'                           **************************************************
'                                   **********************************
'                                           ******************
'                                                   **




'/~ Playing on the stereo as I zip this, (had it in my head all day)..

'/~ AC/DC - Thunderstruck ~/

'/~ I was caught
'/~ In the middle of a railroad track (Thunder)
'/~ I looked round
'/~ And I knew there was no turning back (Thunder)
'/~ My mind raced
'/~ And I thought, what could I do? (Thunder)
'/~ And I knew
'/~ There was no help, no help from you! (Thunder)
'/~ Sound of the drums
'/~ Beatin in my heart
'/~ The thunder of guns
'/~ Tore me apart
'/~ You've been.. Thunderstruck



'                          STEPPENWOLFE
'        STEPPENWOLFE  STEPPENWOLFESTEPPENWOLFEST
'        STEPPENWOLFESTEPPENWOLFE    STEPPENWOLFESTEPPE
'        STEPPENW      STEPPENWOLFESTEPPE    STEPPENWOLFESTEPPENW
'          STEPPE            NWOLFESTEPPENWOLFE  STEP     PENWOLFEST
'          STEPPE                NWOLFESTEPPENWOL                  FEST
'            STEPPE                ST  EPPENWOLFE              STEP  ENWO
'            STEPPE                STEPPENWOLFE              STEPPENWOLFEST
'            STEPPE                  STEP  PENW            STEPPENWOLFESTEPPENW
'            STEPPE                    NWOLFE            STEPPENWOLFESTEPPENWOLFESTEPPENWOLFESTE
'            STEPPE    STEPPENWOLFESTEPPE         STEPPENWOLFE                            STEPPENWO
'            STEPPENWOLFESTEP                    PENWOLFEST                                      EP
'            STEPPENWOLFESTEPPENWOLFE        STEPPENWOLFE                                      STEP
'            STEPPENWOL    ESTEPPENWOLFESTEPPENWOLF                    ESTE                    PPE
'            STEP      STEPPENW    OLFESTEPPENWOLFEST                EPPENWOLFEST           EPPEN
'            STEP      PENWOL                  FESTEPPE          NWOLFESTEPPENWOLFE    STEPPEN
'          STEPPENWOLFESTEPPENWOLFESTEPPENWOLFEST                EPPENWOLFESTEPPENWOLFEST
'          STEPPENWOLFESTEP  PENWOLFWS  TEPPEN                WOLFESTEPP            ENW
'          STEPPENWOL    FESTEP        STEPPENWOLFE        STEPPENWOLFEST          EP
'          STEP      PENWOL      FESTEPPENWOLFESTEPPENWOLFEST      EPPENWOLFE    STEP
'        STEPPE    NWOL      FESTEPPENWOL                FE          STEPPENWOLFEST
'        STEP  PENWOL    FESTEPPEN                    WO

