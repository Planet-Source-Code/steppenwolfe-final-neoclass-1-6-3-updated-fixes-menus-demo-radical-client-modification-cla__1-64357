VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RCMGate"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const VER_PLATFORM_WIN32s        As Integer = 0
Private Const PLATFORM_WIN32_WINDOWS     As Integer = 1
Private Const PLATFORM_WIN32_NT          As Integer = 2

'/* version structure
Private Type OSVersion
    dwOSVersionInfoSize                      As Long
    dwMajorVersion                           As Long
    dwMinorVersion                           As Long
    dwBuildNumber                            As Long
    dwPlatformId                             As Long
    szCSDVersion                             As String * 128
End Type

Private Type BITMAP
    bmType                                   As Long
    bmWidth                                  As Long
    bmHeight                                 As Long
    bmWidthBytes                             As Long
    bmPlanes                                 As Integer
    bmBitsPixel                              As Integer
    bmBits                                   As Long
End Type


Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As OSVersion) As Boolean

Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, _
                                                                      ByVal nCount As Long, _
                                                                      lpObject As Any) As Long

Public Event eGTerminate()

Private m_pCaption                       As StdPicture
Private m_pBorders                       As StdPicture
Private m_pCBoxMin                       As StdPicture
Private m_pCBoxMax                       As StdPicture
Private m_pCBoxRst                       As StdPicture
Private m_pCBoxCls                       As StdPicture
Private m_TransEnabled                   As Boolean
Private m_bCustomCaption                 As Boolean
Private m_CaptionFntClr                  As Long
Private m_bCustomIcon                    As Boolean
Private m_lCaptionOffset                 As Long
Private m_lButtonWidth                   As Long
Private m_lButtonHeight                  As Long
Private m_bCustomButtonPosition          As Boolean
Private m_lControlOffsetX                As Long
Private m_lControlOffsetY                As Long
Private m_lCaptionImgHeight              As Long
Private m_lLeftEnd                       As Long
Private m_lRightStart                    As Long
Private m_lRightEnd                      As Long
Private m_lOffset                        As Long
Private m_lLeftBorderWidth               As Long
Private m_lRightBorderWidth              As Long
Private m_lBottomBorderHeight            As Long
Private m_lTopBorderHeight               As Long
Private m_lTransIdx                      As Long
Private m_bShadowForm                    As Boolean
Private m_bShadowInset                   As Boolean
Private m_lShadowInDepth                 As Long
Private m_lShadowDepth                   As Long
Private m_bBorderHasInactive             As Long
Private m_oFrmMain                       As Form
Private m_bMaxSize                       As Boolean
Private m_lMaxHeight                     As Long
Private m_lMaxWidth                      As Long
Private WithEvents m_RCM                 As RCMSuperClass
Attribute m_RCM.VB_VarHelpID = -1
Private m_Neo                            As cNeoClass


Public Property Get p_oFrmMain() As Object
    Set p_oFrmMain = m_oFrmMain
End Property

Public Property Set p_oFrmMain(ByVal Propval As Object)
    Set m_oFrmMain = Propval
End Property

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

Public Property Let p_CaptionOffset(ByVal Propval As Long)
    m_lCaptionOffset = Propval
End Property

Public Property Get p_CaptionHeight() As Long
'/* left border width
    p_CaptionHeight = m_lCaptionImgHeight
End Property

Public Property Let p_CaptionHeight(ByVal Propval As Long)
    m_lCaptionImgHeight = Propval
End Property

Public Property Get p_CustomCaption() As Boolean
'/* use caption text from form
    p_CustomCaption = m_bCustomCaption
End Property

Public Property Let p_CustomCaption(ByVal Propval As Boolean)
    m_bCustomCaption = Propval
End Property

Public Property Get p_CaptionFntClr() As Long
'/* caption font color
    p_CaptionFntClr = m_CaptionFntClr
End Property

Public Property Let p_CaptionFntClr(ByVal Propval As Long)
    m_CaptionFntClr = Propval
End Property

Public Property Get p_CustomIcon() As Boolean
'/* use caption icon from form
    p_CustomIcon = m_bCustomIcon
End Property

Public Property Let p_CustomIcon(ByVal Propval As Boolean)
    m_bCustomIcon = Propval
End Property

Public Property Get p_Offset() As Long
'/* title offset
    p_Offset = m_lOffset
End Property

Public Property Let p_Offset(ByVal Propval As Long)
    m_lOffset = Propval
End Property

Public Property Get p_RightBorderWidth() As Long
'/* right border width
    p_RightBorderWidth = m_lRightBorderWidth
End Property

Public Property Let p_RightBorderWidth(ByVal Propval As Long)
    m_lRightBorderWidth = Propval
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

Public Property Get p_TransForm() As Boolean
'/* skin app switch
    p_TransForm = m_TransEnabled
End Property

Public Property Let p_TransForm(ByVal Propval As Boolean)
    m_TransEnabled = Propval
End Property

Public Property Get p_RightEnd() As Long
'/* caption bar right end
    p_RightEnd = m_lRightEnd
End Property

Public Property Let p_RightEnd(ByVal Propval As Long)
    m_lRightEnd = Propval
End Property

Public Property Get p_MaxSize() As Boolean
'/* use caption icon from form
    p_MaxSize = m_bMaxSize
End Property

Public Property Let p_MaxSize(ByVal Propval As Boolean)
    m_bMaxSize = Propval
End Property

Public Property Get p_MaxHeight() As Long
'/* use caption icon from form
    p_MaxHeight = m_lMaxHeight
End Property

Public Property Let p_MaxHeight(ByVal Propval As Long)
    m_lMaxHeight = Propval
End Property

Public Property Get p_MaxWidth() As Long
'/* use caption icon from form
    p_MaxWidth = m_lMaxWidth
End Property

Public Property Let p_MaxWidth(ByVal Propval As Long)
    m_lMaxWidth = Propval
End Property

Public Sub Start()

Dim tBmp            As BITMAP
Dim lBtnWidth       As Long
Dim lBtnHeight      As Long
Dim lCapWidth       As Long
Dim lCapHeight      As Long
Dim lBorderWdth     As Long

    '/* get global image metrics
    '/* caption buttons
    GetObjectAPI p_ICBoxMin.handle, Len(tBmp), tBmp
    With tBmp
        lBtnWidth = (.bmWidth / 3)
        lBtnHeight = .bmHeight
    End With
    
    '/* caption bar
    GetObjectAPI p_ICaption.handle, Len(tBmp), tBmp
    With tBmp
        lCapWidth = .bmWidth
        lCapHeight = .bmHeight
    End With
    
    '/* borders
    GetObjectAPI p_IBorders.handle, Len(tBmp), tBmp
    With tBmp
        lBorderWdth = (.bmWidth / 7)
    End With
    
    If Identify_OS And m_TransEnabled Then
        If Not m_RCM Is Nothing Then
            Set m_RCM = Nothing
        End If
        Set m_RCM = New RCMSuperClass
        With m_RCM
            '/* image set
            Set .p_ICaption = m_pCaption
            Set .p_IBorders = m_pBorders
            Set .p_ICBoxMin = m_pCBoxMin
            Set .p_ICBoxMax = m_pCBoxMax
            Set .p_ICBoxRst = m_pCBoxRst
            Set .p_ICBoxCls = m_pCBoxCls
            Set .m_oFrmMain = m_oFrmMain
            '/* max size constraint defaults
            If m_bMaxSize Then
                If m_lMaxHeight = 0 Then
                    .p_MaxHeight = 480
                Else
                    .p_MaxHeight = m_lMaxHeight
                End If
                
                If m_lMaxWidth = 0 Then
                    .p_MaxWidth = 640
                Else
                    .p_MaxWidth = m_lMaxWidth
                End If
                .p_MaxSize = m_bMaxSize
            End If
            
            '/* caption dimensions
            .p_CaptionHeight = lCapHeight
            .p_LeftEnd = m_lLeftEnd
            .p_ActiveRight = m_lRightStart
            .p_RightEnd = lCapWidth
            .p_CustomCaption = m_bCustomCaption
            .p_CaptionFntClr = m_CaptionFntClr
            .p_CustomIcon = m_bCustomIcon
            .p_CaptionOffset = m_lCaptionOffset
            '/* control buttons
            .p_ButtonHeight = lBtnHeight
            .p_ButtonWidth = lBtnWidth
            .p_ControlButtonPosition = m_bCustomButtonPosition
            .p_ButtonOffsetX = m_lControlOffsetX
            .p_ButtonOffsetY = m_lControlOffsetY
            '/* frame borders
            .p_LeftBorderWidth = lBorderWdth
            .p_RightBorderWidth = lBorderWdth
            '/* sizing borders
            .p_BottomSizingBorder = m_lBottomBorderHeight
            .p_TopSizingBorder = m_lTopBorderHeight
            '/* transparency index
            .p_TransIdx = m_lTransIdx
            '/* shadow control
            .p_ShadowForm = m_bShadowForm
            .p_ShadowDepth = m_lShadowDepth
            .p_ShadowInset = m_bShadowInset
            .p_ShadowInDepth = m_lShadowInDepth
            '/* start
            .Set_Host
        End With
    Else
        Set m_Neo = Nothing
        Set m_Neo = New cNeoClass
        With m_Neo
            Set .p_ICaption = m_pCaption
            Set .p_IBorders = m_pBorders
            Set .p_ICBoxMin = m_pCBoxMin
            Set .p_ICBoxMax = m_pCBoxMax
            Set .p_ICBoxRst = m_pCBoxRst
            Set .p_ICBoxCls = m_pCBoxCls
            Set .p_oFrmMain = m_oFrmMain
            '/* caption
            .p_LeftEnd = m_lLeftEnd
            .p_ActiveRight = m_lRightStart
            .p_CustomCaption = m_bCustomCaption
            .p_CaptionFntClr = m_CaptionFntClr
            .p_CaptionOffset = m_lCaptionOffset
            .p_CustomIcon = m_bCustomIcon
            .p_BottomSizingBorder = m_lBottomBorderHeight
            .p_TopSizingBorder = m_lTopBorderHeight
            '/* buttons
            .p_ControlButtonPosition = m_bCustomButtonPosition
            .p_ButtonOffsetX = m_lControlOffsetX
            .p_ButtonOffsetY = m_lControlOffsetY
            .Attach
        End With
    End If

End Sub

Public Function Identify_OS() As Boolean
'/* set os version flag

Dim rOsVersion As OSVersion

    rOsVersion.dwOSVersionInfoSize = Len(rOsVersion)
    If GetVersionEx(rOsVersion) Then
        If rOsVersion.dwPlatformId >= PLATFORM_WIN32_NT Then
            Identify_OS = True
        End If
    End If

End Function

Public Sub Class_Unload()

    If Not m_RCM Is Nothing Then
        m_RCM.Class_Unload
    End If
    
End Sub

Private Sub Class_Terminate()

    If Not m_RCM Is Nothing Then
        Set m_RCM = Nothing
    End If
    If Not m_Neo Is Nothing Then
        Set m_Neo = Nothing
    End If

End Sub

Private Sub m_RCM_eTerminate()
    RaiseEvent eGTerminate
End Sub




