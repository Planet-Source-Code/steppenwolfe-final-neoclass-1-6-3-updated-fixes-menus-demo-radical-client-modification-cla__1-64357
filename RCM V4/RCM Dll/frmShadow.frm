VERSION 5.00
Begin VB.Form frmShadow 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4440
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   3300
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/* adapted (and rewritten/improved) from an example on vbaccelerator.com

Private Type SAFEARRAYBOUND
    cElements                          As Long
    lLbound                            As Long
End Type

Private Type SAFEARRAY2D
    cDims                              As Integer
    fFeatures                          As Integer
    cbElements                         As Long
    cLocks                             As Long
    pvData                             As Long
    Bounds(0 To 1)                     As SAFEARRAYBOUND
End Type

Private Type RGBQUAD
    rgbBlue                            As Byte
    rgbGreen                           As Byte
    rgbRed                             As Byte
    rgbReserved                        As Byte
End Type

Private Type BITMAPINFOHEADER
    biSize                             As Long
    biWidth                            As Long
    biHeight                           As Long
    biPlanes                           As Integer
    biBitCount                         As Integer
    biCompression                      As Long
    biSizeImage                        As Long
    biXPelsPerMeter                    As Long
    biYPelsPerMeter                    As Long
    biClrUsed                          As Long
    biClrImportant                     As Long
End Type

Private Type BITMAPINFO
    bmiHeader                          As BITMAPINFOHEADER
    bmiColors                          As RGBQUAD
End Type

Private Type SIZEAPI
    cx                                 As Long
    cy                                 As Long
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

Private Type BLENDFUNCTION
    BlendOp                            As Byte
    BlendFlags                         As Byte
    SourceConstantAlpha                As Byte
    AlphaFormat                        As Byte
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

Public Enum EShadowType
    RightShadow
    BottomShadow
End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                     lpvSource As Any, _
                                                                     ByVal cbCopy As Long)

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
                                                       pBitmapInfo As BITMAPINFO, _
                                                       ByVal un As Long, _
                                                       lplpVoid As Long, _
                                                       ByVal handle As Long, _
                                                       ByVal dw As Long) As Long

Private Declare Function UpdateLayeredWindow Lib "user32" (ByVal hWnd As Long, _
                                                           ByVal hdcDst As Long, _
                                                           pptDst As Any, _
                                                           psize As Any, _
                                                           ByVal hdcSrc As Long, _
                                                           pptSrc As Any, _
                                                           ByVal crKey As Long, _
                                                           pblend As BLENDFUNCTION, _
                                                           ByVal dwFlags As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, _
                                                     lpRect As RECT) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long


Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    ByVal Y As Long, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long, _
                                                  ByVal nWidth As Long, _
                                                  ByVal nHeight As Long, _
                                                  ByVal bRepaint As Long) As Long

Private m_bInSizeMove              As Boolean
Private m_hDIb                     As Long
Private m_hBmpOld                  As Long
Private m_hDC                      As Long
Private m_lPtr                     As Long
Private m_lShadowSize              As Long
Private m_hWndAttach               As Long
Private m_tBI                      As BITMAPINFO
Private m_eShadowType              As EShadowType
Private m_oShadow                  As Object


Public Property Get p_ShadowSize() As Long
    p_ShadowSize = m_lShadowSize
End Property

Public Property Let p_ShadowSize(ByVal lSize As Long)
    m_lShadowSize = lSize
End Property

Public Property Get p_Shadow() As Object
    Set p_Shadow = m_oShadow
End Property

Public Property Let p_Shadow(ByVal oShadow As Object)
    Set m_oShadow = oShadow
End Property

Public Property Set p_Shadow(ByVal oShadow As Object)
    Set m_oShadow = oShadow
End Property

Public Property Get p_ShadowType() As EShadowType
    p_ShadowType = m_eShadowType
End Property

Public Property Let p_ShadowType(ByVal value As EShadowType)
    m_eShadowType = value
End Property

Private Property Get p_BytesPerLine() As Long
    p_BytesPerLine = m_tBI.bmiHeader.biWidth * 4
End Property

Private Property Get p_DibWidth() As Long
    p_DibWidth = m_tBI.bmiHeader.biWidth
End Property

Private Property Get p_DibHeight() As Long
    p_DibHeight = m_tBI.bmiHeader.biHeight
End Property

Private Function Create_DIB(ByVal lhDC As Long, _
                            ByVal lWidth As Long, _
                            ByVal lHeight As Long, _
                            ByRef hDib As Long) As Boolean

'/* create dib section

    With m_tBI.bmiHeader
        .biSize = Len(m_tBI.bmiHeader)
        .biWidth = lWidth
        .biHeight = lHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
        .biSizeImage = p_BytesPerLine * .biHeight
    End With
    
    hDib = CreateDIBSection(lhDC, m_tBI, 0, m_lPtr, 0, 0)
    Create_DIB = (hDib <> 0)

End Function

Private Function Instance_Shadow(ByVal lWidth As Long, _
                                 ByVal lHeight As Long) As Boolean

'/* create a new dc

    Clean_Up
    m_hDC = CreateCompatibleDC(0)
    If m_hDC <> 0 Then
        If (Create_DIB(m_hDC, lWidth, lHeight, m_hDIb)) Then
            m_hBmpOld = SelectObject(m_hDC, m_hDIb)
            Instance_Shadow = True
        Else
            DeleteObject m_hDC
            m_hDC = 0
        End If
    End If

End Function

Private Sub Clean_Up()

'/* remove devices

    If m_hDC <> 0 Then
        If m_hDIb <> 0 Then
            SelectObject m_hDC, m_hBmpOld
            DeleteObject m_hDIb
        End If
        DeleteObject m_hDC
    End If
    
    m_hDC = 0
    m_hDIb = 0
    m_hBmpOld = 0
    m_lPtr = 0

End Sub

Private Sub Create_Shadow(ByVal bHorizontal As Boolean, _
                          ByVal bLeftTop As Boolean)

'/* apply shadow to form

Dim bDib()      As Byte
Dim X           As Long
Dim Y           As Long
Dim lC          As Long
Dim lInitC      As Long
Dim lSize       As Long
Dim tSA         As SAFEARRAY2D

    With tSA
        .cbElements = 1
        .cDims = 2
        .Bounds(0).lLbound = 0
        .Bounds(0).cElements = m_tBI.bmiHeader.biHeight
        .Bounds(1).lLbound = 0
        .Bounds(1).cElements = p_BytesPerLine()
        .pvData = m_lPtr
    End With
    
    CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
    If bHorizontal Then
        lSize = p_DibHeight
        If Not bLeftTop Then
            For X = 0 To p_BytesPerLine - 1 Step 4
                If (X < lSize * 4) Then
                    lInitC = (255 * X) \ (lSize * 4)
                ElseIf (X >= (p_BytesPerLine - lSize * 4)) Then
                    lInitC = (((p_BytesPerLine - X) * 255) \ (4 * lSize))
                Else
                    lInitC = 255
                End If
                For Y = 0 To p_DibHeight - 1
                    lC = (lInitC * Y) \ p_DibHeight
                    bDib(X + 3, Y) = lC
                    bDib(X + 2, Y) = 0
                    bDib(X + 1, Y) = 0
                    bDib(X, Y) = 0
                Next Y
            Next X
        End If
    Else
        lSize = p_BytesPerLine \ 4
        If Not bLeftTop Then
            For Y = 0 To p_DibHeight - 1
                If (Y >= (p_DibHeight - lSize)) Then
                    lInitC = (255 * (p_DibHeight - Y)) \ lSize
                Else
                    lInitC = 255
                End If
                For X = 0 To p_BytesPerLine - 1 Step 4
                    lC = (lInitC * (p_BytesPerLine - X)) \ p_BytesPerLine
                    bDib(X + 3, Y) = lC
                    bDib(X + 2, Y) = 0
                    bDib(X + 1, Y) = 0
                    bDib(X, Y) = 0
                Next X
            Next Y
        End If
    End If
    CopyMemory ByVal VarPtrArray(bDib), 0&, 4

End Sub

Private Sub Form_Initialize()

    m_lShadowSize = 5

End Sub

Private Sub Form_Load()
'/* set style bits

Dim lExStyle    As Long
Dim lStyle      As Long
Dim tr          As RECT

    lExStyle = &H80000 Or &H20& Or &H8& Or &H80&
    lStyle = &H80000000 Or &H10000000
    SetWindowLong Me.hWnd, (-16), lStyle
    SetWindowLong Me.hWnd, (-20), lExStyle
    GetWindowRect m_oShadow.hWnd, tr
    
    With tr
        Render_Shadow .Left, .Right, .Top, .Bottom, True
    End With
    

End Sub

Friend Sub Render_Shadow(ByVal lLeft As Long, _
                         ByVal lRight As Long, _
                         ByVal lTop As Long, _
                         ByVal lBottom As Long, _
                         ByVal bChange As Boolean)

'/* create the shadow effect

Dim tSize       As SIZEAPI
Dim tBlend      As BLENDFUNCTION
Dim tPtSrc      As POINTAPI
Dim tr          As RECT

On Error Resume Next

    With tr
        .Left = lLeft
        .Right = lRight
        .Top = lTop
        .Bottom = lBottom
    End With
    
    If bChange Then
        With tSize
            If m_eShadowType = BottomShadow Then
                .cx = (tr.Right - tr.Left)
                .cy = m_lShadowSize
            Else
                .cx = m_lShadowSize
                .cy = (tr.Bottom - tr.Top) - m_lShadowSize
            End If
            Instance_Shadow .cx, .cy
        End With
        
        If m_eShadowType = BottomShadow Then
            Create_Shadow True, False
        Else
            Create_Shadow False, False
        End If
        
        With tBlend
            .BlendOp = &H0&
            .BlendFlags = 0
            .AlphaFormat = &H1
            .SourceConstantAlpha = 96
        End With
        With tPtSrc
            .X = 0
            .Y = 0
        End With
        UpdateLayeredWindow Me.hWnd, ByVal 0&, _
        ByVal 0&, tSize, m_hDC, tPtSrc, 0, tBlend, &H2&
    End If
    
    If m_eShadowType = RightShadow Then
        With tr
            MoveWindow Me.hWnd, .Right, (.Top + m_lShadowSize), _
            m_lShadowSize, (.Bottom - .Top), True
        End With
    Else
        With tr
            MoveWindow Me.hWnd, (.Left + m_lShadowSize), _
            .Bottom, (.Right - .Left), m_lShadowSize, True
        End With
    End If
    
On Error GoTo 0

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    Clean_Up

End Sub


