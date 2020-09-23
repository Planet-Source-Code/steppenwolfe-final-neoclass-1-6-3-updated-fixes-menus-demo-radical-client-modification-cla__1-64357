VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "NeoClass - Lite"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4890
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbout 
      Height          =   2655
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmAbout.frx":0442
      Top             =   180
      Width           =   4245
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   3270
      Picture         =   "frmAbout.frx":0448
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   1350
      Picture         =   "frmAbout.frx":0F49
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   2310
      Picture         =   "frmAbout.frx":1A4A
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   0
      Left            =   390
      Picture         =   "frmAbout.frx":254B
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   0
      Left            =   420
      Picture         =   "frmAbout.frx":304C
      ScaleHeight     =   120
      ScaleWidth      =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   0
      Left            =   390
      Picture         =   "frmAbout.frx":3284
      ScaleHeight     =   420
      ScaleWidth      =   3750
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_NCDA  As cNeoClass

Private Sub Form_Load()
    Set_Skin frmMain.m_iCurrSkin
    Set_About
End Sub

Public Sub Set_Skin(ByVal iStyle As Integer)
'/* load skin

Dim lWidth As Long

On Error Resume Next

    lWidth = frmMain.Text_Size(Me)
    Set m_NCDA = New cNeoClass
    
    Select Case iStyle
    '/* aero
    Case 0
        With m_NCDA
            Set .p_ICaption = frmMain.picBar(0).Picture
            Set .p_IBorders = frmMain.picFrame(0).Picture
            Set .p_ICBoxMin = frmMain.picMin(0).Picture
            Set .p_ICBoxMax = frmMain.picMax(0).Picture
            Set .p_ICBoxRst = frmMain.picRst(0).Picture
            Set .p_ICBoxCls = frmMain.picCls(0).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CaptionOffset = ((Me.Width / Screen.TwipsPerPixelX) - lWidth) / 2
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 0
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
        
    '/* mac
    Case 1
        With m_NCDA
            Set .p_ICaption = frmMain.picBar(1).Picture
            Set .p_IBorders = frmMain.picFrame(1).Picture
            Set .p_ICBoxMin = frmMain.picMin(1).Picture
            Set .p_ICBoxMax = frmMain.picMax(1).Picture
            Set .p_ICBoxRst = frmMain.picRst(1).Picture
            Set .p_ICBoxCls = frmMain.picCls(1).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &H333333
            .p_CaptionOffset = 16
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 5
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
        
    '/* leaf
    Case 2
        With m_NCDA
            Set .p_ICaption = frmMain.picBar(2).Picture
            Set .p_IBorders = frmMain.picFrame(2).Picture
            Set .p_ICBoxMin = frmMain.picMin(2).Picture
            Set .p_ICBoxMax = frmMain.picMax(2).Picture
            Set .p_ICBoxRst = frmMain.picRst(2).Picture
            Set .p_ICBoxCls = frmMain.picCls(2).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CaptionOffset = 16
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 5
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
    
    '/* new age
    Case 3
        With m_NCDA
            Set .p_ICaption = frmMain.picBar(3).Picture
            Set .p_IBorders = frmMain.picFrame(3).Picture
            Set .p_ICBoxMin = frmMain.picMin(3).Picture
            Set .p_ICBoxMax = frmMain.picMax(3).Picture
            Set .p_ICBoxRst = frmMain.picRst(3).Picture
            Set .p_ICBoxCls = frmMain.picCls(3).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CaptionOffset = 18
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -5
            .p_ButtonOffsetY = 1
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
        
    '/* silver
    Case 4
        With m_NCDA
            Set .p_ICaption = frmMain.picBar(4).Picture
            Set .p_IBorders = frmMain.picFrame(4).Picture
            Set .p_ICBoxMin = frmMain.picMin(4).Picture
            Set .p_ICBoxMax = frmMain.picMax(4).Picture
            Set .p_ICBoxRst = frmMain.picRst(4).Picture
            Set .p_ICBoxCls = frmMain.picCls(4).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CaptionOffset = 16
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -4
            .p_ButtonOffsetY = 9
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
    
    '/* panel
    Case 5
        With m_NCDA
            '/* set caption and border images
            Set .p_ICaption = frmMain.picBar(5).Picture
            Set .p_IBorders = frmMain.picFrame(5).Picture
            Set .p_ICBoxMin = frmMain.picMin(5).Picture
            Set .p_ICBoxMax = frmMain.picMax(5).Picture
            Set .p_ICBoxRst = frmMain.picRst(5).Picture
            Set .p_ICBoxCls = frmMain.picCls(5).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CaptionOffset = ((Me.Width / Screen.TwipsPerPixelX) - lWidth) / 2
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -14
            .p_ButtonOffsetY = 5
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .Attach
        End With
    End Select
    
On Error GoTo 0

End Sub

Private Sub Set_About()

Dim sAbout As String

    sAbout = "~ NeoClass Lite 6.3 ~" & vbCrLf & _
    vbCrLf & "Barebones version. Removed DSIE and everything extraneous for speed. Base was NC 6.1.2." & _
    vbCrLf & vbCrLf & "Fixes: March 28, 2006 " & _
    vbCrLf & "Flicker when loading/unloading a second form is fixed." & _
    vbCrLf & "Flicker on caption bar down is fixed." & _
    vbCrLf & "Redraw speed on sizing/moving has been improved." & _
    vbCrLf & "Resource leaks (on exit) are fixed." & _
    vbCrLf & "Properties for image size are now done with api calc, and have been removed." & _
    vbCrLf & "Properties for caption text color and offset have been added."

    txtAbout.Text = sAbout
        
End Sub
