VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "NeoClass V6"
   ClientHeight    =   5370
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8070
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5370
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   0
      ScaleHeight     =   285
      ScaleWidth      =   7845
      TabIndex        =   47
      Top             =   0
      Width           =   7845
      Begin VB.Label lblHelp 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1110
         TabIndex        =   50
         Top             =   30
         Width           =   315
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   570
         TabIndex        =   49
         Top             =   30
         Width           =   255
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   48
         Top             =   30
         Width           =   240
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5355
      Left            =   -90
      ScaleHeight     =   5355
      ScaleWidth      =   9345
      TabIndex        =   4
      Tag             =   "xb"
      Top             =   210
      Width           =   9345
      Begin VB.CommandButton cmdAbout 
         Caption         =   "About"
         Height          =   615
         Left            =   6210
         TabIndex        =   40
         Top             =   3990
         Width           =   1275
      End
      Begin VB.PictureBox picCommand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   2340
         ScaleHeight     =   765
         ScaleWidth      =   5715
         TabIndex        =   39
         Tag             =   "xb"
         Top             =   4110
         Width           =   5715
      End
      Begin VB.Frame fmStyle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Examples"
         Height          =   2115
         Left            =   420
         TabIndex        =   26
         Tag             =   "xb"
         Top             =   330
         Width           =   1815
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Blue Panel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   5
            Left            =   120
            TabIndex        =   32
            Tag             =   "xb"
            Top             =   1680
            Width           =   1155
         End
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Silver"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Tag             =   "xb"
            Top             =   1410
            Width           =   795
         End
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "New Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   30
            Tag             =   "xb"
            Top             =   1140
            Width           =   1035
         End
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Leaf"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Tag             =   "xb"
            Top             =   870
            Width           =   855
         End
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mac OS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Tag             =   "xb"
            Top             =   600
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Black Steel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   27
            Tag             =   "xb"
            Top             =   330
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   450
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   25
         Top             =   1020
         Width           =   1095
      End
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   90
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   90
      ScaleWidth      =   630
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5220
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   4
      Left            =   3510
      Picture         =   "frmMain.frx":04BD
      ScaleHeight     =   510
      ScaleWidth      =   5745
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3630
      Visible         =   0   'False
      Width           =   5745
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   4
      Left            =   8190
      Picture         =   "frmMain.frx":9DFF
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4170
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":A689
      ScaleHeight     =   480
      ScaleWidth      =   5250
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   4680
      Visible         =   0   'False
      Width           =   5250
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6900
      Picture         =   "frmMain.frx":12A4B
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   7680
      Picture         =   "frmMain.frx":134E9
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   8460
      Picture         =   "frmMain.frx":13F87
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   6120
      Picture         =   "frmMain.frx":14A25
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3360
      Visible         =   0   'False
      Width           =   765
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":154C3
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4410
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":15CB6
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   4200
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":1641D
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   3960
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   60
      Picture         =   "frmMain.frx":16AF2
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3720
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   8400
      Picture         =   "frmMain.frx":17187
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   4620
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   6480
      Picture         =   "frmMain.frx":17E89
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   4620
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   7440
      Picture         =   "frmMain.frx":18B8B
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   4620
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   5520
      Picture         =   "frmMain.frx":1988D
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   4620
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   150
      Index           =   0
      Left            =   8280
      Picture         =   "frmMain.frx":1A58F
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   5340
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   5610
      Picture         =   "frmMain.frx":1AE19
      ScaleHeight     =   435
      ScaleWidth      =   3750
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   4890
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   7230
      Picture         =   "frmMain.frx":2038B
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   8280
      Picture         =   "frmMain.frx":2167D
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   7230
      Picture         =   "frmMain.frx":2296F
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   8280
      Picture         =   "frmMain.frx":23C61
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   105
      Index           =   3
      Left            =   8580
      Picture         =   "frmMain.frx":24F53
      ScaleHeight     =   105
      ScaleWidth      =   735
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2460
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   5580
      Picture         =   "frmMain.frx":253A1
      ScaleHeight     =   405
      ScaleWidth      =   3750
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   60
      Picture         =   "frmMain.frx":2A333
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   60
      Picture         =   "frmMain.frx":2B075
      ScaleHeight     =   375
      ScaleWidth      =   3750
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1110
      Picture         =   "frmMain.frx":2FA27
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   60
      Picture         =   "frmMain.frx":30769
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   1110
      Picture         =   "frmMain.frx":314AB
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1170
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   60
      Picture         =   "frmMain.frx":321ED
      ScaleHeight     =   225
      ScaleWidth      =   1575
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1830
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox picCls 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   2910
      Picture         =   "frmMain.frx":334B3
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picRst 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   990
      Picture         =   "frmMain.frx":33FB4
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picMax 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   1950
      Picture         =   "frmMain.frx":34AB5
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picMin 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   300
      Index           =   1
      Left            =   30
      Picture         =   "frmMain.frx":355B6
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2130
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.PictureBox picFrame 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   120
      Index           =   1
      Left            =   60
      Picture         =   "frmMain.frx":360B7
      ScaleHeight     =   120
      ScaleWidth      =   840
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2850
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.PictureBox picBar 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   420
      Index           =   1
      Left            =   30
      Picture         =   "frmMain.frx":362EF
      ScaleHeight     =   420
      ScaleWidth      =   3750
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   3750
   End
   Begin VB.Menu mFile 
      Caption         =   "File"
      Begin VB.Menu mF 
         Caption         =   "Save"
         Index           =   0
      End
      Begin VB.Menu mF 
         Caption         =   "Save As"
         Index           =   1
      End
      Begin VB.Menu mF 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mF 
         Caption         =   "Exit"
         Index           =   3
      End
   End
   Begin VB.Menu mEdit 
      Caption         =   "Edit"
      Begin VB.Menu mE 
         Caption         =   "Cut"
         Index           =   0
      End
      Begin VB.Menu mE 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu mE 
         Caption         =   "Paste"
         Index           =   2
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "Help"
      Begin VB.Menu mH 
         Caption         =   "Contents"
         Index           =   0
      End
      Begin VB.Menu mH 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mH 
         Caption         =   "About"
         Index           =   2
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    x                                  As Long
    y                                  As Long
End Type

Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, _
                                                                                         ByVal lpsz As String, _
                                                                                         ByVal cbString As Long, _
                                                                                         lpSize As POINTAPI) As Long

Private m_Neo       As cNeoClass
Public m_iCurrSkin  As Integer

Private Sub cmdAbout_Click()
    frmAbout.Show vbModeless, Me
End Sub

Private Sub Form_Load()

    Set m_Neo = New cNeoClass
    optStyle(0).Value = True
    
End Sub

Public Sub Set_Skin(ByVal iStyle As Integer)
'/* load skin

Dim lWidth As Long

On Error Resume Next

    '/* recreate the instance
    '/* needed to flush message queue
    Set m_Neo = Nothing
    Set m_Neo = New cNeoClass
    
    lWidth = frmMain.Text_Size(Me)
    m_iCurrSkin = iStyle
    
    Select Case iStyle
    '/* aero
    Case 0
        With m_Neo
            Set .p_ICaption = picBar(0).Picture
            Set .p_IBorders = picFrame(0).Picture
            Set .p_ICBoxMin = picMin(0).Picture
            Set .p_ICBoxMax = picMax(0).Picture
            Set .p_ICBoxRst = picRst(0).Picture
            Set .p_ICBoxCls = picCls(0).Picture
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
        With m_Neo
            Set .p_ICaption = picBar(1).Picture
            Set .p_IBorders = picFrame(1).Picture
            Set .p_ICBoxMin = picMin(1).Picture
            Set .p_ICBoxMax = picMax(1).Picture
            Set .p_ICBoxRst = picRst(1).Picture
            Set .p_ICBoxCls = picCls(1).Picture
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
        With m_Neo
            Set .p_ICaption = picBar(2).Picture
            Set .p_IBorders = picFrame(2).Picture
            Set .p_ICBoxMin = picMin(2).Picture
            Set .p_ICBoxMax = picMax(2).Picture
            Set .p_ICBoxRst = picRst(2).Picture
            Set .p_ICBoxCls = picCls(2).Picture
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
        With m_Neo
            Set .p_ICaption = picBar(3).Picture
            Set .p_IBorders = picFrame(3).Picture
            Set .p_ICBoxMin = picMin(3).Picture
            Set .p_ICBoxMax = picMax(3).Picture
            Set .p_ICBoxRst = picRst(3).Picture
            Set .p_ICBoxCls = picCls(3).Picture
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
        With m_Neo
            Set .p_ICaption = picBar(4).Picture
            Set .p_IBorders = picFrame(4).Picture
            Set .p_ICBoxMin = picMin(4).Picture
            Set .p_ICBoxMax = picMax(4).Picture
            Set .p_ICBoxRst = picRst(4).Picture
            Set .p_ICBoxCls = picCls(4).Picture
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
        With m_Neo
            '/* set caption and border images
            Set .p_ICaption = picBar(5).Picture
            Set .p_IBorders = picFrame(5).Picture
            Set .p_ICBoxMin = picMin(5).Picture
            Set .p_ICBoxMax = picMax(5).Picture
            Set .p_ICBoxRst = picRst(5).Picture
            Set .p_ICBoxCls = picCls(5).Picture
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

Private Sub Form_Resize()
'/* resize/repaint controls

On Error Resume Next

    With picMain
        .Top = 50
        .Left = 0
    End With
    With picMenu
        .Top = 0
        .Left = 0
    End With

On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set m_Neo = Nothing
End Sub

'/~ menu support is the simplest thing in the world..
Private Sub lblEdit_Click()
    Me.PopupMenu mEdit, 0, lblEdit.Left, lblEdit.Height
End Sub

Private Sub lblFile_Click()
    Me.PopupMenu mFile, 0, lblFile.Left, lblFile.Height
End Sub

Private Sub lblHelp_Click()
    Me.PopupMenu mHelp, 0, lblHelp.Left, lblHelp.Height
End Sub

Private Sub optStyle_Click(Index As Integer)
'/* skins

    Select Case Index
    Case 0
        Set_Skin 0
    Case 1
        Set_Skin 1
    Case 2
        Set_Skin 2
    Case 3
        Set_Skin 3
    Case 4
        Set_Skin 4
    Case 5
        Set_Skin 5
    End Select

End Sub

Public Function Text_Size(frm As Form) As Long

Dim tPnt    As POINTAPI

    '/* get text height/width
    GetTextExtentPoint32 frm.hdc, frm.Caption, Len(frm.Caption) + 5, tPnt
    Text_Size = tPnt.x

End Function
