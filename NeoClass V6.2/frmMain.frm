VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "NeoClass V6"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9840
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9840
   StartUpPosition =   1  'CenterOwner
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
      Height          =   5625
      Left            =   0
      ScaleHeight     =   5625
      ScaleWidth      =   9825
      TabIndex        =   6
      Tag             =   "xb"
      Top             =   0
      Width           =   9825
      Begin VB.PictureBox picCommand 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   765
         Left            =   2340
         ScaleHeight     =   765
         ScaleWidth      =   5715
         TabIndex        =   48
         Tag             =   "xb"
         Top             =   4110
         Width           =   5715
         Begin VB.CommandButton Command1 
            Caption         =   "No Skin"
            Height          =   435
            Left            =   4110
            TabIndex        =   52
            Tag             =   "NO"
            Top             =   150
            Width           =   1185
         End
         Begin VB.CommandButton cmdtest 
            Caption         =   "No Icon"
            Height          =   435
            Index           =   2
            Left            =   1470
            TabIndex        =   51
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton cmdtest 
            Caption         =   "Icon Right"
            Height          =   435
            Index           =   1
            Left            =   2790
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   50
            Tag             =   "IR2"
            Top             =   150
            Width           =   1215
         End
         Begin VB.CommandButton cmdtest 
            Caption         =   "Icon Left"
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   49
            Tag             =   "IL3"
            Top             =   150
            Width           =   1245
         End
      End
      Begin VB.TextBox txtDialog 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3075
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   420
         Width           =   6705
      End
      Begin VB.Frame fmStyle 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Examples"
         Height          =   2115
         Left            =   420
         TabIndex        =   33
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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
            TabIndex        =   36
            Tag             =   "xb"
            Top             =   870
            Width           =   855
         End
         Begin VB.OptionButton optStyle 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Frost"
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
            TabIndex        =   35
            Tag             =   "xb"
            Top             =   600
            Value           =   -1  'True
            Width           =   765
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
            TabIndex        =   34
            Tag             =   "xb"
            Top             =   330
            Width           =   1275
         End
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   450
         ScaleHeight     =   1335
         ScaleWidth      =   1095
         TabIndex        =   32
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
      Left            =   0
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   90
      ScaleWidth      =   630
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4290
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
      Left            =   3450
      Picture         =   "frmMain.frx":007B
      ScaleHeight     =   510
      ScaleWidth      =   5745
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2700
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
      Left            =   8130
      Picture         =   "frmMain.frx":99BD
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3240
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
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":A247
      ScaleHeight     =   480
      ScaleWidth      =   5250
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2070
      Width           =   5250
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
      Index           =   1
      Left            =   30
      Picture         =   "frmMain.frx":12609
      ScaleHeight     =   90
      ScaleWidth      =   630
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2580
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
      Height          =   480
      Index           =   2
      Left            =   0
      Picture         =   "frmMain.frx":1294B
      ScaleHeight     =   480
      ScaleWidth      =   5250
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3750
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
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":1AD0D
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1140
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
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":1BF4F
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1380
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
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":1D191
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   1620
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
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":1E3D3
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1830
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
      Height          =   255
      Index           =   4
      Left            =   6840
      Picture         =   "frmMain.frx":1F615
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   2430
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
      Left            =   7620
      Picture         =   "frmMain.frx":200B3
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   2430
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
      Left            =   8400
      Picture         =   "frmMain.frx":20B51
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   2430
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
      Left            =   6060
      Picture         =   "frmMain.frx":215EF
      ScaleHeight     =   255
      ScaleWidth      =   765
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2430
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
      Left            =   0
      Picture         =   "frmMain.frx":2208D
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   3480
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
      Left            =   0
      Picture         =   "frmMain.frx":22880
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   3270
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
      Left            =   0
      Picture         =   "frmMain.frx":22FE7
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   3030
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
      Left            =   0
      Picture         =   "frmMain.frx":236BC
      ScaleHeight     =   240
      ScaleWidth      =   1440
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2790
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":23D51
      ScaleHeight     =   315
      ScaleWidth      =   1350
      TabIndex        =   31
      Top             =   4830
      Width           =   1350
   End
   Begin MSComctlLib.ImageList imlList 
      Left            =   8640
      Top             =   4860
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2436E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24688
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29E7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30114
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31396
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":314F0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picButtons 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   1
      Left            =   1440
      Picture         =   "frmMain.frx":33CA2
      ScaleHeight     =   300
      ScaleWidth      =   2100
      TabIndex        =   40
      Top             =   4860
      Width           =   2100
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
      Left            =   8340
      Picture         =   "frmMain.frx":35DB4
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3690
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
      Left            =   6420
      Picture         =   "frmMain.frx":36AB6
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3690
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
      Left            =   7380
      Picture         =   "frmMain.frx":377B8
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   3690
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
      Left            =   5460
      Picture         =   "frmMain.frx":384BA
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   3690
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
      Left            =   8220
      Picture         =   "frmMain.frx":391BC
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   4410
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
      Left            =   5550
      Picture         =   "frmMain.frx":39A46
      ScaleHeight     =   435
      ScaleWidth      =   3750
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   3960
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
      Left            =   7170
      Picture         =   "frmMain.frx":3EFB8
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   420
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
      Left            =   8220
      Picture         =   "frmMain.frx":402AA
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   420
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
      Left            =   7170
      Picture         =   "frmMain.frx":4159C
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   750
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
      Left            =   8220
      Picture         =   "frmMain.frx":4288E
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   750
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
      Left            =   8520
      Picture         =   "frmMain.frx":43B80
      ScaleHeight     =   105
      ScaleWidth      =   735
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1530
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
      Left            =   5520
      Picture         =   "frmMain.frx":43FCE
      ScaleHeight     =   405
      ScaleWidth      =   3750
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1110
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
      Left            =   0
      Picture         =   "frmMain.frx":48F60
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
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
      Left            =   0
      Picture         =   "frmMain.frx":49CA2
      ScaleHeight     =   375
      ScaleWidth      =   3750
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   510
      Width           =   3750
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
      Left            =   0
      Picture         =   "frmMain.frx":4E654
      ScaleHeight     =   225
      ScaleWidth      =   1575
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   900
      Visible         =   0   'False
      Width           =   1575
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
      Left            =   1050
      Picture         =   "frmMain.frx":4F91A
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
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
      Left            =   0
      Picture         =   "frmMain.frx":5065C
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   240
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
      Left            =   1050
      Picture         =   "frmMain.frx":5139E
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Neo           As cNeoClass

Private Sub Form_Load()

    Set m_Neo = New cNeoClass
    optStyle(0).Value = True

End Sub

Public Sub Set_Skin(ByVal iStyle As Integer)
'/* load skin

On Error Resume Next

    '/* recreate the instance
    '/* needed to flush message queue
    Set m_Neo = Nothing
    Set m_Neo = New cNeoClass
    
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
            Set .p_ICommand = picButtons(0).Picture
            '/* skin command buttons
            .p_SkinCommand = False
            .p_CmdFntClr = &H0
            .p_SkinForm = True
            .p_BorderHasInactive = False
            .p_ControlHasInactive = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 0
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_Offset = 0
            .p_BottomBorderHeight = 2
            .p_TopBorderHeight = 2
            Set .p_ImlRef = imlList
        End With
        
    '/* frost
    Case 1
        With m_Neo
            Set .p_ICaption = picBar(1).Picture
            Set .p_IBorders = picFrame(1).Picture
            Set .p_ICBoxMin = picMin(1).Picture
            Set .p_ICBoxMax = picMax(1).Picture
            Set .p_ICBoxRst = picRst(1).Picture
            Set .p_ICBoxCls = picCls(1).Picture
            Set .p_ICommand = picButtons(0).Picture
            '/* skin command buttons
            .p_SkinCommand = True
            .p_SkinForm = True
            .p_BorderHasInactive = False
            .p_ControlHasInactive = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -8
            .p_ButtonOffsetY = 6
            .p_LeftEnd = 250
            .p_ActiveRight = 251
            .p_Offset = 0
            .p_BottomBorderHeight = 2
            .p_TopBorderHeight = 2
            Set .p_ImlRef = imlList
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
            Set .p_ICommand = picButtons(0).Picture
            '/* skin command buttons
            .p_SkinCommand = True
            .p_SkinForm = True
            .p_BorderHasInactive = False
            .p_ControlHasInactive = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -8
            .p_ButtonOffsetY = 6
            .p_LeftEnd = 250
            .p_ActiveRight = 251
            .p_Offset = 0
            .p_BottomBorderHeight = 2
            .p_TopBorderHeight = 2
            Set .p_ImlRef = imlList
        End With
    
    '/* indigo
    Case 3
        With m_Neo
            Set .p_ICaption = picBar(3).Picture
            Set .p_IBorders = picFrame(3).Picture
            Set .p_ICBoxMin = picMin(3).Picture
            Set .p_ICBoxMax = picMax(3).Picture
            Set .p_ICBoxRst = picRst(3).Picture
            Set .p_ICBoxCls = picCls(3).Picture
            Set .p_ICommand = picButtons(0).Picture
            '/* skin command buttons
            .p_SkinCommand = True
            .p_SkinForm = True
            .p_BorderHasInactive = False
            .p_ControlHasInactive = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -5
            .p_ButtonOffsetY = 1
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_Offset = 0
            .p_BottomBorderHeight = 2
            .p_TopBorderHeight = 2
            Set .p_ImlRef = imlList
        End With
        
    '/* alpha
    Case 4
        With m_Neo
            Set .p_ICaption = picBar(4).Picture
            Set .p_IBorders = picFrame(4).Picture
            Set .p_ICBoxMin = picMin(4).Picture
            Set .p_ICBoxMax = picMax(4).Picture
            Set .p_ICBoxRst = picRst(4).Picture
            Set .p_ICBoxCls = picCls(4).Picture
            Set .p_ICommand = picButtons(0).Picture
            '/* skin command buttons
            .p_SkinCommand = True
            .p_SkinForm = True
            .p_BorderHasInactive = False
            .p_ControlHasInactive = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -5
            .p_ButtonOffsetY = 9
            .p_LeftEnd = 250
            .p_ActiveRight = 251
            .p_Offset = 0
            .p_BottomBorderHeight = 2
            .p_TopBorderHeight = 2
            Set .p_ImlRef = imlList
        End With
    '/* metal
    Case 5
        With m_Neo
            '/* set caption and border images
            Set .p_ICaption = picBar(5).Picture
            Set .p_IBorders = picFrame(5).Picture
            Set .p_ICBoxMin = picMin(5).Picture
            Set .p_ICBoxMax = picMax(5).Picture
            Set .p_ICBoxRst = picRst(5).Picture
            Set .p_ICBoxCls = picCls(5).Picture
            Set .p_ICommand = picButtons(0).Picture
            '/* skin command buttons
            .p_SkinCommand = False
            '/* optional command font color
            .p_CmdFntClr = &H444444
            '/* skin form
            .p_SkinForm = True
            '/* border has inactive image
            .p_BorderHasInactive = False
            '/* button has change image
            .p_ControlHasInactive = True
            '/* use user defined button offsets
            .p_ControlButtonPosition = True
            '/* offset x
            .p_ButtonOffsetX = -15
            '/* offset y
            .p_ButtonOffsetY = 5
            '/* left end of caption image
            .p_LeftEnd = 224
            '/* start right side
            .p_ActiveRight = 225
            '/* caption offset
            .p_Offset = 0
            '/* bottom sizing border height
            .p_BottomBorderHeight = 2
            '/* top sizing border height
            .p_TopBorderHeight = 2
            '/* link image list
            Set .p_ImlRef = imlList
        End With
    End Select
    
    '/* start subclass
    m_Neo.Attach Me
    '/* set control styles
    Set_Style iStyle
    Me.Refresh
    
On Error GoTo 0

End Sub

Private Sub Form_Resize()
'/* resize/repaint controls

Dim lW      As Long
Dim lh      As Long
Dim mCtrl   As Control

On Error Resume Next

    With Me
        lW = .ScaleWidth - 2
        lh = .ScaleHeight - 2
    End With
    
    With picMain
        .Width = lW
        .Height = lh
        .Top = 0
        .Left = 0
    End With
    
    With txtDialog
        .Width = Me.Width - (fmStyle.Width + 1500)
        .Height = (Me.Height - 2400)
        .Top = fmStyle.Top
        .Left = fmStyle.Width + 500
    End With
    
    With picCommand
        .Top = txtDialog.Height + txtDialog.Top + 200
        .Left = txtDialog.Left
    End With
    
On Error GoTo 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set m_Neo = Nothing
    
End Sub

Private Sub Set_Style(ByVal iStyle As Integer)

Dim mControl    As Control
Dim lForeColor  As Long
Dim lBackColor  As Long
Dim lAltColor   As Long

On Error Resume Next

    Select Case iStyle
    '/* vista frost
    Case 0
        lForeColor = &H883F29
        lBackColor = &HF8EEEE
        
    '/* vista leaf
    Case 1
        lForeColor = &H836323
        lBackColor = &HEEEFEE
        
    '/* vista midnight
    Case 2
        lForeColor = &H8E503D
        lBackColor = &HFFFFFF
        
    '/* vista ocean
    Case 3
        lForeColor = &H8E503D
        lBackColor = &HF8F4F3
        
    '/* alpha
    Case 4
        lForeColor = &H444444
        lBackColor = &HFEFEFE
        lAltColor = &HA5AFB4
        
    '/* black steel
    Case 5
        lForeColor = &H883F29
        lBackColor = &HF8EEEE
    End Select
    
    '/* set control styles
    For Each mControl In Me
        With mControl
            If .Tag = "xb" Then
                .ForeColor = lForeColor
                .BackColor = lBackColor
            ElseIf .Tag = "xa" Then
                .ForeColor = lForeColor
                .BackColor = lAltColor
            End If
        End With
    Next mControl
    
    Set_Dialogue iStyle
    
On Error GoTo 0

End Sub

Private Sub Set_Dialogue(ByVal iStyle As Integer)
'/* tips

Dim sDialog As String

    Select Case iStyle
    Case 0
        sDialog = "I have included several hastily rendered skin examples for you to draw upon, " & _
        "(excuse the pun). The next three forms, are my initial " & _
        "interpretation of the upcoming Vista stylings. I can hardly wait.."
             
    Case 1
        sDialog = "The included examples use 24 bit bitmaps for caption frame and button images, " & _
        "but gifs, and jpegs will work fine as well.. "

    Case 2
        sDialog = "The caption button sets are tri-state; up, over and down. " & _
        "Each button state uses it's own unique image set. " & _
        "The caption is a single image, and the frame is seven images;  left side, " & _
        "left corner active, left corner inactive, bottom, right side, and right corner active and inactive."
    
    Case 3
        sDialog = "The subclassing engine is inline and uses the MGSubclass class and " & _
        "MISubclass implements interface. Subclass instances are generated dynamically using " & _
        "the DSIE dynamic instancing engine. This creates a unique subclass instance for every " & _
        "control in the subclass member queue. In this way, there is no need to dim and set a new instance for " & _
        "a control. Instances are added automatically using the MGSubclass class array, and a seven " & _
        "dimensional variant array tracks object state and class heirarchy."
    
    Case 4
        sDialog = "An example of using this technique to subclass common controls in the " & _
        "client area is included as an example. Most controls could be added as additional " & _
        "'spokes' and leverage parallel message processing to respond to events originated " & _
        "by a subclass member object in a way appropriate to that object type."
        
    Case 5
        sDialog = "Well, that's about it, read the notes, poke around a bit, and let me know " & _
        "if you liked it.." & vbNewLine & _
        "John"
    End Select
    
    txtDialog.Text = ""
    txtDialog.Text = sDialog

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
