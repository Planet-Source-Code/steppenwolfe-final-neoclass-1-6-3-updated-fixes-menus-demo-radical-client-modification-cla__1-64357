VERSION 5.00
Begin VB.Form frmTest 
   BackColor       =   &H00FFFFFF&
   Caption         =   "RCM Test"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   7605
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   5085
      Left            =   120
      ScaleHeight     =   5085
      ScaleWidth      =   7215
      TabIndex        =   0
      Top             =   60
      Width           =   7215
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sample Styles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   240
         TabIndex        =   16
         Top             =   330
         Width           =   2595
         Begin VB.PictureBox Picture1 
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
            Height          =   1095
            Left            =   90
            ScaleHeight     =   1095
            ScaleWidth      =   2385
            TabIndex        =   17
            Top             =   270
            Width           =   2385
            Begin VB.OptionButton optStyles 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Vista Aero"
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
               Left            =   90
               TabIndex        =   20
               Top             =   90
               Width           =   1335
            End
            Begin VB.OptionButton optStyles 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Indigo"
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
               Index           =   1
               Left            =   90
               TabIndex        =   19
               Top             =   420
               Width           =   1125
            End
            Begin VB.OptionButton optStyles 
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
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   18
               Top             =   750
               Value           =   -1  'True
               Width           =   1335
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Form Transparency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   240
         TabIndex        =   11
         Top             =   2010
         Width           =   2625
         Begin VB.PictureBox Picture2 
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
            Height          =   1035
            Left            =   90
            ScaleHeight     =   1035
            ScaleWidth      =   2445
            TabIndex        =   12
            Top             =   210
            Width           =   2445
            Begin VB.TextBox txtTrans 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   1710
               TabIndex        =   14
               Text            =   "225"
               Top             =   390
               Width           =   435
            End
            Begin VB.CheckBox chkTrans 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Enable (XP/2K)"
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
               Left            =   90
               TabIndex        =   13
               Top             =   420
               Value           =   1  'Checked
               Width           =   1545
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Trans Index (100 - 255)"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   810
               TabIndex        =   15
               Top             =   780
               Width           =   1350
            End
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Shadow Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2445
         Left            =   3270
         TabIndex        =   2
         Top             =   330
         Width           =   2985
         Begin VB.TextBox txtExterior 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1980
            TabIndex        =   6
            Text            =   "12"
            Top             =   600
            Width           =   345
         End
         Begin VB.CheckBox chkExterior 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable"
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
            Left            =   180
            TabIndex        =   5
            Top             =   690
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.TextBox txtInterior 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   1980
            TabIndex        =   4
            Text            =   "8"
            Top             =   1710
            Width           =   345
         End
         Begin VB.CheckBox chkInterior 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Enable"
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
            Left            =   180
            TabIndex        =   3
            Top             =   1800
            Value           =   1  'Checked
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depth (2 - 15)"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1530
            TabIndex        =   10
            Top             =   960
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Depth (2 - 15)"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   1530
            TabIndex        =   9
            Top             =   2100
            Width           =   795
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exterior Shadow"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   8
            Top             =   330
            Width           =   1185
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Interior Shadow"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   1440
            Width           =   1155
         End
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "Reload"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   4560
         TabIndex        =   1
         Top             =   3060
         Width           =   1695
      End
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
      Index           =   0
      Left            =   90
      Picture         =   "frmTest.frx":0000
      ScaleHeight     =   375
      ScaleWidth      =   3750
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   930
      Visible         =   0   'False
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
      Height          =   255
      Index           =   0
      Left            =   90
      Picture         =   "frmTest.frx":61EA
      ScaleHeight     =   255
      ScaleWidth      =   1785
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   1785
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
      Index           =   0
      Left            =   1140
      Picture         =   "frmTest.frx":81C8
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   450
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
      Height          =   240
      Index           =   0
      Left            =   1140
      Picture         =   "frmTest.frx":934A
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   690
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
      Index           =   0
      Left            =   90
      Picture         =   "frmTest.frx":A4CC
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   450
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
      Index           =   0
      Left            =   90
      Picture         =   "frmTest.frx":B64E
      ScaleHeight     =   240
      ScaleWidth      =   1035
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   690
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
      Index           =   1
      Left            =   1140
      Picture         =   "frmTest.frx":C7D0
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2040
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
      Index           =   1
      Left            =   90
      Picture         =   "frmTest.frx":E0DE
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2040
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
      Index           =   1
      Left            =   1140
      Picture         =   "frmTest.frx":F9EC
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   1035
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
      Index           =   1
      Left            =   90
      Picture         =   "frmTest.frx":112FA
      ScaleHeight     =   345
      ScaleWidth      =   1035
      TabIndex        =   30
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
      Index           =   1
      Left            =   120
      Picture         =   "frmTest.frx":12C08
      ScaleHeight     =   105
      ScaleWidth      =   735
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2820
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
      Index           =   1
      Left            =   90
      Picture         =   "frmTest.frx":131A6
      ScaleHeight     =   405
      ScaleWidth      =   3750
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   3750
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
      Index           =   2
      Left            =   2970
      Picture         =   "frmTest.frx":19B60
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   3030
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
      Index           =   2
      Left            =   1050
      Picture         =   "frmTest.frx":1AC5E
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   3030
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
      Index           =   2
      Left            =   2010
      Picture         =   "frmTest.frx":1BD5C
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   3030
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
      Index           =   2
      Left            =   90
      Picture         =   "frmTest.frx":1CE5A
      ScaleHeight     =   255
      ScaleWidth      =   945
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   3030
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
      Index           =   2
      Left            =   90
      Picture         =   "frmTest.frx":1DF58
      ScaleHeight     =   150
      ScaleWidth      =   1050
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   3750
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
      Index           =   2
      Left            =   90
      Picture         =   "frmTest.frx":1EA8A
      ScaleHeight     =   435
      ScaleWidth      =   3750
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   3300
      Visible         =   0   'False
      Width           =   3750
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_RCMG   As RCMgate
Attribute m_RCMG.VB_VarHelpID = -1
Private m_bNoTerminate      As Boolean
Private m_bTransEnabled     As Boolean

Private Sub chkTrans_Click()

    cmdReload_Click
    
End Sub

Private Sub Form_Load()

    m_bTransEnabled = (chkTrans.Value = 1)
    Set_Style 2

End Sub

Private Sub cmdReload_Click()

Dim lSel As Long
    
    m_bTransEnabled = (chkTrans.Value = 1)
    m_bNoTerminate = True
    Select Case True
    Case optStyles(0).Value
        lSel = 0
    Case optStyles(1).Value
        lSel = 1
    Case optStyles(2).Value
        lSel = 2
    End Select
    
    Set_Style lSel
    m_bNoTerminate = False
    
End Sub

Private Sub m_RCMG_eGTerminate()

    If Not m_bNoTerminate Then
        Unload Me
    Else
        Me.Visible = True
    End If
    
End Sub

Private Sub optStyles_Click(Index As Integer)

    Select Case Index
    '/* aero
    Case 0
        txtTrans.Text = "160"
        txtExterior.Text = "12"
        txtInterior.Text = "14"
    
    '/* indigo
    Case 1
        txtTrans.Text = "160"
        txtExterior.Text = "8"
        txtInterior.Text = "3"
    
    '/* black steel
    Case 2
        txtTrans.Text = "200"
        txtExterior.Text = "12"
        txtInterior.Text = "4"
    
    End Select
    
End Sub

Private Sub Set_Style(ByVal lStyle As Long)

Dim lTrans      As Long
Dim lShDepth    As Long
Dim lShInDepth  As Long

    If Not m_RCMG Is Nothing Then
        m_RCMG.Class_Unload
    End If
    
    Set m_RCMG = Nothing
    Set m_RCMG = New RCMgate
    
    '/* user (sleeping) checks
    If Not CLng(txtTrans.Text) > 255 Or CLng(txtTrans.Text) < 100 Then
        lTrans = CLng(txtTrans.Text)
    Else
        MsgBox "Traslucency must be numeric and between 100 and 255!", _
        vbExclamation, "Invalid Paramater"
        txtTrans.Text = "225"
        Exit Sub
    End If
    
    If chkTrans.Value = 0 Then
        lTrans = 255
    End If
    
    If Not CLng(txtExterior.Text) < 1 Or CLng(txtExterior.Text) > 15 Then
        lShDepth = CLng(txtExterior.Text)
    Else
        MsgBox "Exterior shadow must be set between values 1 and 15!", _
        vbExclamation, "Invalid Paramater"
        txtExterior.Text = "8"
        Exit Sub
    End If
    
    If Not CLng(txtInterior.Text) < 1 Or CLng(txtInterior.Text) > 15 Then
        lShInDepth = CLng(txtInterior.Text)
    Else
        MsgBox "Interior shadow must be set between values 1 and 15!", _
        vbExclamation, "Invalid Paramater"
        txtInterior.Text = "12"
        Exit Sub
    End If
    
    '/* os check
    If Not m_RCMG.Identify_OS Then
        chkTrans.Value = 0
        chkTrans.Enabled = False
    End If
    
    Select Case lStyle
    '/* aero
    Case 0
        With m_RCMG
            '/* image set
            Set .p_ICaption = picBar(0).Picture
            Set .p_IBorders = picFrame(0).Picture
            Set .p_ICBoxMin = picMin(0).Picture
            Set .p_ICBoxMax = picMax(0).Picture
            Set .p_ICBoxRst = picRst(0).Picture
            Set .p_ICBoxCls = picCls(0).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CustomIcon = True
            .p_CaptionOffset = 32
            '/* control buttons
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -15
            .p_ButtonOffsetY = 5
            '/* caption dimensions
            '.p_CaptionHeight = 25
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_Offset = 0
            '/* frame borders
            '/* sizing borders
            .p_BottomSizingBorder = 6
            .p_TopSizingBorder = 6
            '/* transparency index
            .p_TransIdx = lTrans
            '/* shadow control
            .p_MaxSize = True
            .p_ShadowForm = (chkExterior.Value = 1)
            .p_ShadowDepth = lShDepth
            .p_ShadowInset = (chkInterior.Value = 1)
            .p_ShadowInDepth = lShInDepth
        End With
    
    '/* indigo
    Case 1
        lTrans = 220
        lShDepth = 14
        lShInDepth = 6
        With m_RCMG
            Set .p_ICaption = picBar(1).Picture
            Set .p_IBorders = picFrame(1).Picture
            Set .p_ICBoxMin = picMin(1).Picture
            Set .p_ICBoxMax = picMax(1).Picture
            Set .p_ICBoxRst = picRst(1).Picture
            Set .p_ICBoxCls = picCls(1).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -5
            .p_ButtonOffsetY = 1
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_Offset = 0
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 6
            .p_TransIdx = lTrans
            .p_MaxSize = True
            .p_ShadowForm = (chkExterior.Value = 1)
            .p_ShadowDepth = lShDepth
            .p_ShadowInset = (chkInterior.Value = 1)
            .p_ShadowInDepth = lShInDepth
        End With
        
    '/* metal
    Case 2
        With m_RCMG
            Set .p_ICaption = picBar(2).Picture
            Set .p_IBorders = picFrame(2).Picture
            Set .p_ICBoxMin = picMin(2).Picture
            Set .p_ICBoxMax = picMax(2).Picture
            Set .p_ICBoxRst = picRst(2).Picture
            Set .p_ICBoxCls = picCls(2).Picture
            Set .p_oFrmMain = Me
            .p_CustomCaption = True
            .p_CaptionFntClr = &HFFFFFF
            .p_CustomIcon = True
            .p_ControlButtonPosition = True
            .p_ButtonOffsetX = -9
            .p_ButtonOffsetY = 0
            .p_LeftEnd = 224
            .p_ActiveRight = 225
            .p_Offset = 0
            .p_BottomSizingBorder = 3
            .p_TopSizingBorder = 0
            .p_TransIdx = lTrans
            .p_MaxSize = True
            .p_ShadowForm = (chkExterior.Value = 1)
            .p_ShadowDepth = lShDepth
            .p_ShadowInset = (chkInterior.Value = 1)
            .p_ShadowInDepth = lShInDepth
        End With
    End Select
    m_RCMG.p_TransForm = m_bTransEnabled
    m_RCMG.Start
    
End Sub

Private Sub Form_Resize()
    
    With picBg
        .Left = 0
        .Top = 200
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Not m_RCMG Is Nothing Then
        Set m_RCMG = Nothing
    End If
    
End Sub


