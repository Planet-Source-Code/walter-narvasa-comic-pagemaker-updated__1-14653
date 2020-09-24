VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Comic Pagemaker Ver. 1.0"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   12990
   Icon            =   "Comic Pagemaker Main.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MousePointer    =   1  'Arrow
   Picture         =   "Comic Pagemaker Main.frx":0442
   ScaleHeight     =   8505
   ScaleWidth      =   12990
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   8
      Left            =   2400
      Picture         =   "Comic Pagemaker Main.frx":7D02
      ScaleHeight     =   6060
      ScaleWidth      =   7875
      TabIndex        =   114
      ToolTipText     =   "Statue"
      Top             =   1800
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   7
      Left            =   2400
      Picture         =   "Comic Pagemaker Main.frx":A3464
      ScaleHeight     =   6060
      ScaleWidth      =   7875
      TabIndex        =   113
      ToolTipText     =   "Statue"
      Top             =   1800
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   6120
      Index           =   6
      Left            =   2400
      Picture         =   "Comic Pagemaker Main.frx":13EBC6
      ScaleHeight     =   6060
      ScaleWidth      =   7875
      TabIndex        =   112
      ToolTipText     =   "Statue"
      Top             =   1800
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   4200
      Index           =   2
      Left            =   2880
      Picture         =   "Comic Pagemaker Main.frx":1DA328
      ScaleHeight     =   4140
      ScaleWidth      =   6420
      TabIndex        =   95
      ToolTipText     =   "Statue"
      Top             =   1680
      Visible         =   0   'False
      Width           =   6480
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   2685
      Index           =   1
      Left            =   2640
      Picture         =   "Comic Pagemaker Main.frx":1DD108
      ScaleHeight     =   2625
      ScaleWidth      =   4380
      TabIndex        =   94
      ToolTipText     =   "Statue"
      Top             =   1680
      Visible         =   0   'False
      Width           =   4440
   End
   Begin VB.PictureBox picCurrent 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   2520
      ScaleHeight     =   315
      ScaleWidth      =   375
      TabIndex        =   90
      Top             =   5400
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox picTile 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   2190
      Index           =   0
      Left            =   2400
      Picture         =   "Comic Pagemaker Main.frx":1DEB7C
      ScaleHeight     =   2130
      ScaleWidth      =   2895
      TabIndex        =   89
      ToolTipText     =   "Statue"
      Top             =   1680
      Visible         =   0   'False
      Width           =   2955
   End
   Begin VB.PictureBox MoreTools 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   5350
      Left            =   10480
      MousePointer    =   1  'Arrow
      ScaleHeight     =   5295
      ScaleWidth      =   2220
      TabIndex        =   46
      ToolTipText     =   "More Tools Or Information"
      Top             =   2300
      Width           =   2275
      Begin VB.PictureBox LineStyleBox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   120
         ScaleHeight     =   1455
         ScaleWidth      =   1935
         TabIndex        =   64
         Top             =   2160
         Width           =   1935
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   74
            Top             =   240
            Value           =   -1  'True
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   73
            Top             =   480
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   72
            Top             =   720
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   71
            Top             =   960
            Width           =   255
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Option1"
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   70
            Top             =   1200
            Width           =   255
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   69
            Top             =   240
            Width           =   1575
            Begin VB.Line Line1 
               BorderColor     =   &H00FFFFFF&
               DrawMode        =   1  'Blackness
               Index           =   0
               X1              =   120
               X2              =   1560
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   68
            Top             =   480
            Width           =   1575
            Begin VB.Line Line1 
               BorderStyle     =   2  'Dash
               DrawMode        =   2  'Blackness
               Index           =   1
               X1              =   120
               X2              =   1560
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   67
            Top             =   720
            Width           =   1575
            Begin VB.Line Line1 
               BorderStyle     =   3  'Dot
               DrawMode        =   1  'Blackness
               Index           =   2
               X1              =   120
               X2              =   1560
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   66
            Top             =   960
            Width           =   1575
            Begin VB.Line Line1 
               BorderStyle     =   4  'Dash-Dot
               DrawMode        =   1  'Blackness
               Index           =   3
               X1              =   120
               X2              =   1560
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   240
            ScaleHeight     =   255
            ScaleWidth      =   1575
            TabIndex        =   65
            Top             =   1200
            Width           =   1575
            Begin VB.Line Line1 
               BorderStyle     =   5  'Dash-Dot-Dot
               DrawMode        =   1  'Blackness
               Index           =   4
               X1              =   120
               X2              =   1560
               Y1              =   120
               Y2              =   120
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Line Style"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   75
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox DrawWidth1 
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   120
         ScaleHeight     =   855
         ScaleWidth      =   2055
         TabIndex        =   60
         Top             =   3600
         Width           =   2055
         Begin VB.PictureBox SampleLine 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   0
            ScaleHeight     =   330
            ScaleWidth      =   1815
            TabIndex        =   62
            Top             =   200
            Width           =   1815
         End
         Begin ComctlLib.Slider Slider1 
            Height          =   255
            Left            =   0
            TabIndex        =   61
            Top             =   600
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   450
            _Version        =   327682
            Min             =   1
            Max             =   20
            SelStart        =   1
            Value           =   1
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Line Width"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.PictureBox Shapeoptions 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   120
         ScaleHeight     =   735
         ScaleWidth      =   1935
         TabIndex        =   56
         Top             =   4560
         Width           =   1935
         Begin VB.OptionButton Option2 
            Caption         =   "Filled"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   58
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Clear"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   57
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Shape Options"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   0
            TabIndex        =   59
            Top             =   0
            Width           =   1260
         End
      End
      Begin VB.PictureBox FillSelect 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   0
         ScaleHeight     =   2055
         ScaleWidth      =   2175
         TabIndex        =   47
         Top             =   0
         Width           =   2175
         Begin VB.OptionButton Option3 
            Caption         =   "Horizontal Line"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   54
            Top             =   360
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Vertical Line"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Upward Diagonal"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   52
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Downward Diagonal"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   51
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Cross"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   50
            Top             =   1320
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Diagonal Cross"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   49
            Top             =   1560
            Width           =   1455
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Solid"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   48
            Top             =   1800
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Fill Pattern"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   120
            Width           =   930
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   2760
      Top             =   7080
   End
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      Height          =   6490
      Left            =   2320
      MousePointer    =   1  'Arrow
      ScaleHeight     =   6495
      ScaleWidth      =   7905
      TabIndex        =   1
      Top             =   1730
      Width           =   7900
      Begin VB.PictureBox Progpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   1320
         Picture         =   "Comic Pagemaker Main.frx":1DFBF3
         ScaleHeight     =   200
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   17
         Top             =   3480
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.PictureBox Progpic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   6600
         ScaleHeight     =   200
         ScaleMode       =   0  'User
         ScaleWidth      =   100
         TabIndex        =   21
         Top             =   9840
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.PictureBox RedoBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   4440
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   217
         TabIndex        =   14
         Top             =   6600
         Visible         =   0   'False
         Width           =   3255
         Begin VB.PictureBox Tempory 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   720
            ScaleHeight     =   81
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   161
            TabIndex        =   18
            Top             =   120
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.PictureBox TempZoom 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         DrawStyle       =   2  'Dot
         DrawWidth       =   16891
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   11520
         ScaleHeight     =   73
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   13
         Top             =   8280
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1320
         Top             =   7080
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.HScrollBar PicXScroll 
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   6240
         Width           =   7635
      End
      Begin VB.VScrollBar PicYScroll 
         Height          =   6195
         Left            =   7645
         MousePointer    =   1  'Arrow
         TabIndex        =   3
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox Corner 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   12960
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   5
         Top             =   10200
         Width           =   375
      End
      Begin VB.PictureBox MainPic 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Lucida Handwriting"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6435
         Left            =   0
         MousePointer    =   1  'Arrow
         ScaleHeight     =   429
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   525
         TabIndex        =   2
         Top             =   0
         Width           =   7875
         Begin VB.PictureBox picTile 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            ClipControls    =   0   'False
            Height          =   4215
            Index           =   5
            Left            =   600
            Picture         =   "Comic Pagemaker Main.frx":1E0245
            ScaleHeight     =   4155
            ScaleWidth      =   6450
            TabIndex        =   106
            ToolTipText     =   "Statue"
            Top             =   1920
            Visible         =   0   'False
            Width           =   6510
         End
         Begin VB.PictureBox picTile 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            ClipControls    =   0   'False
            Height          =   2670
            Index           =   4
            Left            =   240
            Picture         =   "Comic Pagemaker Main.frx":1E32C3
            ScaleHeight     =   2610
            ScaleWidth      =   4440
            TabIndex        =   105
            ToolTipText     =   "Statue"
            Top             =   3480
            Visible         =   0   'False
            Width           =   4500
         End
         Begin VB.PictureBox picTile 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            ClipControls    =   0   'False
            Height          =   2190
            Index           =   3
            Left            =   0
            Picture         =   "Comic Pagemaker Main.frx":1E4D83
            ScaleHeight     =   2130
            ScaleWidth      =   2895
            TabIndex        =   101
            ToolTipText     =   "Statue"
            Top             =   3960
            Visible         =   0   'False
            Width           =   2955
         End
         Begin VB.PictureBox PasteBox 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   4680
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   20
            Top             =   4200
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Shape CopyBox 
            BorderStyle     =   3  'Dot
            DrawMode        =   1  'Blackness
            Height          =   750
            Left            =   2520
            Top             =   3840
            Visible         =   0   'False
            Width           =   750
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   19
            Visible         =   0   'False
            X1              =   152
            X2              =   232
            Y1              =   344
            Y2              =   344
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   18
            Visible         =   0   'False
            X1              =   544
            X2              =   624
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   17
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   16
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   15
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   14
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   13
            Visible         =   0   'False
            X1              =   544
            X2              =   624
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   12
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   11
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   10
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   9
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   8
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   7
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   6
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   5
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   4
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   3
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   2
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   1
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Line PolySide 
            BorderStyle     =   3  'Dot
            Index           =   0
            Visible         =   0   'False
            X1              =   536
            X2              =   616
            Y1              =   384
            Y2              =   384
         End
         Begin VB.Label CloneFrom 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   3240
            TabIndex        =   16
            Top             =   4200
            Visible         =   0   'False
            Width           =   60
         End
         Begin VB.Label TextPic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   615
            Left            =   120
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   5175
         End
         Begin VB.Shape Shape 
            Height          =   2655
            Left            =   360
            Top             =   600
            Visible         =   0   'False
            Width           =   3495
         End
         Begin VB.Line LineDr 
            Visible         =   0   'False
            X1              =   56
            X2              =   216
            Y1              =   32
            Y2              =   224
         End
      End
      Begin VB.PictureBox LBPIC 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   9360
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   6
         Top             =   7800
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox RBPIC 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   10560
         ScaleHeight     =   735
         ScaleWidth      =   615
         TabIndex        =   7
         Top             =   7680
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.PictureBox UndoPicBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   0
         Left            =   480
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   8
         Top             =   6480
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox UndoPicBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   1
         Left            =   840
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   9
         Top             =   6600
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox UndoPicBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   2
         Left            =   840
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   10
         Top             =   6960
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox UndoPicBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   3
         Left            =   1440
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   11
         Top             =   6720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox UndoPicBox 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   2655
         Index           =   4
         Left            =   2400
         ScaleHeight     =   177
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   177
         TabIndex        =   12
         Top             =   6720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.PictureBox ClipDataPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   3000
         ScaleHeight     =   113
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   185
         TabIndex        =   19
         Top             =   6480
         Visible         =   0   'False
         Width           =   2775
      End
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   8235
      Width           =   12990
      _ExtentX        =   22913
      _ExtentY        =   476
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   11
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6509
            MinWidth        =   6509
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Progress Indication"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Picture Filename"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Cursor Position"
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Picture Size"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "1/22/01"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "System Date"
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "3:39 AM"
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "System Time"
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   ""
            Object.Tag             =   ""
            Object.ToolTipText     =   "Selected Tool"
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   882
            MinWidth        =   882
            Key             =   "LBBUT"
            Object.Tag             =   "LBBUT"
            Object.ToolTipText     =   "Left Button Colour"
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   573
            MinWidth        =   573
            Picture         =   "Comic Pagemaker Main.frx":1E5E1E
            Key             =   "CHBUT"
            Object.Tag             =   "CHBUT"
            Object.ToolTipText     =   "Swap Button colours"
         EndProperty
         BeginProperty Panel11 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   882
            MinWidth        =   882
            Key             =   "RBBUT"
            Object.Tag             =   "RBBUT"
            Object.ToolTipText     =   "Right Button Colour"
         EndProperty
      EndProperty
      MousePointer    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Tools 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   6510
      Left            =   410
      MousePointer    =   1  'Arrow
      ScaleHeight     =   6450
      ScaleWidth      =   1395
      TabIndex        =   22
      ToolTipText     =   "Tools"
      Top             =   800
      Width           =   1455
      Begin VB.PictureBox Picture3 
         AutoSize        =   -1  'True
         Height          =   1665
         Left            =   0
         MouseIcon       =   "Comic Pagemaker Main.frx":1E5F28
         MousePointer    =   99  'Custom
         Picture         =   "Comic Pagemaker Main.frx":1E6232
         ScaleHeight     =   1605
         ScaleWidth      =   1335
         TabIndex        =   88
         Top             =   4800
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   11
         Left            =   930
         Picture         =   "Comic Pagemaker Main.frx":1E9143
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   10
         Left            =   680
         Picture         =   "Comic Pagemaker Main.frx":1E9215
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   9
         Left            =   415
         Picture         =   "Comic Pagemaker Main.frx":1E92E7
         Style           =   1  'Graphical
         TabIndex        =   85
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   8
         Left            =   160
         Picture         =   "Comic Pagemaker Main.frx":1E93B9
         Style           =   1  'Graphical
         TabIndex        =   84
         Top             =   4440
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   7
         Left            =   930
         Picture         =   "Comic Pagemaker Main.frx":1E948B
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   4200
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   6
         Left            =   680
         Picture         =   "Comic Pagemaker Main.frx":1E955D
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   4200
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   5
         Left            =   930
         Picture         =   "Comic Pagemaker Main.frx":1E962F
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   4
         Left            =   680
         Picture         =   "Comic Pagemaker Main.frx":1E9701
         Style           =   1  'Graphical
         TabIndex        =   80
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   3
         Left            =   415
         Picture         =   "Comic Pagemaker Main.frx":1E9803
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   4200
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   2
         Left            =   160
         Picture         =   "Comic Pagemaker Main.frx":1E98D5
         Style           =   1  'Graphical
         TabIndex        =   78
         Top             =   4200
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   1
         Left            =   415
         Picture         =   "Comic Pagemaker Main.frx":1E99A7
         Style           =   1  'Graphical
         TabIndex        =   77
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Index           =   0
         Left            =   160
         Picture         =   "Comic Pagemaker Main.frx":1E9A79
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3960
         Width           =   255
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   35
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1E9B4B
         Style           =   1  'Graphical
         TabIndex        =   118
         ToolTipText     =   "Capture"
         Top             =   3140
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   22
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1EC2ED
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Trace"
         Top             =   3140
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   21
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1EC3EF
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Stiple Brush"
         Top             =   3140
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   20
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1EC4F1
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Spray"
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   19
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1EC5F3
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Select Area To Copy"
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   18
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1EC6F5
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Draw Polygon"
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Enabled         =   0   'False
         Height          =   375
         Index           =   17
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1ECA41
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Redo Last Undo"
         Top             =   2390
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   16
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1ECB43
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Undo Last"
         Top             =   2390
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   15
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1ECE85
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Select Colour"
         Top             =   2390
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   14
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1ED1C7
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Zoom Window"
         Top             =   2010
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   13
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1ED52F
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Insert Text"
         Top             =   2010
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   12
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1ED631
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Clone Tool"
         Top             =   2010
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   11
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1ED733
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Draw Rectangle"
         Top             =   1635
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   10
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1ED835
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Draw Elipse"
         Top             =   1635
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   9
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1ED937
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Draw Line"
         Top             =   1635
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   8
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1EDA39
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Steal Colour tool"
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   7
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1EDD85
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Fill Tool"
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   6
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1EE0EF
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Smudge"
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   5
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1EE431
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "Brush Tool"
         Top             =   885
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   4
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1EE746
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Pencil Tool"
         Top             =   885
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   3
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1EEA56
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Print Picture"
         Top             =   885
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   2
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1EED98
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Save Picture File"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   1
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1EF0DA
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Open Picture File"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1EF41C
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "New Picture File"
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   34
         Left            =   870
         Picture         =   "Comic Pagemaker Main.frx":1EF75E
         Style           =   1  'Graphical
         TabIndex        =   115
         ToolTipText     =   "Paste"
         Top             =   500
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   33
         Left            =   500
         Picture         =   "Comic Pagemaker Main.frx":1EF8A0
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Cut"
         Top             =   500
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   32
         Left            =   120
         Picture         =   "Comic Pagemaker Main.frx":1EF99A
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Cut"
         Top             =   500
         UseMaskColor    =   -1  'True
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Brush Option"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   165
         TabIndex        =   119
         Top             =   3720
         Width           =   1215
      End
   End
   Begin VB.PictureBox Advanced 
      Height          =   6510
      Left            =   410
      ScaleHeight     =   6450
      ScaleWidth      =   1395
      TabIndex        =   97
      Top             =   800
      Visible         =   0   'False
      Width           =   1455
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   31
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F0664
         Style           =   1  'Graphical
         TabIndex        =   109
         ToolTipText     =   "Comic Border Type 3"
         Top             =   5760
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   30
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F132E
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "Comic Border Type 2"
         Top             =   5160
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   29
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F1FF8
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "Comic Border Type 1"
         Top             =   4560
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   28
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F2CC2
         Style           =   1  'Graphical
         TabIndex        =   104
         ToolTipText     =   "Balloon Dialog Type 2 Size 3"
         Top             =   3480
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   27
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F398C
         Style           =   1  'Graphical
         TabIndex        =   103
         ToolTipText     =   "Balloon Dialog Type 2 Size 2"
         Top             =   2880
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   26
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F4656
         Style           =   1  'Graphical
         TabIndex        =   102
         ToolTipText     =   "Balloon Dialog Type 2 Size 1"
         Top             =   2280
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   25
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F5320
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Balloon Dialog Type 1 Size 3"
         Top             =   1680
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   24
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F5FEA
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Balloon Dialog Type 1 Size 2"
         Top             =   1080
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Height          =   550
         Index           =   23
         Left            =   180
         Picture         =   "Comic Pagemaker Main.frx":1F6CB4
         Style           =   1  'Graphical
         TabIndex        =   98
         ToolTipText     =   "Balloon Dialog Type 1 Size 1"
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Comic Border"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   180
         TabIndex        =   111
         Top             =   4120
         Width           =   780
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Balloon Dialog"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   180
         TabIndex        =   110
         Top             =   60
         Width           =   1215
      End
   End
   Begin VB.Image MenuButton 
      Height          =   255
      Left            =   840
      MouseIcon       =   "Comic Pagemaker Main.frx":1F797E
      MousePointer    =   99  'Custom
      ToolTipText     =   "Click here to enable Advanced Toolbox."
      Top             =   7560
      Width           =   495
   End
   Begin VB.Label MenuLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Basic Toolbox"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   600
      TabIndex        =   96
      Top             =   7350
      Width           =   1035
   End
   Begin VB.Label lblImgNum 
      Caption         =   "lblImgNum"
      Height          =   255
      Left            =   7680
      TabIndex        =   93
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblType 
      Caption         =   "lblType"
      Height          =   255
      Left            =   8040
      TabIndex        =   92
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblName 
      Caption         =   "lblName"
      Height          =   255
      Left            =   8040
      TabIndex        =   91
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Menu Fil 
      Caption         =   "File"
      Begin VB.Menu NewImg 
         Caption         =   "&New Image"
      End
      Begin VB.Menu OpFil 
         Caption         =   "&Open Picture"
         Shortcut        =   ^O
      End
      Begin VB.Menu SavFil 
         Caption         =   "Sa&ve Picutre"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSeparator1 
         Caption         =   "-"
      End
      Begin VB.Menu Twan 
         Caption         =   "Twain"
         Begin VB.Menu SelSrc 
            Caption         =   "Select Source"
         End
         Begin VB.Menu AquImge 
            Caption         =   "Aquire Image"
         End
      End
      Begin VB.Menu Pr 
         Caption         =   "Printer"
         Begin VB.Menu PrSet 
            Caption         =   "Printer Setup"
         End
         Begin VB.Menu PrintPicNow 
            Caption         =   "Print Picture"
         End
      End
      Begin VB.Menu sepfil 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRecentFile 
         Caption         =   ""
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu SE 
         Caption         =   "-"
      End
      Begin VB.Menu Endprog 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edi 
      Caption         =   "Edit"
      Begin VB.Menu CutPic 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu CopyPic 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu PastePic 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu UndoLast 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu Redolast 
         Caption         =   "&Redo"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu CapScre 
         Caption         =   "Capture Screen"
         Shortcut        =   ^{F12}
      End
   End
   Begin VB.Menu Vew 
      Caption         =   "View"
      Begin VB.Menu Rilers 
         Caption         =   "Horizontal Ruler"
      End
      Begin VB.Menu Rulers 
         Caption         =   "Vertical Ruler"
      End
      Begin VB.Menu seper1 
         Caption         =   "-"
      End
      Begin VB.Menu BasicToolbox 
         Caption         =   "Basic Toolbox"
      End
      Begin VB.Menu AdvancedToolbox 
         Caption         =   "Advanced Toolbox"
      End
      Begin VB.Menu seper2 
         Caption         =   "-"
      End
      Begin VB.Menu Proview 
         Caption         =   "Progressbar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu SpecEff 
      Caption         =   "Special Effects"
      Begin VB.Menu Filt 
         Caption         =   "Filters"
         Begin VB.Menu Blurr 
            Caption         =   "Blur"
            Begin VB.Menu Blur 
               Caption         =   "Blur"
            End
            Begin VB.Menu Blur_more 
               Caption         =   "Blur More..."
            End
         End
         Begin VB.Menu Emoss 
            Caption         =   "Emboss"
         End
         Begin VB.Menu Sharp 
            Caption         =   "Sharpen"
            Begin VB.Menu Sharpen 
               Caption         =   "Sharpen"
            End
            Begin VB.Menu Sharpen_more 
               Caption         =   "Sharpen More..."
            End
         End
         Begin VB.Menu Difu 
            Caption         =   "Diffuse"
            Begin VB.Menu Diffuse 
               Caption         =   "Diffuse"
            End
            Begin VB.Menu Diffuse_more 
               Caption         =   "Diffuse More..."
            End
         End
         Begin VB.Menu bright 
            Caption         =   "Brightness"
         End
         Begin VB.Menu aqua 
            Caption         =   "Aqua"
         End
         Begin VB.Menu grey 
            Caption         =   "Grey Scale"
         End
         Begin VB.Menu invert 
            Caption         =   "Invert"
         End
         Begin VB.Menu BandW 
            Caption         =   "Black and White"
         End
         Begin VB.Menu pixelatew 
            Caption         =   "Pixelate"
         End
         Begin VB.Menu Cir 
            Caption         =   "Circular"
         End
      End
      Begin VB.Menu Efect 
         Caption         =   "Effects"
         Begin VB.Menu Flip 
            Caption         =   "Flip"
            Begin VB.Menu Flip_horiz 
               Caption         =   "Horizontal"
            End
            Begin VB.Menu Flip_vert 
               Caption         =   "Vertical"
            End
         End
         Begin VB.Menu Rotate 
            Caption         =   "Rotate"
         End
         Begin VB.Menu Rep_col 
            Caption         =   "Replace Colour"
         End
         Begin VB.Menu WAVE 
            Caption         =   "Wave"
         End
         Begin VB.Menu RepIndCol 
            Caption         =   "Replace Individual Colour"
         End
         Begin VB.Menu Lite 
            Caption         =   "Lighting"
         End
         Begin VB.Menu Ham 
            Caption         =   "Hammered"
         End
      End
   End
   Begin VB.Menu Hlpe 
      Caption         =   "Help"
      Begin VB.Menu Hlp 
         Caption         =   "Comic Pagemaker Help Index"
         Index           =   1
      End
      Begin VB.Menu Abt 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl, yl, xtl, ytl, sx, sy, sha As Boolean, fillshape As Boolean, cx, cy, clone As Boolean, cl, xoff, yoff, filarea, texx, texy, re, gre, bl, dwn
Dim XPos As Long
    'this carries the X information which the tiles
    'will be BitBlted on
Dim YPos As Long
    'this carries the Y information which the tiles
    'will be BitBlted on


Private Sub Abt_Click()
    Form8.Show
End Sub

Private Sub AdvancedToolbox_Click()
    BasicToolbox.Checked = False
    AdvancedToolbox.Checked = True
    Tools.Visible = False
    Advanced.Visible = True
End Sub

Private Sub aqua_Click()
Dim tColQ As Long
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
    For i = 0 To xf
        For j = 0 To yf
            tColQ = GetPixel(MainPic.hdc, i, j)
            r = tColQ Mod 256
            g = (tColQ \ 256) Mod 256
            b = tColQ \ 256 \ 256
            r = (g - b) ^ 2 / 125
            g = (r - b) ^ 2 / 125
            b = (r - g) ^ 2 / 125
            SetPixelV MainPic.hdc, i, j, RGB(r, g, b)
        Next j
        progress = i * 100 \ (xf - 1)
        Call progressbar
    Next i
    MainPic.Refresh
End Sub

Private Sub AquImge_Click()
On Error GoTo BadScan
Screen.MousePointer = 11
filne = App.Path + "\scan.bmp"
S% = TWAIN_AcquireToFilename(Me.hwnd, filne)
If S% = 0 Then
   MainPic.Picture = LoadPicture(filne)
   Kill filne
Else
  GoTo BadScan
End If
PicName = "Scan.BMP"
StatusBar1.Panels.Item(2) = PicName
xs = ScaleX(MainPic.Width, vbTwips, vbPixels) - 4
ys = ScaleY(MainPic.Height, vbTwips, vbPixels) - 4
siz = xs & "," & ys
StatusBar1.Panels.Item(4) = siz
Call scroll_val(MainPic.Width, MainPic.Height)
Call clearundo
Screen.MousePointer = 0
Exit Sub
BadScan:
MsgBox "Scan has been aborted", vbInformation, ""
Screen.MousePointer = 0
End Sub

Private Sub BandW_Click()
Dim col As Long
Call Oldpic
    For i = 0 To MainPic.ScaleWidth
        For j = 0 To MainPic.ScaleHeight
            col = GetPixel(MainPic.hdc, i, j)
            r = col Mod 256
            g = (col Mod 256) \ 256
            b = col \ 256 \ 256
            If r < 200 And g < 200 And b < 200 Then
                col = vbBlack
            Else
                col = vbWhite
            End If
            SetPixelV MainPic.hdc, i, j, col
        Next j
        progress = (100 / MainPic.ScaleWidth) * i
        Call progressbar
    Next i
MainPic.Refresh
End Sub

Private Sub BasicToolbox_Click()
    BasicToolbox.Checked = True
    AdvancedToolbox.Checked = False
    Tools.Visible = True
    Advanced.Visible = False
End Sub

Private Sub Blur_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
Call loading(i, j)
    For i = 1 To yf - 2
        For j = 1 To xf - 2
            Red = ImageArray(0, i - 1, j - 1) + ImageArray(0, i - 1, j) + ImageArray(0, i - 1, j + 1) + _
            ImageArray(0, i, j - 1) + ImageArray(0, i, j) + ImageArray(0, i, j + 1) + _
            ImageArray(0, i + 1, j - 1) + ImageArray(0, i + 1, j) + ImageArray(0, i + 1, j + 1)
            Green = ImageArray(1, i - 1, j - 1) + ImageArray(1, i - 1, j) + ImageArray(1, i - 1, j + 1) + _
            ImageArray(1, i, j - 1) + ImageArray(1, i, j) + ImageArray(1, i, j + 1) + _
            ImageArray(1, i + 1, j - 1) + ImageArray(1, i + 1, j) + ImageArray(1, i + 1, j + 1)
            Blue = ImageArray(2, i - 1, j - 1) + ImageArray(2, i - 1, j) + ImageArray(2, i - 1, j + 1) + _
            ImageArray(2, i, j - 1) + ImageArray(2, i, j) + ImageArray(2, i, j + 1) + _
            ImageArray(2, i + 1, j - 1) + ImageArray(2, i + 1, j) + ImageArray(2, i + 1, j + 1)
            SetPixelV MainPic.hdc, j, i, RGB(Red / 9, Green / 9, Blue / 9)
        Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub
Private Sub Blur_more_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
Call loading(i, j)
    For i = 2 To yf - 2
        For j = 2 To xf - 2
            Red = ImageArray(0, i - 2, j - 2) + ImageArray(0, i - 2, j - 1) + ImageArray(0, i - 2, j) + _
            ImageArray(0, i - 2, j + 1) + ImageArray(0, i - 2, j + 2) + ImageArray(0, i - 1, j - 2) + _
            ImageArray(0, i - 1, j - 1) + ImageArray(0, i - 1, j) + ImageArray(0, i - 1, j + 1) + _
            ImageArray(0, i - 1, j + 2) + ImageArray(0, i, j - 2) + _
            ImageArray(0, i, j - 1) + ImageArray(0, i, j) + ImageArray(0, i, j + 1) + _
            ImageArray(0, i, j + 2) + ImageArray(0, i + 1, j - 2) + _
            ImageArray(0, i + 1, j - 1) + ImageArray(0, i + 1, j) + ImageArray(0, i + 1, j + 1) + _
            ImageArray(0, i + 1, j + 2) + ImageArray(0, i + 2, j - 2) + ImageArray(0, i + 2, j - 1) + _
            ImageArray(0, i + 2, j) + ImageArray(0, i + 2, j + 1) + ImageArray(0, i + 2, j + 2)
            Green = ImageArray(1, i - 2, j - 2) + ImageArray(1, i - 2, j - 1) + ImageArray(1, i - 2, j) + _
            ImageArray(1, i - 2, j + 1) + ImageArray(1, i - 2, j + 2) + ImageArray(1, i - 1, j - 2) + _
            ImageArray(1, i - 1, j - 1) + ImageArray(1, i - 1, j) + ImageArray(1, i - 1, j + 1) + _
            ImageArray(1, i - 1, j + 2) + ImageArray(1, i, j - 2) + _
            ImageArray(1, i, j - 1) + ImageArray(1, i, j) + ImageArray(1, i, j + 1) + _
            ImageArray(1, i, j + 2) + ImageArray(1, i + 1, j - 2) + _
            ImageArray(1, i + 1, j - 1) + ImageArray(1, i + 1, j) + ImageArray(1, i + 1, j + 1) + _
            ImageArray(1, i + 1, j + 2) + ImageArray(1, i + 2, j - 2) + ImageArray(1, i + 2, j - 1) + _
            ImageArray(1, i + 2, j) + ImageArray(1, i + 2, j + 1) + ImageArray(1, i + 2, j + 2)
            Blue = ImageArray(2, i - 2, j - 2) + ImageArray(2, i - 2, j - 1) + ImageArray(2, i - 2, j) + _
            ImageArray(2, i - 2, j + 1) + ImageArray(2, i - 2, j + 2) + ImageArray(2, i - 1, j - 2) + _
            ImageArray(2, i - 1, j - 1) + ImageArray(2, i - 1, j) + ImageArray(2, i - 1, j + 1) + _
            ImageArray(2, i - 1, j + 2) + ImageArray(2, i, j - 2) + _
            ImageArray(2, i, j - 1) + ImageArray(2, i, j) + ImageArray(2, i, j + 1) + _
            ImageArray(2, i, j + 2) + ImageArray(2, i + 1, j - 2) + _
            ImageArray(2, i + 1, j - 1) + ImageArray(2, i + 1, j) + ImageArray(2, i + 1, j + 1) + _
            ImageArray(2, i + 1, j + 2) + ImageArray(2, i + 2, j - 2) + ImageArray(2, i + 2, j - 1) + _
            ImageArray(2, i + 2, j) + ImageArray(2, i + 2, j + 1) + ImageArray(2, i + 2, j + 2)
            SetPixelV MainPic.hdc, j, i, RGB(Red / 25, Green / 25, Blue / 25)
        Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub
Private Sub bright_Click()
txt = InputBox("Enter Brightness Level (0-200), below 100 darker, above 100 lighter", "Brightness", 100)
txtbrightness = Val(txt)
If txtbrightness < 0 Or txtbrightness > 200 Then Exit Sub
Dim Brightness As Single
Dim NewColor As Long
Dim X, Y As Integer
Dim r, g, b As Integer
Brightness = txtbrightness / 100
For X = 0 To MainPic.ScaleWidth
For Y = 0 To MainPic.ScaleHeight
NewColor = GetPixel(MainPic.hdc, X, Y)
r = (NewColor Mod 256)
b = (Int(NewColor / 65536))
g = ((NewColor - (b * 65536) - r) / 256)
r = r * Brightness
b = b * Brightness
g = g * Brightness
If r > 255 Then r = 255
If r < 0 Then r = 0
If b > 255 Then b = 255
If b < 0 Then b = 0
If g > 255 Then g = 255
If g < 0 Then g = 0
SetPixelV MainPic.hdc, X, Y, RGB(r, g, b)
Next Y
progress = X * (100 / MainPic.ScaleWidth)
Call progressbar
Next X
MainPic.Refresh
End Sub
Private Sub CapScre_Click()
Me.WindowState = vbMinimized
DoEvents
screencapture 0, 0, Screen.Width / Screen.TwipsPerPixelX, Screen.Height / Screen.TwipsPerPixelY
MainPic.Picture = Clipboard.GetData()
Me.WindowState = vbMaximized
Me.SetFocus
PicName = "Screen Capture.BMP"
StatusBar1.Panels.Item(2) = PicName
xs = ScaleX(MainPic.Width, vbTwips, vbPixels) - 4
ys = ScaleY(MainPic.Height, vbTwips, vbPixels) - 4
siz = xs & "," & ys
StatusBar1.Panels.Item(4) = siz
Call scroll_val(MainPic.Width, MainPic.Height)
Call clearundo
End Sub

Private Sub Cir_Click()
Dim pix As Long
ans = InputBox("Enter Value (1-10)", "Pixel Size", 5)
If ans = "" Then Exit Sub
pix = Val(ans)
If pix < 1 Or pix > 10 Then Exit Sub
Call Oldpic
Call circpix(pix)
End Sub

Private Sub Command1_Click(Index As Integer)
If Index <> 18 Then CopyBox.Visible = False
If Index = 0 Then StatusBar1.Panels.Item(8) = "": Call newpic(0)
If Index = 1 Then OpFil_Click: ShowBalloonDialog = False
If Index = 2 Then StatusBar1.Panels.Item(8) = "": ShowBalloonDialog = False
If Index = 3 Then StatusBar1.Panels.Item(8) = "": Call PrintPicture: ShowBalloonDialog = False
If Index = 4 Then StatusBar1.Panels.Item(8) = "Pencil": ShowBalloonDialog = False
If Index = 5 Then StatusBar1.Panels.Item(8) = "Brush": ShowBalloonDialog = False
If Index = 6 Then StatusBar1.Panels.Item(8) = "Smudge": ShowBalloonDialog = False
If Index = 7 Then StatusBar1.Panels.Item(8) = "Fill Region": ShowBalloonDialog = False
If Index = 8 Then StatusBar1.Panels.Item(8) = "Steal Colour": ShowBalloonDialog = False
If Index = 9 Then StatusBar1.Panels.Item(8) = "Draw Line": ShowBalloonDialog = False
If Index = 10 Then StatusBar1.Panels.Item(8) = "Draw Elipse": ShowBalloonDialog = False
If Index = 11 Then StatusBar1.Panels.Item(8) = "Draw Rectangle": ShowBalloonDialog = False
If Index = 12 Then StatusBar1.Panels.Item(8) = "Clone Tool": ShowBalloonDialog = False
If Index = 13 Then StatusBar1.Panels.Item(8) = "Insert Text": ShowBalloonDialog = False
If Index = 14 Then StatusBar1.Panels.Item(8) = "View Zoom": Form2.Show: ShowBalloonDialog = False
If Index = 15 Then StatusBar1.Panels.Item(8) = "": CommonDialog1.ShowColor: lb = CommonDialog1.Color: Call show_cols: ShowBalloonDialog = False
If Index = 16 Then Call UndoLast_Click: ShowBalloonDialog = False
If Index = 17 Then Call Redolast_Click: ShowBalloonDialog = False
If Index = 18 Then NumSides = InputBox("How Many Sides (3 to 20)", "Number of Sides to Polygon", 6): StatusBar1.Panels.Item(8) = "Polygon": ShowBalloonDialog = False
If Index = 19 Then StatusBar1.Panels.Item(8) = "Select Area": ShowBalloonDialog = False
If Index = 20 Then StatusBar1.Panels.Item(8) = "Spray Can": ShowBalloonDialog = False
If Index = 21 Then StatusBar1.Panels.Item(8) = "Stiple": ShowBalloonDialog = False
If Index = 22 Then StatusBar1.Panels.Item(8) = "Trace": GoTo trcset: ShowBalloonDialog = False
If Index = 23 Then TileSize = 200: Call picTile_Click(0)
If Index = 24 Then TileSize = 295: Call picTile_Click(1)
If Index = 25 Then TileSize = 430: Call picTile_Click(2)
If Index = 26 Then TileSize = 200: Call picTile_Click(3)
If Index = 27 Then TileSize = 295: Call picTile_Click(4)
If Index = 28 Then TileSize = 430: Call picTile_Click(5)
If Index = 29 Then TileSize = 550: Call picTile_Click(6)
If Index = 30 Then TileSize = 550: Call picTile_Click(7)
If Index = 31 Then TileSize = 550: Call picTile_Click(8)
If Index = 32 Then Call CutPic_Click
If Index = 33 Then Call CopyPic_Click
If Index = 34 Then Call PastePic_Click
If Index = 35 Then Call CapScre_Click
If NumSides > 12 Then NumSides = 20
If NumSides < 3 Then NumSides = 3
GoTo doo
trcset:
Form7.Show
wid = MainPic.Width
hei = MainPic.Height
Form7.Width = wid + 20
Form7.Height = hei + 650
Form7.Picture1.Width = MainPic.Width
Form7.Picture1.Height = MainPic.Height
Form7.Command1.left = (Form7.Width / 2) - (Form7.Command1.Width / 2)
Form7.Command1.top = Form7.Picture1.Height + 50

doo:
End Sub
Private Sub Command2_Click(Index As Integer)
BrushType = Index
End Sub

Private Sub CopyPic_Click()
Call putToClipBoard(0)
End Sub

Private Sub CutPic_Click()
    Call putToClipBoard(1)
End Sub

Private Sub Diffuse_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
Call loading(i, j)
    For i = 2 To yf - 2
        For j = 2 To xf - 2
            AletatorioX = Rnd * 3 - 2
            AletatorioY = Rnd * 3 - 2
            Red = ImageArray(0, i + AletatorioX, j + AletatorioY)
            Green = ImageArray(1, i + AletatorioX, j + AletatorioY)
            Blue = ImageArray(2, i + AletatorioX, j + AletatorioY)
            SetPixelV MainPic.hdc, j, i, RGB(Red, Green, Blue)
        Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub
Private Sub Diffuse_more_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
Call loading(i, j)
For i = 2 To yf - 2
    For j = 2 To xf - 2
        AletatorioX = Rnd * 6 - 2
        AletatorioY = Rnd * 6 - 2
        Red = ImageArray(0, i + AletatorioX, j + AletatorioY)
        Green = ImageArray(1, i + AletatorioX, j + AletatorioY)
        Blue = ImageArray(2, i + AletatorioX, j + AletatorioY)
        SetPixelV MainPic.hdc, j, i, RGB(Red, Green, Blue)
    Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub
Private Sub Emoss_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
   Call loading(i, j)
    For i = 1 To yf - 2
        For j = 1 To xf - 2
            Red = Abs(ImageArray(0, i, j) - ImageArray(0, i + 1, j + 1) + 128)
            Green = Abs(ImageArray(1, i, j) - ImageArray(1, i + 1, j + 1) + 128)
            Blue = Abs(ImageArray(2, i, j) - ImageArray(2, i + 1, j + 1) + 128)
            SetPixelV MainPic.hdc, j, i, RGB(Red, Green, Blue)
        Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub

Private Sub Endprog_Click()
Unload Me
End
End Sub
Private Sub Flip_horiz_Click()
Call Oldpic
    Set Tempory.Picture = MainPic.Picture
    px% = MainPic.ScaleWidth
    py% = MainPic.ScaleHeight
    Set Tempory = LoadPicture()
    ret = StretchBlt(Tempory.hdc, px%, 0, -px%, py%, MainPic.hdc, 0, 0, px%, py%, SRCCOPY): Tempory.Refresh
tmpfil = App.Path & "\temp.bmp"
SavePicture Tempory.Image, tmpfil
Set MainPic = LoadPicture(tmpfil)
End Sub
Private Sub Flip_vert_Click()
Call Oldpic
Set Tempory.Picture = MainPic.Picture
    px% = MainPic.ScaleWidth
    py% = MainPic.ScaleHeight
    Set Tempory = LoadPicture()
ret = StretchBlt(Tempory.hdc, 0, py%, px%, -py%, MainPic.hdc, 0, 0, px%, py%, SRCCOPY): Tempory.Refresh
tmpfil = App.Path & "\temp.bmp"
SavePicture Tempory.Image, tmpfil
Set MainPic = LoadPicture(tmpfil)
End Sub
Private Sub Form_Load()
If TWAIN_IsAvailable() = 0 Then Twan.Enabled = False Else Twan.Enabled = True
Call newpic(1)
fillshape = False: clone = False: filarea = 0: zoomactive = False
Form1.MainPic.FontSize = 10
'LineStyleBox.left = 50: LineStyleBox.top = 200
'Shapeoptions.left = 50: Shapeoptions.top = 1920
'FillSelect.left = 50: FillSelect.top = 2890
'BrushShapes.left = 2200: BrushShapes.top = 200
'DrawWidth1.left = 2200: DrawWidth1.top = 1700
'ColSelectBox.left = 2200: ColSelectBox.top = 2800
GetRecentFiles
Call CreateTiles
ShowBalloonDialog = False
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y > 10605 Then GoTo en
If X > (PicBack.Width - 10) And X < (PicBack.Width + 10) Then Form1.MousePointer = 9 Else Form1.MousePointer = 1
If X > Screen.Width - 650 Then GoTo en
If X < 400 Then GoTo en
If Button = 1 Then resiz (X)
en:
End Sub
Private Sub resiz(X)
'PicBack.Width = X: Picture2.left = X + 30: Picture2.Width = Screen.Width - (X + 50): Tools.Width = Picture2.Width: MoreTools.Width = Picture2.Width
'PicYScroll.left = PicBack.Width - 300: PicXScroll.Width = PicYScroll.left
'Corner.top = PicXScroll.top: Corner.left = PicYScroll.left
'TitleBar.Width = Tools.Width - 50
'TOT = Picture2.Width
'TOT = TOT - 120
'TOT = TOT / 480
'TOT = Fix(TOT)
'num = 120: numd = 120
'If TOT = 1 Then For a = 0 To (Command1.Count - 1): Command1(a).left = 120: Command1(a).top = numd: numd = numd + 480: Next a: GoTo ed
'For a = 0 To (Command1.Count - 1)
'If num > ((TOT * 480) + 60) Then num = 120: numd = numd + 480
'Command1(a).left = num
'Command1(a).top = numd
'num = num + 480
'Next a
'ed:
'Call scroll_val(MainPic.Width, MainPic.Height)
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub
Private Sub grey_Click()
Dim TColA
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
    For i = 0 To xf
        For j = 0 To yf
            TColA = GetPixel(MainPic.hdc, i, j)
            r = TColA Mod 256
            g = (TColA \ 256) Mod 256
            b = TColA \ 256 \ 256
            r = Abs((g * b) / 256)
            g = Abs((b * r) / 256)
            b = Abs((r * g) / 256)
            SetPixelV MainPic.hdc, i, j, RGB(r, g, b)
        Next j
        progress = i * 100 \ (xf - 1)
        Call progressbar
    Next i
    MainPic.Refresh
End Sub

Private Sub Ham_Click()
Call use_hammer
End Sub

Private Sub Hlp_Click(Index As Integer)
    ShellExecute 0, "open", App.Path + "\HELP\Comic Pagemaker Help Index.htm", vbNullString, vbNullString, SW_MAXIMIZE
End Sub

Private Sub invert_Click()
Call Oldpic
Tempory.Width = MainPic.ScaleWidth
Tempory.Height = MainPic.ScaleHeight
Tempory.ScaleMode = vbPixels
Tempory.Cls
Call BitBlt(Tempory.hdc, 0, 0, MainPic.ScaleWidth, MainPic.ScaleHeight, MainPic.hdc, 0, 0, SRCINVERT)
MainPic.Picture = Tempory.Image
MainPic.Refresh
End Sub

Private Sub Lite_Click()
StatusBar1.Panels.Item(8) = "Lighting"
End Sub

Private Sub MainPic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim avr, avb, avg, cr, cg, cb

If ShowBalloonDialog = True Then
    Dim retval As String
    'used for InputBox
    Dim i As Integer
    'set i as integer
    i = Val(lblImgNum)
    'set i's value, it turns the string in
    'lblImgNum into an integer
    XPos = XPos * TileSize
    'XPos * TileSize because each tile will be
    'BitBlted in a different X value, so whereever
    'the mouse's X position is, it was divided
    'by 64, but in integer only (no decimals), and
    'then multiply it by the TileSize (64), because
    'remember, we need to BitBlt it to the right
    'spot, making it right next to each other
    YPos = YPos * TileSize
    'samething as the above description, except
    'replace X with Y
    Call PaintTile(XPos, YPos, i)
    'call the sub PaintTile
End If

Rx = ScaleX(X, vbPixels, vbCentimeters)
Ry = ScaleY(Y, vbPixels, vbCentimeters)
If StatusBar1.Panels.Item(8).Text = "Draw Line" And Button = 1 Then LineDr.X1 = X: LineDr.Y1 = Y: LineDr.X2 = X: LineDr.Y2 = Y: LineDr.Visible = True
If StatusBar1.Panels.Item(8).Text = "Pencil" And Button = 1 Then Call Oldpic: xl = X: yl = Y: MainPic.Line (X, Y)-(X, Y), lb
If StatusBar1.Panels.Item(8).Text = "Pencil" And Button = 2 Then Call Oldpic: xl = X: yl = Y: MainPic.Line (X, Y)-(X, Y), rb
If StatusBar1.Panels.Item(8).Text = "Draw Rectangle" And Button = 1 Then Call Oldpic: Shape.Shape = 0: Shape.Visible = True: Shape.left = X: Shape.top = Y: Shape.Width = 0: Shape.Height = 0: sx = X: sy = Y
If StatusBar1.Panels.Item(8).Text = "Draw Elipse" And Button = 1 Then Call Oldpic: Shape.Shape = 2: Shape.Visible = True: shapeleft = X: Shape.top = Y: Shape.Width = 0: Shape.Height = 0: sx = X: sy = Y
If StatusBar1.Panels.Item(8).Text = "Steal Colour" And Button = 1 Then lb = MainPic.Point(X, Y): Call show_cols
If StatusBar1.Panels.Item(8).Text = "Steal Colour" And Button = 2 Then rb = MainPic.Point(X, Y): Call show_cols
If StatusBar1.Panels.Item(8).Text = "Clone Tool" And Button = 1 And Shift = 4 Then cx = X: cy = Y: clone = True: cl = 1: CloneFrom.left = X: CloneFrom.top = Y
If StatusBar1.Panels.Item(8).Text = "Clone Tool" And Button = 1 Then Call Oldpic: CloneFrom.Visible = True
If StatusBar1.Panels.Item(8).Text = "Fill Region" Then GoTo Fill_Region
If StatusBar1.Panels.Item(8).Text = "Insert Text" And Button = 1 Then TextPic.Visible = True: TextPic.left = X + 1: TextPic.top = Y + 1
If StatusBar1.Panels.Item(8).Text = "Smudge" And Button = 1 Then GoTo smu
If StatusBar1.Panels.Item(8).Text = "Brush" And Button = 1 Then dwn = 1: Call draw_brush(X, Y)
If StatusBar1.Panels.Item(8).Text = "Polygon" And Button = 1 Then PolyX = X: PolyY = Y: For a = 0 To NumSides - 1: PolySide(a).Visible = True: PolySide(a).X1 = X: PolySide(a).X2 = X: PolySide(a).Y1 = Y: PolySide(a).Y2 = Y: Next
If StatusBar1.Panels.Item(8).Text = "Select Area" And Button = 1 Then CopyBox.Visible = True: CopyBox.left = X: CopyBox.top = Y: CopyBox.Width = 1: CopyBox.Height = 1: CopyX = X: CopyY = Y
If StatusBar1.Panels.Item(8).Text = "Paste" And Button = 1 Then Call Oldpic: PasteBox.left = X + 1: PasteBox.top = Y + 1: PasteBox.Visible = True
If StatusBar1.Panels.Item(8).Text = "Spray Can" And Button = 1 Then Call Oldpic: Call usespay(X, Y)
If StatusBar1.Panels.Item(8).Text = "Lighting" And Button = 1 Then Call Oldpic: Call lighting(X, Y): StatusBar1.Panels.Item(8) = ""
If StatusBar1.Panels.Item(8).Text = "Stiple" And Button = 1 Then MainPic.DrawWidth = 1: Call Oldpic: Call stiple(X, Y)
If StatusBar1.Panels.Item(8).Text = "Trace" And Button = 1 Then GoTo trce
GoTo nojmp
trce:
xtl = X: ytl = Y
Form7.Picture1.Line (X, Y)-(X, Y), RGB(0, 0, 0)
GoTo nojmp
smu:
Call smud(X, Y)
GoTo nojmp
Fill_Region:
Call Oldpic
MainPic.FillStyle = filarea
If Button = 1 Then MainPic.FillColor = lb
If Button = 2 Then MainPic.FillColor = rb
ExtFloodFill MainPic.hdc, X, Y, MainPic.Point(X, Y), 1
GoTo nojmp
nojmp:
End Sub

Private Sub PaintTile(SrcX As Long, SrcY As Long, ImageNum As Integer)
    ret = BitBlt(MainPic.hdc, SrcX, SrcY, TileSize, TileSize, Tiles(ImageNum).Src, 0, 0, SRCCOPY)
        'this BitBlts to the picMap, with the information
        'given when this sub was called
    MainPic.Refresh
        'refresh the picturebox, so that it could be
        'redrawn
        'sets the value of the coordinate printed,
        'this is essential because this is the information
        'that is being saved as a *.map file
End Sub

Private Sub MainPic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ShowBalloonDialog = True Then
    XPos = (X \ TileSize)
    'XPos is the current mouse's X value, but divided
    'by the TileSize, returning INTEGER only, that means
    'no decimals! That is what the sign "\" is for,
    'instead of "/" that you usually see
    YPos = (Y \ TileSize)
    'same as above, replace X with Y
    If Button = 1 Then
        'if the user has the mouse button down, then
        Call PaintTile(XPos * TileSize, YPos * TileSize, Val(lblImgNum))
        'call the sub PaintTile
    Else    'otherwise
        Exit Sub
        'exit current sub
    End If
End If
    
MainPic.MousePointer = 1
If StatusBar1.Panels.Item(8).Text = "Pencil" Then filtn = App.Path + "\GRAPHICS\Pencil.ico": MainPic.MousePointer = 99: MainPic.MouseIcon = LoadPicture(filtn)
If StatusBar1.Panels.Item(8).Text = "Steal Colour" Then filtn = App.Path + "\GRAPHICS\Pipet.ico": MainPic.MousePointer = 99: MainPic.MouseIcon = LoadPicture(filtn)
If StatusBar1.Panels.Item(8).Text = "Spray Can" Then filtn = App.Path + "\GRAPHICS\Spray.ico": MainPic.MousePointer = 99: MainPic.MouseIcon = LoadPicture(filtn)
If StatusBar1.Panels.Item(8).Text = "Fill Region" Then filtn = App.Path + "\GRAPHICS\Flood.ico": MainPic.MousePointer = 99: MainPic.MouseIcon = LoadPicture(filtn)
Rx = ScaleX(X, vbPixels, vbCentimeters)
Ry = ScaleY(Y, vbPixels, vbCentimeters)
siz = X & "," & Y
StatusBar1.Panels.Item(3) = siz
If StatusBar1.Panels.Item(8).Text = "Draw Line" And Button = 1 Then LineDr.X2 = X: LineDr.Y2 = Y
If StatusBar1.Panels.Item(8).Text = "Pencil" And Button = 1 Then MainPic.Line (xl, yl)-(X, Y), lb: xl = X: yl = Y
If StatusBar1.Panels.Item(8).Text = "Pencil" And Button = 2 Then MainPic.Line (xl, yl)-(X, Y), rb: xl = X: yl = Y
If StatusBar1.Panels.Item(8).Text = "Draw Rectangle" And Button = 1 Then GoTo rect
If StatusBar1.Panels.Item(8).Text = "Draw Elipse" And Button = 1 Then GoTo eclip
If StatusBar1.Panels.Item(8).Text = "Steal Colour" And Button = 1 Then lb = MainPic.Point(X, Y): Call show_cols
If StatusBar1.Panels.Item(8).Text = "Steal Colour" And Button = 2 Then rb = MainPic.Point(X, Y): Call show_cols
If StatusBar1.Panels.Item(8).Text = "Clone Tool" And Button = 1 And Shift = 4 Then cx = X: cy = Y: clone = True: GoTo dn
If StatusBar1.Panels.Item(8).Text = "Clone Tool" And Button = 1 And clone = True Then GoTo clon
If StatusBar1.Panels.Item(8).Text = "Insert Text" And Button = 1 Then TextPic.left = X + 1: TextPic.top = Y + 1
If StatusBar1.Panels.Item(8).Text = "Smudge" And Button = 1 Then GoTo smude
If StatusBar1.Panels.Item(8).Text = "Brush" And Button = 1 Then dwn = 0: Call draw_brush(X, Y)
If StatusBar1.Panels.Item(8).Text = "Polygon" And Button = 1 Then GoTo polydraw
If StatusBar1.Panels.Item(8).Text = "Select Area" And Button = 1 Then GoTo copyAreaSel
If StatusBar1.Panels.Item(8).Text = "Paste" And Button = 1 Then PasteBox.left = X + 1: PasteBox.top = Y + 1
If StatusBar1.Panels.Item(8).Text = "Spray Can" And Button = 1 Then Call usespay(X, Y)
If StatusBar1.Panels.Item(8).Text = "Stiple" And Button = 1 Then Call stiple(X, Y)
If StatusBar1.Panels.Item(8).Text = "Trace" And Button = 1 Then Form7.Picture1.Line (xtl, ytl)-(X, Y), RGB(0, 0, 0): xtl = X: ytl = Y
GoTo dn
copyAreaSel:
If CopyX > X Then CopyBox.left = X: CopyBox.Width = (CopyX - X)
If CopyX < X Then CopyBox.left = CopyX: CopyBox.Width = (X - CopyX)
If CopyY > Y Then CopyBox.top = Y: CopyBox.Height = (CopyY - Y)
If CopyY < Y Then CopyBox.top = CopyY: CopyBox.Height = (Y - CopyY)
GoTo dn
polydraw:
If X > PolyX Then XWID = X - PolyX
If X < PolyX Then XWID = PolyX - X
If Y > PolyY Then YWID = Y - PolyY
If Y < PolyY Then YWID = PolyY - Y
If PolyY < Y Then GoTo hhhh
If PolyY > Y Then GoTo hhhh1
GoTo dn
hhhh:
radis = Sqr((XWID * XWID) + (YWID * YWID))
num = (X - PolyX) / radis
AtAngle = Atn(-num / Sqr(-num * num + 1)) + 2 * Atn(1)
ADAN = (Pi * 2) / NumSides
For a = 0 To NumSides - 1
PolySide(a).X1 = PolyX + radis * Cos(AtAngle)
PolySide(a).X2 = PolyX + radis * Cos(AtAngle + ADAN)
PolySide(a).Y1 = PolyY + radis * Sin(AtAngle)
PolySide(a).Y2 = PolyY + radis * Sin(AtAngle + ADAN)
AtAngle = AtAngle + ADAN
Next a
GoTo dn
hhhh1:
radis = Sqr((XWID * XWID) + (YWID * YWID))
num = (X - PolyX) / radis
num = -num
AtAngle = Atn(-num / Sqr(-num * num + 1)) + 2 * Atn(1)
ADAN = (Pi * 2) / NumSides
For a = 0 To NumSides - 1
PolySide(a).X1 = PolyX + radis * Cos(AtAngle)
PolySide(a).X2 = PolyX + radis * Cos(AtAngle + ADAN)
PolySide(a).Y1 = PolyY + radis * Sin(AtAngle)
PolySide(a).Y2 = PolyY + radis * Sin(AtAngle + ADAN)
AtAngle = AtAngle + ADAN
Next a
GoTo dn
smude:
Call smud(X, Y)
GoTo dn
zom:
siz = Form2.Slider1.Value
wid = 200 / siz
TempZoom.Width = ScaleX(wid, vbPixels, vbTwips)
TempZoom.Height = ScaleY(wid, vbPixels, vbTwips)
TempZoom.Cls
TempZoom.Picture = LoadPicture()
Call BitBlt(TempZoom.hdc, 0, 0, wid, wid, MainPic.hdc, X - (wid / 2), Y - (wid / 2), SRCCOPY): TempZoom.Refresh
Form2.ZoomStretch.Picture = TempZoom.Image
Return
clon:
If cl = 1 Then xoff = X - cx: yoff = Y - cy: cl = 2
Call BitBlt(MainPic.hdc, X - 2, Y - 2, 4, 4, MainPic.hdc, X - xoff, Y - yoff, SRCCOPY): MainPic.Refresh
CloneFrom.top = Y - yoff
CloneFrom.left = X - xoff
GoTo dn
rect:
If sx > X Then Shape.left = X: Shape.Width = (sx - X)
If sx < X Then Shape.left = sx: Shape.Width = (X - sx)
If sy > Y Then Shape.top = Y: Shape.Height = (sy - Y)
If sy < Y Then Shape.top = sy: Shape.Height = (Y - sy)
GoTo dn
eclip:
If sx > X Then Shape.left = X: Shape.Width = (sx - X)
If sx < X Then Shape.left = sx: Shape.Width = (X - sx)
If sy > Y Then Shape.top = Y: Shape.Height = (sy - Y)
If sy < Y Then Shape.top = sy: Shape.Height = (Y - sy)
GoTo dn
dn:
If zoomactive = True Then GoSub zom
End Sub

Private Sub MainPic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Rx = ScaleX(X, vbPixels, vbCentimeters)
Ry = ScaleY(Y, vbPixels, vbCentimeters)
X1 = LineDr.X1
X2 = LineDr.X2
Y1 = LineDr.Y1
Y2 = LineDr.Y2
If StatusBar1.Panels.Item(8).Text = "Draw Line" And Button = 1 Then Call Oldpic: MainPic.Line (X1, Y1)-(X2, Y2), lb: LineDr.Visible = False:
If StatusBar1.Panels.Item(8).Text = "Draw Rectangle" And Button = 1 Then GoTo refin
If StatusBar1.Panels.Item(8).Text = "Draw Elipse" And Button = 1 Then GoTo ecli
If StatusBar1.Panels.Item(8).Text = "Insert Text" And Button = 1 Then GoTo dotext
If StatusBar1.Panels.Item(8).Text = "Clone Tool" And Button = 1 Then GoTo clonetolsel
If StatusBar1.Panels.Item(8).Text = "Polygon" And Button = 1 Then GoTo drpol
If StatusBar1.Panels.Item(8).Text = "Paste" And Button = 1 Then GoTo pstclip
GoTo don
pstclip:
Call BitBlt(MainPic.hdc, X + 1, Y + 1, PasteBox.Width, PasteBox.Height, PasteBox.hdc, 0, 0, SRCCOPY): MainPic.Refresh
PasteBox.Visible = False
StatusBar1.Panels.Item(8) = ""
GoTo don
drpol:
Call Oldpic
For a = 0 To NumSides - 1
PolySide(a).Visible = False
MainPic.Line (PolySide(a).X1, PolySide(a).Y1)-(PolySide(a).X2, PolySide(a).Y2), lb
Next a
GoTo don
clonetolsel:
CloneFrom.Visible = False
GoTo don
dotext:
Form3.Show
GoTo don
refin:
If sx > X Then w = sx - X Else w = X - sx
If sx > X Then sx = X
If sy > Y Then h = sy - Y Else h = Y - sy
If sy > Y Then sy = Y
w = w - 1
h = h - 1
If fillshape = True Then MainPic.FillStyle = 0: MainPic.Line (sx, sy)-(sx + w, sy + h), rb, BF: MainPic.Line (sx, sy)-(sx + w, sy + h), lb, B: MainPic.FillStyle = 1
If fillshape = False Then MainPic.FillStyle = 1: MainPic.Line (sx, sy)-(sx + w, sy + h), lb, B
Shape.Visible = False
GoTo don
ecli:
If sx > X Then w = sx - X Else w = X - sx
If sx > X Then sx = X
If sy > Y Then h = sy - Y Else h = Y - sy
If sy > Y Then sy = Y
w = w - 1
h = h - 1
sx = sx + (w / 2)
sy = sy + (h / 2)
If w > h Then radi = w / 2 Else radi = h / 2
asp = h / w
If radi < 0 Then Shape.Visible = False: GoTo don
If fillshape = True Then MainPic.FillColor = rb: MainPic.FillStyle = 0: MainPic.Circle (sx, sy), radi, lb, , , asp: MainPic.FillStyle = 1
If fillshape = False Then MainPic.FillStyle = 1: MainPic.Circle (sx, sy), radi, lb, , , asp
Shape.Visible = False
GoTo don
don:
End Sub

Private Sub MenuButton_Click()
    If MenuLabel.Caption = "Basic Toolbox" Then
        MenuLabel.Caption = "Advanced Toolbox"
        MenuButton.ToolTipText = "Click here to enable Basic Toolbox."
        Tools.Visible = False
        Advanced.Visible = True
    Else
        MenuLabel.Caption = "Basic Toolbox"
        MenuButton.ToolTipText = "Click here to enable Advanced Toolbox."
        Advanced.Visible = False
        Tools.Visible = True
    End If
End Sub

Private Sub NewImg_Click()
xx = InputBox("Enter Picture Width", "Width (Pixels)", 640)
If xx = "" Then Exit Sub
yy = InputBox("Enter Picture Height", "Height (Pixels)", 400)
If yy = "" Then Exit Sub
MainPic.Picture = LoadPicture()
MainPic.Cls
MainPic.Width = ScaleX(xx + 4, vbPixels, vbTwips)
MainPic.Height = ScaleY(yy + 4, vbPixels, vbTwips)
Call newpic(2)
End Sub

Private Sub OpFil_Click()
CommonDialog1.DialogTitle = "Open a Picture File"
CommonDialog1.Filter = "JPEG Files (*.JPG)|*.JPG|Bitmap Picture (*.BMP)|*.BMP|"
CommonDialog1.ShowOpen
Set MainPic = LoadPicture(CommonDialog1.Filename)
filname = CommonDialog1.Filename
dn = 0
For a = Len(filname) To 1 Step -1
If Mid$(filname, a, 1) = "\" And dn = 0 Then dn = a
Next a
PicName = Mid$(filname, dn + 1, (Len(filname) - dn))
StatusBar1.Panels.Item(2) = PicName
Form5.Width = Form1.MainPic.Width + 30: Call Drw_Rul
xs = ScaleX(MainPic.Width, vbTwips, vbPixels) - 4
ys = ScaleY(MainPic.Height, vbTwips, vbPixels) - 4
siz = xs & "," & ys
StatusBar1.Panels.Item(4) = siz
Call scroll_val(MainPic.Width, MainPic.Height)
Call clearundo
UpdateFileMenu (filname)
End Sub
Private Sub Option1_Click(Index As Integer)
MainPic.DrawStyle = Index
LineDr.BorderStyle = Index + 1
End Sub
Private Sub Option2_Click(Index As Integer)
If Index = 0 Then fillshape = True
If Index = 1 Then fillshape = False
End Sub
Private Sub Option3_Click(Index As Integer)
If Index = 6 Then filarea = 0
If Index = 5 Then filarea = 7
If Index = 4 Then filarea = 6
If Index = 3 Then filarea = 5
If Index = 2 Then filarea = 4
If Index = 1 Then filarea = 3
If Index = 0 Then filarea = 2
End Sub

Private Sub PastePic_Click()
StatusBar1.Panels.Item(8) = "Paste"
PasteBox.Picture = Clipboard.GetData()
End Sub

Private Sub picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.MousePointer = 7
If Y > Tools.Height And Y < Tools.Height + 500 Then Picture2.MousePointer = 7 Else Picture2.MousePointer = 1
If Button = 1 Then resi (Y)
End Sub
Private Sub resi(Y)
'Tools.Height = Y
'MoreTools.top = Tools.Height + 50
'MoreTools.Height = Picture2.Height - (Tools.Height + 60)
End Sub

Private Sub picTile_Click(Index As Integer)
    ShowBalloonDialog = True
    Form1.lblName = picTile(Index).Index
        'view the name of the current tile
    Form1.lblType = picTile(Index).Index
        'view the type of the current tile
    Form1.lblImgNum = picTile(Index).Index
    ret = BitBlt(picCurrent.hdc, 0, 0, TileSize, TileSize, Tiles(Index).Src, 0, 0, SRCCOPY)
    'BitBlt the tile that user wanted to paint with
    picCurrent.Refresh
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lb = Picture3.Point(X, Y)
If Button = 2 Then rb = Picture3.Point(X, Y)
Call show_cols
End Sub

Private Sub PicXScroll_Change()
MainPic.left = -PicXScroll.Value
End Sub
Private Sub PicXScroll_Scroll()
MainPic.left = -PicXScroll.Value
End Sub
Private Sub PicYScroll_Change()
MainPic.top = -PicYScroll.Value
End Sub
Private Sub PicYScroll_Scroll()
MainPic.top = -PicYScroll.Value
End Sub
Private Sub scroll_val(X, Y)
If MainPic.Width > 32767 Or MainPic.Height > 32767 Then GoTo toobig
maxx = PicYScroll.left
maxy = PicXScroll.top
If MainPic.Width < maxx Then maxx = MainPic.Width
PicYScroll.Enabled = True: PicYScroll.Visible = True
PicXScroll.Enabled = True: PicXScroll.Visible = True
yr = 0
If Y > maxy Then PicYScroll.Max = (Y - maxy): yr = 1 Else PicYScroll.Enabled = False: PicYScroll.Visible = False
If yr = 0 Then maxx = maxx + 280
If X > maxx Then PicXScroll.Max = (X - maxx) Else PicXScroll.Enabled = False: PicXScroll.Visible = False
PicXScroll.LargeChange = maxx / 25
PicXScroll.SmallChange = maxx / 100
PicYScroll.LargeChange = maxy / 25
PicYScroll.SmallChange = maxy / 10
Form1.Refresh
GoTo nnd
toobig:
ans = MsgBox("The Picture is too large to load", vbCritical, "Memory Error")
Call newpic(1)
nnd:
End Sub
Private Sub show_cols()
LBPIC.BackColor = lb
RBPIC.BackColor = rb
StatusBar1.Panels.Item(9).Picture = LBPIC.Image
StatusBar1.Panels.Item(11).Picture = RBPIC.Image
End Sub
Private Sub pixelatew_Click()
Dim pix As Long
ans = InputBox("Enter Value (1-10)", "Pixel Size", 5)
If ans = "" Then Exit Sub
pix = Val(ans)
If pix < 1 Or pix > 10 Then Exit Sub
Call Oldpic
Call pixelate(pix)
End Sub

Private Sub PrintPicNow_Click()
Call PrintPicture
End Sub
Private Sub Proview_Click()
If Proview.Checked = True Then Proview.Checked = False: Form1.StatusBar1.Panels.Item(1).Picture = Form1.Progpic(1).Image Else Proview.Checked = True: Form1.StatusBar1.Panels.Item(1).Picture = Form1.Progpic(0).Image
End Sub
Private Sub PrSet_Click()
CommonDialog1.ShowPrinter
End Sub

Private Sub Redolast_Click()
Call Oldpic
MainPic.Picture = RedoBox.Image
Redolast.Enabled = False
Command1(17).Enabled = False
End Sub
Private Sub Rep_col_Click()
Form4.Show
End Sub

Private Sub RepIndCol_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
For X = 1 To xf - 2
For Y = 1 To yf - 2
If GetPixel(MainPic.hdc, X, Y) = lb Then SetPixelV MainPic.hdc, X, Y, rb
Next Y
progress = (100 / (xf - 2)) * X
Call progressbar
Next X
MainPic.Refresh
End Sub

Private Sub Rilers_Click()
If Rilers.Checked = False Then Form5.Show: Rilers.Checked = True: GoTo hhjk
If Rilers.Checked = True Then Unload Form5: Rilers.Checked = False: GoTo hhjk
hhjk:
End Sub

Private Sub Rotate_Click()
Call Oldpic
    Tempory.Cls
    Call bmp_rotate(Pi / 6)
End Sub

Private Sub Rulers_Click()
If Rulers.Checked = False Then Form6.Show: Rulers.Checked = True: GoTo hhjk2
If Rulers.Checked = True Then Unload Form6: Rulers.Checked = False: GoTo hhjk2
hhjk2:
End Sub

Private Sub SavFil_Click()
CommonDialog1.DialogTitle = "Save Current Picture File"
CommonDialog1.Filter = "Bitmap Files (*.BMP)|*.BMP|JPEG Files (*.JPG)|*.JPG|"
CommonDialog1.ShowSave
If CommonDialog1.Filename = "" Then GoTo dds
If FileExist(CommonDialog1.Filename) = True Then ans = MsgBox("That File Already Exists, Overwright", vbYesNo, "Are you Sure")
If ans = vbNo Then GoTo dds
SavePicture MainPic.Image, CommonDialog1.Filename
dds:
End Sub
Private Sub SelSrc_Click()
TWAIN_SelectImageSource (Me.hwnd)
End Sub
Private Sub Sharpen_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
   Call loading(i, j)
    For i = 1 To yf - 2
        For j = 1 To xf - 2
            Red = ImageArray(0, i, j) + 0.5 * (ImageArray(0, i, j) - ImageArray(0, i - 1, j - 1))
            Green = ImageArray(1, i, j) + 0.5 * (ImageArray(1, i, j) - ImageArray(1, i - 1, j - 1))
            Blue = ImageArray(2, i, j) + 0.5 * (ImageArray(2, i, j) - ImageArray(2, i - 1, j - 1))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            SetPixelV MainPic.hdc, j, i, RGB(Red, Green, Blue)
        Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub
Private Sub Sharpen_more_Click()
Call Oldpic
xf = MainPic.ScaleWidth
yf = MainPic.ScaleHeight
Call loading(i, j)
    For i = 1 To yf - 2
        For j = 1 To xf - 2
            Red = ImageArray(0, i, j) + 0.8 * (ImageArray(0, i, j) - ImageArray(0, i - 1, j - 1))
            Green = ImageArray(1, i, j) + 0.8 * (ImageArray(1, i, j) - ImageArray(1, i - 1, j - 1))
            Blue = ImageArray(2, i, j) + 0.8 * (ImageArray(2, i, j) - ImageArray(2, i - 1, j - 1))
            If Red > 255 Then Red = 255
            If Red < 0 Then Red = 0
            If Green > 255 Then Green = 255
            If Green < 0 Then Green = 0
            If Blue > 255 Then Blue = 255
            If Blue < 0 Then Blue = 0
            SetPixelV MainPic.hdc, j, i, RGB(Red, Green, Blue)
        Next
        progress = 50 + (i * 100 / (yf - 1)) / 2
        Call progressbar
    Next
End Sub
Private Sub Slider1_Change()
SampleLine.Cls
SampleLine.DrawWidth = Slider1.Value
SampleLine.Line (0, SampleLine.Height / 2)-(7000, SampleLine.Height / 2), RGB(0, 0, 0)
MainPic.DrawWidth = Slider1.Value
End Sub
Private Sub Slider1_Click()
SampleLine.Cls
SampleLine.DrawWidth = Slider1.Value
SampleLine.Line (0, SampleLine.Height / 2)-(7000, SampleLine.Height / 2), RGB(0, 0, 0)
MainPic.DrawWidth = Slider1.Value
End Sub
Private Sub Slider1_Scroll()
SampleLine.Cls
SampleLine.DrawWidth = Slider1.Value
SampleLine.Line (0, SampleLine.Height / 2)-(7000, SampleLine.Height / 2), RGB(0, 0, 0)
MainPic.DrawWidth = Slider1.Value
End Sub
Private Sub StatusBar1_PanelClick(ByVal Panel As ComctlLib.Panel)
If Panel.key = "CHBUT" Then l = rb: rb = lb: lb = l
If Panel.key = "LBBUT" Or Panel.key = "RBBUT" Then CommonDialog1.ShowColor
If Panel.key = "LBBUT" Then lb = CommonDialog1.Color
If Panel.key = "RBBUT" Then rb = CommonDialog1.Color
Call show_cols
Form1.TextPic.ForeColor = lb
End Sub

Private Sub Timer1_Timer()
If CopyBox.Visible = True Then CutPic.Enabled = True Else CutPic.Enabled = False
If CopyBox.Visible = True Then CopyPic.Enabled = True Else CopyPic.Enabled = False
If Clipboard.GetFormat(2) = True Then PastePic.Enabled = True Else PastePic.Enabled = False
If Rilers.Checked = False Then GoTo norl
Form5.Line2.X1 = Rx: Form5.Line2.X2 = Rx
norl:
If Rulers.Checked = False Then GoTo norl1
Form6.Line2.Y1 = Ry: Form6.Line2.Y2 = Ry
norl1:
End Sub

Private Sub TitleBar_Click()

End Sub

Private Sub UndoLast_Click()
Redolast.Enabled = True
Command1(17).Enabled = True
RedoBox.Picture = MainPic.Image
MainPic.Picture = UndoPicBox(0).Image
For a = 1 To 4: UndoPicBox(a - 1).Picture = UndoPicBox(a).Image
Next a
End Sub
Private Sub clearundo()
For a = 0 To 4: UndoPicBox(a).Picture = MainPic.Image: Next
End Sub
Private Sub screencapture(left As Long, top As Long, right As Long, bottom As Long)
Dim capWidth As Long, capHeight As Long
capWidth = right - left
capHeight = bottom - top
srcdc = CreateDC("DISPLAY", 0, 0, 0)
destdc = CreateCompatibleDC(srcdc)
hbmp = CreateCompatibleBitmap(srcdc, capWidth, capHeight)
SelectObject destdc, hbmp
BitBlt destdc, 0, 0, capWidth, capHeight, srcdc, left, top, &HCC0020
OpenClipboard Screen.ActiveForm.hwnd
EmptyClipboard
SetClipboardData 2, hbmp
CloseClipBoard
DeleteObject hbmp
DeleteDC destdc
ReleaseDC 0, srcdc
Command1(1).Enabled = True
End Sub
Private Sub newpic(PSS)
If PSS = 1 Then GoTo NOQUEST
If PSS = 2 Then GoTo newsize
ans = MsgBox("Are You Sure", vbYesNo, "User Reminder")
If ans = vbNo Then GoTo notnow
NOQUEST:
MainPic.Picture = LoadPicture()
MainPic.Cls
'MainPic.Width = ScaleX(644, vbPixels, vbTwips)
MainPic.Height = ScaleY(404, vbPixels, vbTwips)
newsize:
Call resiz(11400)
Call resi(3400)
Call scroll_val(MainPic.Width, MainPic.Height)
PicName = "Untitled.BMP"
StatusBar1.Panels.Item(2) = PicName
xs = ScaleX(MainPic.Width, vbTwips, vbPixels) - 4
ys = ScaleY(MainPic.Height, vbTwips, vbPixels) - 4
siz = xs & "," & ys
StatusBar1.Panels.Item(4) = siz
lb = RGB(0, 0, 0)
rb = RGB(255, 255, 255)
Call show_cols
SampleLine.Cls
SampleLine.DrawWidth = Slider1.Value
SampleLine.Line (0, SampleLine.Height / 2)-(7000, SampleLine.Height / 2), RGB(0, 0, 0)
MainPic.DrawWidth = Slider1.Value
Call clearundo
notnow:
End Sub
Public Sub smud(X, Y)
col = MainPic.Point(X, Y)
MainPic.Line (X - 1, Y)-(X + 1, Y), col
MainPic.Line (X, Y - 1)-(X, Y + 1), col
End Sub
Private Sub draw_brush(X, Y)
If bushtype < 0 And BrushType > 11 Then Exit Sub
If dwn = 1 Then Call Oldpic
If BrushType = 0 Then MainPic.Line (X - 5, Y + 5)-(X + 5, Y - 5), lb
If BrushType = 1 Then MainPic.Line (X + 5, Y + 5)-(X - 5, Y - 5), lb
If BrushType = 2 Then MainPic.Line (X, Y + 5)-(X, Y - 5), lb: MainPic.Line (X + 5, Y)-(X - 5, Y), lb
If BrushType = 3 Then MainPic.Line (X, Y + 5)-(X, Y - 5), lb: MainPic.Line (X + 5, Y)-(X - 5, Y), lb: MainPic.Line (X - 5, Y + 5)-(X + 5, Y - 5), lb: MainPic.Line (X + 5, Y + 5)-(X - 5, Y - 5), lb
If BrushType = 4 Then For a = 0 To 5 Step 0.5: MainPic.Circle (X, Y), a, lb: Next a
If BrushType = 5 Then MainPic.Line (X - 5, Y - 5)-(X + 5, Y + 5), lb, BF
If BrushType = 6 Then For a = -5 To 5: MainPic.Line (X + a, Y + 5)-(X, Y - 5), lb: Next
If BrushType = 7 Then For a = -5 To 5: MainPic.Line (X + a, Y - 5)-(X, Y + 5), lb: Next
If BrushType = 8 Then For a = -5 To 5: MainPic.Line (X - 5, Y)-(X + 5, Y + a), lb: Next
If BrushType = 9 Then For a = -5 To 5: MainPic.Line (X - 5, Y + a)-(X + 5, Y), lb: Next
If BrushType = 10 Then MainPic.Line (X - 5, Y + 5)-(X + 5, Y - 2), lb: MainPic.Line (X + 5, Y - 2)-(X - 5, Y - 2), lb: MainPic.Line (X - 5, Y - 2)-(X + 5, Y + 5), lb: MainPic.Line (X + 5, Y + 5)-(X, Y - 5), lb: MainPic.Line (X, Y - 5)-(X - 5, Y + 5), lb
If BrushType = 11 Then For a = X - 5 To X + 5: MainPic.Line (X, Y - 5)-(a, Y), lb: MainPic.Line (X, Y + 5)-(a, Y), lb: Next a: MainPic.Line (X - 5, Y)-(X + 6, Y), lb
End Sub
Sub loading(i, j)
    xf = MainPic.ScaleWidth
    yf = MainPic.ScaleHeight
    Dim Color As Long
    For i = 0 To yf - 1
        For j = 0 To xf - 1
            pixel& = Form1.MainPic.Point(j, i)
            Red = pixel& Mod 256
            Green = ((pixel& And &HFF00) / 256&) Mod 256&
            Blue = (pixel& And &HFF0000) / 65536
            ImageArray(0, i, j) = Red
            ImageArray(1, i, j) = Green
            ImageArray(2, i, j) = Blue
        Next
        progress = Abs(i * 100 / (yf - 1) / 2)
        Call progressbar
    Next
End Sub
Private Sub pixelate(size As Long)
f = size: f2 = f / 2 - 1
All = (MainPic.ScaleWidth - f) * (MainPic.ScaleHeight - f) / f / f
For i = f2 To MainPic.ScaleWidth - f2 Step f
For j = f2 To MainPic.ScaleHeight - f2 Step f
r = 0: g = 0: b = 0
For k = -f2 To f2 Step f2 / 2: For l = -f2 To f2 Step f2 / 2
r = r + TakeRGB(MainPic.Point(i + k, j + l), 0)
g = g + TakeRGB(MainPic.Point(i + k, j + l), 1)
b = b + TakeRGB(MainPic.Point(i + k, j + l), 2)
Next l, k
MainPic.Line (i - f2, j - f2)-(i + f2, j + f2), RGB(r / 25, g / 25, b / 25), BF
h = h + 1
If h > All Then progress = 100 Else progress = h / All * 100
Call progressbar
Next j
Next i
End Sub
Function TakeRGB(Colors As Long, Index As Integer) As Long
IndexColor = Colors
Red = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Red) / 256
Green = IndexColor - Int(IndexColor / 256) * 256: IndexColor = (IndexColor - Green) / 256
Blue = IndexColor
If Index = 0 Then TakeRGB = Red
If Index = 1 Then TakeRGB = Green
If Index = 2 Then TakeRGB = Blue
End Function
Sub bmp_rotate(theta)
    Tempory.Picture = MainPic.Picture
    Tempory.Picture = LoadPicture()
        Dim c1x As Integer, c1y As Integer
    Dim c2x As Integer, c2y As Integer
    Dim a As Single
    Dim p1x As Integer, p1y As Integer
    Dim p2x As Integer, p2y As Integer
    Dim n As Integer, r As Integer
    c1x = MainPic.ScaleWidth \ 2
    c1y = MainPic.ScaleHeight \ 2
    c2x = Tempory.ScaleWidth \ 2
    c2y = Tempory.ScaleHeight \ 2
    If c2x < c2y Then n = c2y Else n = c2x
    n = n - 1
    For p2x = 0 To n
        For p2y = 0 To n
            If p2x = 0 Then a = Pi / 2 Else a = Atn(p2y / p2x)
            r = Sqr(1& * p2x * p2x + 1& * p2y * p2y)
            p1x = r * Cos(a + theta)
            p1y = r * Sin(a + theta)
            c0& = GetPixel(MainPic.hdc, c1x + p1x, c1y + p1y)
            c1& = GetPixel(MainPic.hdc, c1x - p1x, c1y - p1y)
            c2& = GetPixel(MainPic.hdc, c1x + p1y, c1y - p1x)
            c3& = GetPixel(MainPic.hdc, c1x - p1y, c1y + p1x)
            If c0& <> -1 Then xret& = SetPixelV(Tempory.hdc, c2x + p2x, c2y + p2y, c0&)
            If c1& <> -1 Then xret& = SetPixelV(Tempory.hdc, c2x - p2x, c2y - p2y, c1&)
            If c2& <> -1 Then xret& = SetPixelV(Tempory.hdc, c2x + p2y, c2y - p2x, c2&)
            If c3& <> -1 Then xret& = SetPixelV(Tempory.hdc, c2x - p2y, c2y + p2x, c3&)
        Next
        progress = p2x \ n
        Call progressbar
    Next
tmpfil = App.Path & "\temp.bmp"
SavePicture Tempory.Image, tmpfil
MainPic = LoadPicture()
Set MainPic = LoadPicture(tmpfil)
End Sub
Private Sub WAVE_Click()
Call Oldpic
    Dim i As Long, j As Long
    Dim sw As Long, sh As Long
    Dim coli() As Long, posy() As Double
    sw = MainPic.ScaleWidth
    sh = MainPic.ScaleHeight
    ReDim coli(sw, sh)
    ReDim posy(sw, sh)
    For i = 0 To sw
        For j = 0 To sh
            coli(i, j) = GetPixel(MainPic.hdc, i, j)
            posy(i, j) = Sin(i) * 6 + (j - 3)
        Next j
        progress = (i * 100 \ (sw - 1)) \ 2
        Call progressbar
    Next i
    For i = 0 To sw
        For j = 0 To sh
            MainPic.PSet (i, posy(i, j)), coli(i, j)
        Next j
        progress = 50 + (i * 100 \ (sw - 1)) \ 2
        Call progressbar
    Next i
    MainPic.Refresh
End Sub
Private Sub Drw_Rul()
Form5.Cls
For a = 0 To 60
Form5.Line (a, 0.1)-(a, 0.4), RGB(0, 0, 0)
If a < 10 Then ofst = 0.13 Else ofst = 0.2
Form5.Line (a - ofst, 0.5)-(a - ofst, 0.5), RGB(0, 0, 0)
Form5.Print a
Next a
For a = 0 To 60 Step 0.1
Form5.Line (a, 0.15)-(a, 0.35), RGB(0, 0, 0)
Next a
ng = ScaleX(4, vbPixels, vbCentimeters)
Form5.Caption = "X Ruler - " & (ScaleX(Form1.MainPic.Width, vbTwips, vbCentimeters) - ng) & "cm"
End Sub
Private Sub putToClipBoard(inp)
ClipDataPic.Picture = LoadPicture()
ClipDataPic.Cls
ClipDataPic.Width = ScaleX(CopyBox.Width, vbPixels, vbTwips)
ClipDataPic.Height = ScaleY(CopyBox.Height, vbPixels, vbTwips)
Call BitBlt(ClipDataPic.hdc, 0, 0, CopyBox.Width, CopyBox.Height, MainPic.hdc, CopyBox.left, CopyBox.top, SRCCOPY)
ClipDataPic.Refresh
If inp = 1 Then Call Oldpic: MainPic.Line (CopyBox.left, CopyBox.top)-((CopyBox.left + CopyBox.Width) - 1, (CopyBox.top + CopyBox.Height) - 1), rb, BF: MainPic.Refresh
CopyBox.Visible = False
Clipboard.Clear
Clipboard.SetData ClipDataPic.Image, vbCFBitmap
End Sub
Private Sub usespay(X, Y)
For a = 1 To 4
MainPic.DrawWidth = 2
Let xx = (Rnd(1) * (Slider1.Value + 5)) + 1
Let yy = (Rnd(1) * (Slider1.Value + 5)) + 1
Let an = (Rnd(1) * 6.28) + 1
MainPic.Line (X + xx * Cos(an), Y + yy * Sin(an))-(X + xx * Cos(an), Y + yy * Sin(an)), lb
Next a
MainPic.DrawWidth = Slider1.Value
End Sub
Private Sub circpix(size)
f = size: f2 = f / 2 - 1
All = (MainPic.ScaleWidth - f) * (MainPic.ScaleHeight - f) / f / f
For i = f2 To MainPic.ScaleWidth - f2 Step f
For j = f2 To MainPic.ScaleHeight - f2 Step f
r = 0: g = 0: b = 0
For k = -f2 To f2 Step f2 / 2: For l = -f2 To f2 Step f2 / 2
r = r + TakeRGB(MainPic.Point(i + k, j + l), 0)
g = g + TakeRGB(MainPic.Point(i + k, j + l), 1)
b = b + TakeRGB(MainPic.Point(i + k, j + l), 2)
Next l, k
MainPic.Circle (i - f2, j - f2), f, RGB(r / 25, g / 25, b / 25), BF
h = h + 1
If h > All Then progress = 100 Else progress = h / All * 100
Call progressbar
Next j
Next i
End Sub
Private Sub lighting(xp, yp)
Dim Brightness, diffcol As Single
Dim NewColor As Long
Dim X, Y, raditxt As Integer
Dim r, g, b As Integer
txt = InputBox("Enter Brightness Level (0-100)", "Brightness", 50)
txtbrightness = Val(txt)
If txtbrightness < 0 Or txtbrightness > 100 Then Exit Sub
txtbrightness = txtbrightness
txt = InputBox("Enter Radius ", "Light Radius", 50)
raditxt = Val(txt)
Brightness = txtbrightness / 100
diffcol = Brightness / raditxt
Brightness = Brightness + 1
progress = 0: Call progressbar
For radin = 0 To raditxt
For ang = 0 To (2 * Pi) Step 0.01
X = xp + radin * Cos(ang)
Y = yp + radin * Sin(ang)
NewColor = GetPixel(MainPic.hdc, X, Y)
r = (NewColor Mod 256)
b = (Int(NewColor / 65536))
g = ((NewColor - (b * 65536) - r) / 256)
r = r * Brightness
b = b * Brightness
g = g * Brightness
If r > 255 Then r = 255
If r < 0 Then r = 0
If b > 255 Then b = 255
If b < 0 Then b = 0
If g > 255 Then g = 255
If g < 0 Then g = 0
SetPixelV MainPic.hdc, X, Y, RGB(r, g, b)
Next ang
Brightness = Brightness - diffcol
progress = (100 / raditxt) * radin
Call progressbar
Next radin
MainPic.Refresh
End Sub
Private Sub PrintPicture()
Printer.CurrentX = 0
Printer.CurrentY = 0
Printer.PaintPicture MainPic.Picture, 0, 0
Printer.EndDoc
End Sub
Private Sub stiple(X, Y)
MainPic.DrawWidth = 1
Let xxe = Int(Rnd(1) * (Slider1.Value + 10)) + 1
Let yye = Int(Rnd(1) * (Slider1.Value + 10)) + 1
Let r = Int(Rnd(1) * 2) + 1
If r = 1 Then yye = -yye
Let r = Int(Rnd(1) * 2) + 1
If r = 1 Then xxe = -xxe
MainPic.Line (X, Y)-(X + xxe, Y + yye), lb
MainPic.DrawWidth = Slider1.Value
End Sub
Private Sub use_hammer()
Dim Brightness, diffcol As Single
Dim NewColor As Long
Dim X, Y, raditxt As Integer
Dim r, g, b As Integer
progress = 0
For lp = 1 To 2
For yp = 5 To MainPic.ScaleHeight Step 10
For xp = 5 To MainPic.ScaleWidth Step 10
Brightness = 0.995
For radin = 1 To 5
For ang = 0 To (2 * Pi) Step 0.1
X = xp + radin * Cos(ang)
Y = yp + radin * Sin(ang)
NewColor = GetPixel(MainPic.hdc, X, Y)
r = (NewColor Mod 256)
b = (Int(NewColor / 65536))
g = ((NewColor - (b * 65536) - r) / 256)
r = r * Brightness
b = b * Brightness
g = g * Brightness
If r > 255 Then r = 255
If r < 0 Then r = 0
If b > 255 Then b = 255
If b < 0 Then b = 0
If g > 255 Then g = 255
If g < 0 Then g = 0
SetPixelV MainPic.hdc, X, Y, RGB(r, g, b)
Next ang
Brightness = Brightness + 0.001
Next radin
Next xp
If lp = 1 Then progress = ((yp / MainPic.ScaleHeight) * 100) / 2
If lp = 2 Then progress = (((yp / MainPic.ScaleHeight) * 100) / 2) + 50
Call progressbar
Next yp
Next lp
progress = 100
Call progressbar
MainPic.Refresh
End Sub
Private Sub mnuRecentFile_Click(Index As Integer)
On Error GoTo rrtr
openfile = mnuRecentFile(Index).Caption
Set MainPic = LoadPicture(openfile)
filname = openfile
dn = 0
For a = Len(filname) To 1 Step -1
If Mid$(filname, a, 1) = "\" And dn = 0 Then dn = a
Next a
PicName = Mid$(filname, dn + 1, (Len(filname) - dn))
StatusBar1.Panels.Item(2) = PicName
Form5.Width = Form1.MainPic.Width + 30: Call Drw_Rul
xs = ScaleX(MainPic.Width, vbTwips, vbPixels) - 4
ys = ScaleY(MainPic.Height, vbTwips, vbPixels) - 4
siz = xs & "," & ys
StatusBar1.Panels.Item(4) = siz
Call scroll_val(MainPic.Width, MainPic.Height)
Call clearundo
GetRecentFiles
GoTo noer
rrtr:
ans = MsgBox("That File no longer exits", vbCritical, "File Error")
noer:
End Sub

