VERSION 5.00
Begin VB.Form Form8 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Comic Pagmaker"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timMain 
      Interval        =   100
      Left            =   1200
      Top             =   2700
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   450
      Left            =   3210
      TabIndex        =   14
      Top             =   5100
      Width           =   1215
   End
   Begin VB.PictureBox picOut 
      BackColor       =   &H00FFFFFF&
      Height          =   4905
      Left            =   30
      ScaleHeight     =   4845
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   60
      Width           =   4395
      Begin VB.PictureBox picUp 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   -15
         ScaleHeight     =   1035
         ScaleWidth      =   4410
         TabIndex        =   11
         Top             =   -30
         Width           =   4410
         Begin VB.Line Line3 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            X1              =   75
            X2              =   4230
            Y1              =   900
            Y2              =   900
         End
         Begin VB.Label Label2 
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "For Comic Strips, Editorial and Graphics Layout"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   120
            TabIndex        =   13
            Top             =   615
            Width           =   4185
         End
         Begin VB.Label lblMain 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Comic Pagemaker"
            BeginProperty Font 
               Name            =   "Jokerman"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   795
            TabIndex        =   12
            Top             =   105
            Width           =   3090
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            X1              =   90
            X2              =   4245
            Y1              =   570
            Y2              =   570
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   120
            Picture         =   "About.frx":000C
            ToolTipText     =   "Don't double click me please :)"
            Top             =   45
            Width           =   480
         End
      End
      Begin VB.PictureBox picIn 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7275
         Left            =   135
         ScaleHeight     =   7275
         ScaleWidth      =   4110
         TabIndex        =   1
         Top             =   1170
         Width           =   4110
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   4
            X1              =   60
            X2              =   1725
            Y1              =   4935
            Y2              =   4935
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Disclaimer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   4
            Left            =   165
            TabIndex        =   16
            Top             =   4665
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"About.frx":044E
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   885
            Index           =   5
            Left            =   360
            TabIndex        =   15
            Top             =   5040
            Width           =   3495
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   3
            X1              =   75
            X2              =   1740
            Y1              =   3435
            Y2              =   3435
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"About.frx":050D
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   840
            Index           =   4
            Left            =   375
            TabIndex        =   10
            Top             =   3585
            Width           =   3495
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   3
            Left            =   210
            TabIndex        =   9
            Top             =   3180
            Width           =   1095
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   2
            X1              =   30
            X2              =   1695
            Y1              =   6450
            Y2              =   6450
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"About.frx":05A4
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   690
            Index           =   3
            Left            =   330
            TabIndex        =   8
            Top             =   6600
            Width           =   3495
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Redistribution"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   105
            TabIndex        =   7
            Top             =   6180
            Width           =   1320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   1
            X1              =   105
            X2              =   1770
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"About.frx":062D
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   1305
            Index           =   2
            Left            =   390
            TabIndex        =   6
            Top             =   1755
            Width           =   3495
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            Index           =   0
            X1              =   60
            X2              =   1725
            Y1              =   330
            Y2              =   330
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Walter A. Narvasa of WANCOM SYSTEMS"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   480
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   735
            Width           =   3375
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "jawoltze@edsmail.com.ph"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   2175
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Developed By "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   60
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' Hey sir, Kindly rate this code, if you like it.
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Icon Hunter fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in Icon Hunter. Contact me for
' additional help/suggestions
'
' I recently inveted a technology for streaming audio, and is
' now looking promoters/investors to invest in a web-phone network
' project.
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================

Dim CurScheme As Integer
Dim EasterFlag As Boolean

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
CurScheme = 2
EasterFlag = False
timMain.Interval = 30
picIn.top = picOut.ScaleHeight + 20
End Sub

Private Sub Image1_DblClick()
If EasterFlag = False Then
    MsgBox "OK, you cracked one easter egg, one more is there to crack", vbInformation + vbOKOnly, "Icon Hunter"
    Image1.ToolTipText = "Don't right click me please"
    
End If

EasterFlag = True
    CurScheme = CurScheme + 1
    If CurScheme = 5 Then CurScheme = 1
    ChangeState CurScheme
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If EasterFlag = True Then
        frmAuthor.Show vbModal
    End If
End If
End Sub

Private Sub Label1_Click(Index As Integer)
picIn_Click
End Sub

Private Sub Label2_Click()
picIn_Click
End Sub


Private Sub lblCap_Click(Index As Integer)
picIn_Click
End Sub


Private Sub lblMain_Click()
picIn_Click
End Sub


Private Sub picIn_Click()
timMain.Enabled = Not timMain.Enabled
End Sub

Private Sub timMain_Timer()

picIn.top = picIn.top - 10
If picIn.top + picIn.Height < picUp.Height + picUp.top Then picIn.top = picOut.ScaleHeight + 20


If EasterFlag = False Then Exit Sub

If picIn.top = picOut.ScaleHeight + 20 Then
    ChangeState CurScheme
    CurScheme = CurScheme + 1
    If CurScheme = 5 Then CurScheme = 1
End If

End Sub


Sub ChangeState(State As Integer)

Select Case State
    Case 1
        Dim myC As Control
        
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbBlack
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbGreen
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbRed
            End If
        Next myC
     Case 2
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbWhite
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbBlack
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbBlack
            End If
        Next myC

     Case 3
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbBlack
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbRed
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbGreen
            End If
        Next myC

     Case 4
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbBlack
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbWhite
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbWhite
            End If
        Next myC

End Select

End Sub
