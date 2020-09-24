VERSION 5.00
Begin VB.Form Form9 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   1245
   ClientTop       =   1440
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5475
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "Splash.frx":000C
      Top             =   990
      Width           =   480
   End
   Begin VB.Label lblProductName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Program Title"
      BeginProperty Font 
         Name            =   "Jokerman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   555
      Left            =   960
      TabIndex        =   3
      Top             =   960
      Width           =   4125
   End
   Begin VB.Label lblCopyright 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright c (Your Company)"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   0
      TabIndex        =   2
      Top             =   630
      Width           =   5445
   End
   Begin VB.Label lblWarning 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This program is protected by national and international copyright laws as described in Help About."
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   2145
      Width           =   5445
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPlatformAndVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "for Win X Version x.x"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      TabIndex        =   0
      Top             =   1650
      Width           =   5445
   End
   Begin VB.Image BackgroundPicture 
      Height          =   3165
      Left            =   15
      Picture         =   "Splash.frx":044E
      Stretch         =   -1  'True
      Top             =   15
      Width           =   5445
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Stores the number of seconds elapsed since midnight to determine the display time of the Splash window.
Private msngSplashDisplayStartTime As Single

'Platform the application runs on (e.g. "Win 95").
Public Platform As String

Private Sub Form_Activate()
    On Error GoTo HandleErrors
    Refresh
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Private Sub Form_Load()
    On Error GoTo HandleErrors
    'Change the Screens mouse pointer for this application as we dont want the user thinking that
    'they can start working on another form while the Splash form is still displayed.
    Screen.MousePointer = vbHourglass
    '------------------------------------Assign propertys for the Splash form---------------------------------------------
    lblPlatformAndVersion = "For " & Platform & " Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
    lblCopyright = "Copyright " & Chr(169) & " " & Year(Now()) & " " & "Your Company"
    'Include the Applications Title in the splash forms caption so that when displayed at run-time the
    'applications title will appear in the Windows Task Bar. This is especially important with applications
    'with lengthy startup code because the user needs to be informed that the application has indeed started.
    'Note that the Splash forms 'ShowInTaskbar' property must be set to TRUE at design-time to achieve this.
    Caption = App.Title
    'Move the background to the size of the form, minus a border space, so that the background
    'image displays correctly in all screen resolutions.
    BackgroundPicture.Move 15, 15, Me.ScaleWidth - 30, Me.ScaleHeight - 30
    '----------------------------------------------------------------------------------------------------------------------
    'Assign start time of the display of the Splash window.
    msngSplashDisplayStartTime = Timer
ExitHandleErrors:
  Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error GoTo HandleErrors
    'Num. sec. the Splash Window is displayed.
    Const cintDisplayTimeSeconds As Integer = 3
    'Loop until the Display Time has elpased - if the applications loading time took longer than
    'the display time it will not enter this loop.
    Do Until (Timer - msngSplashDisplayStartTime) > cintDisplayTimeSeconds
    Loop
    Screen.MousePointer = vbNormal
ExitHandleErrors:
    Exit Sub
HandleErrors:
  MsgBox Err.Description & " (" & Err.Number & ")", vbCritical, App.Title & " Error"
  Resume ExitHandleErrors
End Sub

