VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About CD Quick View.OCX"
   ClientHeight    =   4350
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5010
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3002.447
   ScaleMode       =   0  'User
   ScaleWidth      =   4704.649
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   3600
      Width           =   660
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "&Help"
      Height          =   345
      Left            =   4320
      TabIndex        =   2
      Top             =   3120
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "Register to unlock view mode for hard disks and to be able to save thumbnails to disk. Cost £5.00"
      Height          =   615
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":030A
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   4095
   End
   Begin VB.Label lblMailTo 
      Caption         =   "derek.hall@virgin.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label lblEMail 
      Caption         =   "E-Mail Author:"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblCompanyRegistration 
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   4620.134
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title: CD Quick View.OCX"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   3165
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4620.134
      Y1              =   1325.218
      Y2              =   1325.218
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:"
      Height          =   225
      Left            =   840
      TabIndex        =   5
      Top             =   360
      Width           =   3285
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Warning...This version of CD Viewer is not registered  contact author:"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   4095
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  
  lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
  lblTitle.Caption = "Application Title:" & "CD PicView"
  lblCompanyRegistration.Caption = _
  "Derek Hall" & Chr$(10) & _
  "The Global Internet Company" & Chr$(10) & _
  "United Kingdom"

End Sub
Sub dez()
'This code will start Word and open this
'     document
    Shell "start testdoc.doc"
    'This code will open default mail client and fill in to address
    Shell "start mailto:gencross@intnet.net"
    'This code will open browser and goto this URL
    Shell "start http://www.planet-source-code.com"
    'Very cool.
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMailTo.ForeColor = &H800000
  lblMailTo.FontUnderline = False
End Sub

Private Sub lblMailTo_Click()
  Shell "start mailto:derek.hall@virgin.net"
End Sub

Private Sub lblMailTo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMailTo.FontUnderline = True
  lblMailTo.ForeColor = &HFF0000
  
End Sub
