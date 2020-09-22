VERSION 5.00
Begin VB.Form frmPicture 
   Appearance      =   0  'Flat
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Picture"
   ClientHeight    =   31995
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   31995
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   31995
   ScaleWidth      =   31995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picFullView 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4320
      Left            =   1080
      ScaleHeight     =   4320
      ScaleWidth      =   5280
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   5280
   End
   Begin VB.Image Image1 
      Height          =   3135
      Left            =   1920
      Stretch         =   -1  'True
      ToolTipText     =   "Arrow Keys = Scroll  ; ' 1 to 0,+,-,I,O' = Voom  ;  C = Center"
      Top             =   720
      Width           =   3255
   End
End
Attribute VB_Name = "frmPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub Form_Activate()
    picFullView.Left = (frmPicture.Width - picFullView.Width) / 2
    picFullView.Top = (frmPicture.Height - picFullView.Height) / 2
    Image1.Width = picFullView.Width
    Image1.Height = picFullView.Height
    Image1.Left = (frmPicture.Width - picFullView.Width) / 2
    Image1.Top = (frmPicture.Height - picFullView.Height) / 2
    Image1 = picFullView
    picFullView = LoadPicture
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
    Case 39  '(39  RIGHT ARROW)
        Image1.Left = Image1.Left - (Image1.Width / 10)
    Case 40  '(40  DOWN ARROW)
        Image1.Top = Image1.Top - (Image1.Height / 10)
    Case 37  '( 37  LEFT ARROW)
        Image1.Left = Image1.Left + (Image1.Width / 10)
    Case 38  '(38  UP ARROW)
        Image1.Top = Image1.Top + (Image1.Height / 10)
    Case 107 '(107 PLUS SIGN (+))
            Image1.Visible = False
            Image1.Width = Image1.Width + (Image1.Width / 5)
            Image1.Height = Image1.Height + (Image1.Height / 5)
            Image1.Left = (frmPicture.Width - Image1.Width) / 2
            Image1.Top = (frmPicture.Height - Image1.Height) / 2
            Image1.Visible = True
    
    Case 73  '(73 I )
        Image1.Visible = False
        Image1.Width = Image1.Width + (Image1.Width / 10)
        Image1.Height = Image1.Height + (Image1.Height / 10)
        Image1.Visible = True
    Case 79 '(79 O)
        If (Not Image1.Width < 400 Or Not Image1.Height < 400) Then
            Image1.Visible = False
            Image1.Width = Image1.Width - (Image1.Width / 15)
            Image1.Height = Image1.Height - (Image1.Height / 15)
            Image1.Visible = True
        End If
    Case 109 '(109 MINUS SIGN (-) )
        If (Not Image1.Width < 400 Or Not Image1.Height < 400) Then
            Image1.Visible = False
            Image1.Width = Image1.Width - (Image1.Width / 15)
            Image1.Height = Image1.Height - (Image1.Height / 15)
            Image1.Left = (frmPicture.Width - Image1.Width) / 2
            Image1.Top = (frmPicture.Height - Image1.Height) / 2
            Image1.Visible = True
    End If
    Case 27, 13  '(27  ESC )(13  ENTER )
        tempTime = 0
        Unload frmPicture
        
    Case 78, 48  '(78 N)(48 0)
        Image1.Visible = False
        Image1.Width = picFullView.Width
        Image1.Height = picFullView.Height
        Image1.Left = (frmPicture.Width - Image1.Width) / 2
        Image1.Top = (frmPicture.Height - Image1.Height) / 2
        Image1.Visible = True
    Case 67  '(67 C) centre the pic
        Image1.Visible = False
        Image1.Left = (frmPicture.Width - Image1.Width) / 2
        Image1.Top = (frmPicture.Height - Image1.Height) / 2
        Image1.Visible = True
     Case 49, 50, 51, 52, 53, 54, 55, 56, 57   ' (1- 9 keys)
        Image1.Visible = False
        Image1.Width = ((picFullView.Width / 50) * (KeyCode - 47)) * (KeyCode - 47) ' * ((KeyCode - 53) - Int((0.25 * (KeyCode - 53)) + 0.25))
        Image1.Height = ((picFullView.Height / 50) * (KeyCode - 47)) * (KeyCode - 47)    '* ((KeyCode - 53) - Int((0.25 * (KeyCode - 53)) + 0.25))
        Image1.Left = (frmPicture.Width - Image1.Width) / 2
        Image1.Top = (frmPicture.Height - Image1.Height) / 2
        Image1.Visible = True
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'  Form1.tempTime = 0
  Unload frmPicture
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Form1.tempTime = 0
  Unload frmPicture
End Sub

Private Sub Image1_Mousedown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Form1.tempTime = 0
Unload frmPicture
End Sub

Public Sub Showtime()
    Image1.Visible = False
    picFullView.Left = (frmPicture.Width - picFullView.Width) / 2
    picFullView.Top = (frmPicture.Height - picFullView.Height) / 2
    Image1.Width = picFullView.Width
    Image1.Height = picFullView.Height
    Image1.Left = (frmPicture.Width - picFullView.Width) / 2
    Image1.Top = (frmPicture.Height - picFullView.Height) / 2
    Image1 = picFullView
    Image1.Visible = True
End Sub

Sub coolfade(tofade As Object)


    X = tofade.Width
    Y = tofade.Height
    tofade.FillStyle = 0
    red = 0
    blue = tofade.Width


    Do Until blue = 255
        blue = blue + 5
        red = red - tofade.Width / 255 * 5
        tofade.FillColor = RGB(0, 0, blue)
        If red < 0 Then Exit Do
        tofade.Circle (tofade.Width / 2, tofade.Height / 2), red, RGB(0, red, 255)
    Loop

End Sub

