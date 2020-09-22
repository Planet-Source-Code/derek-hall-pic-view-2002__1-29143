VERSION 5.00
Begin VB.UserControl PictureViewer 
   ClientHeight    =   5070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5385
   ScaleHeight     =   5070
   ScaleWidth      =   5385
   Begin VB.PictureBox CreateThumb 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1320
      ScaleHeight     =   555
      ScaleWidth      =   735
      TabIndex        =   5
      Top             =   3660
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox Original 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2940
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   3660
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer ScrollBarTimer 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4020
      Top             =   180
   End
   Begin VB.PictureBox PicWindow 
      Appearance      =   0  'Flat
      BackColor       =   &H00A0A0A0&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   600
      ScaleHeight     =   3345
      ScaleWidth      =   3345
      TabIndex        =   0
      Top             =   180
      Width           =   3375
      Begin VB.CommandButton cmdSlideShow 
         Caption         =   "S"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         ToolTipText     =   "Activate Slide Show"
         Top             =   2400
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox txt_Data 
         Alignment       =   2  'Center
         BackColor       =   &H00989898&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "100"
         ToolTipText     =   "Thumb Size"
         Top             =   1920
         Width           =   855
      End
      Begin VB.CommandButton cmdThumbSizeDown 
         Caption         =   "<"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         ToolTipText     =   "Thumb Size Down"
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton cmdThumbSizeUp 
         Caption         =   ">"
         Height          =   255
         Left            =   1680
         TabIndex        =   7
         ToolTipText     =   "Thumb Size Up"
         Top             =   2400
         Width           =   255
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "A"
         Height          =   255
         Left            =   2880
         TabIndex        =   6
         ToolTipText     =   "About"
         Top             =   3120
         Width           =   255
      End
      Begin VB.TextBox txt_Data 
         BackColor       =   &H00989898&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txt_Data 
         BackColor       =   &H00989898&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   3120
         Width           =   855
      End
      Begin VB.VScrollBar ScrollBar 
         Height          =   495
         Left            =   1200
         Max             =   0
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   240
         Width           =   255
      End
      Begin VB.Line OutLine 
         BorderColor     =   &H00000000&
         Index           =   4
         X1              =   0
         X2              =   1800
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Line OutLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   3
         Visible         =   0   'False
         X1              =   0
         X2              =   840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line OutLine 
         BorderColor     =   &H00000000&
         Index           =   2
         Visible         =   0   'False
         X1              =   0
         X2              =   840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line OutLine 
         BorderColor     =   &H00FFFFFF&
         Index           =   0
         Visible         =   0   'False
         X1              =   0
         X2              =   840
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line OutLine 
         BorderColor     =   &H00000000&
         Index           =   1
         Visible         =   0   'False
         X1              =   0
         X2              =   840
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Image Thumb 
         Appearance      =   0  'Flat
         Height          =   1035
         Index           =   10000
         Left            =   120
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
   End
End
Attribute VB_Name = "PictureViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'********BitBlit
Private Const SRCCOPY = &HCC0020     ' dest = source
Private Declare Function BitBlt Lib "GDI32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
'********
Const TimesThumbs = 3
Const DivideThumbs = 4

Private ThumbSize As Integer
Private Const New_ThumbSize = 1500

Private FilePathAndName As String
Private Const New_FilePathAndName = ""

Event Thumbclick(FilePathAndName As String, Button As Integer)
Event ThumbDblClick(FilePathAndName As String)

Dim DATENOW As Date
Dim PicsAcross As Integer
Dim PicsDown As Integer
Dim Spaceing As Integer
Dim LastThumbIndex As Integer
Dim MaxThumbs As Long
Dim LastScrollSetting As Integer
Dim ExitEvents As Boolean ', CurrentThumb As Integer
Dim RegisteredCopy As Integer

Public Sub ClearAllThumbs()
  ExitEvents = True
  Dim i As Integer
  For i = 0 To Thumb.Count - 2
    DoEvents
    Thumb(i) = LoadPicture
    Unload Thumb(i)
  Next i
  InvisibleOutlines
  ClearTxtData
  ExitEvents = False
End Sub
Public Sub DeleteAllThumbs()
  InvisibleOutlines
  ClearAllThumbs
  ClearInfo
  ClearTxtData
End Sub

Private Sub SortThumbs()
If MaxThumbs = 0 Then Exit Sub
InvisibleOutlines
Dim i As Integer
  For i = 0 To Thumb.Count - 2
    DoEvents
    Thumb(i).Visible = False
    Thumb(i).Top = ((Int(i / PicsAcross) * (ThumbSize + Spaceing)) + Spaceing)
    Thumb(i).Left = ((i Mod PicsAcross) * (ThumbSize + Spaceing)) + Spaceing
    If ExitEvents Then Exit Sub
  Next i
  For i = 0 To Thumb.Count - 2
    DoEvents
    Thumb(i).Visible = True
    If ExitEvents Then Exit Sub
  Next i
  If Thumb.Count > 1 And OutLine(1).Visible = True Then DrawBoarder LastThumbIndex, False
End Sub

Private Sub cmdAbout_Click()
  frmAbout.Show , Me
End Sub


Private Sub cmdSlideShow_Click()

  If UBound(Info) > 0 Then
  On Error Resume Next
  PicShowOn = True
  Dim i As Integer
  frmPicture.Show , Me
  For i = 0 To UBound(Info) - 1
    frmPicture.picFullView.Picture = LoadPicture(Thumb(i).Tag)
    
    DoEvents
    s_Wait 2
    frmPicture.Form_Activate
    If ExitEvents Then Exit Sub
  Next i
  Unload frmPicture
  Else
    MsgBox "You need to select some pictures to activate the slideshow"
  End If
End Sub

Private Sub cmdThumbSizeDown_Click()
  ChangeThumbSize = (ThumbSize / 15) - 20
End Sub

Private Sub cmdThumbSizeUp_Click()
  ChangeThumbSize = (ThumbSize / 15) + 20
End Sub



Private Sub Thumb_Click(Index As Integer)
  If Thumb(Index).Tag = "" Then Exit Sub
  frmPicture.Show , Me
  CurrentFilePathAndName = Thumb(Index).Tag
  frmPicture.picFullView.Picture = LoadPicture(Thumb(Index).Tag)
  RaiseEvent ThumbDblClick(CurrentFilePathAndName)
End Sub

'Private Sub Thumb_DblClick(Index As Integer)
'  If Thumb(Index).Tag = "" Then Exit Sub
'  frmPicture.Show , Me
'  CurrentFilePathAndName = Thumb(Index).Tag
'  frmPicture.picFullView.Picture = LoadPicture(Thumb(Index).Tag)
'  RaiseEvent ThumbDblClick(CurrentFilePathAndName)
'End Sub
Private Sub InvisibleOutlines()
  Dim i As Integer
  For i = 0 To 3
    OutLine(i).Visible = False
    If ExitEvents Then Exit Sub
  Next i
  CurrentFilePathAndName = ""
End Sub
Private Sub DrawBoarder(Index As Integer, EdgeStyle As Boolean)
  On Error Resume Next
  Dim i As Integer, Distance As Integer
  CurrentThumb = Index
  Distance = 15
  InvisibleOutlines
  Select Case EdgeStyle
    Case True
      OutLine(0).BorderColor = 0
      OutLine(1).BorderColor = &HFFFFFF
      OutLine(2).BorderColor = &HFFFFFF
      OutLine(3).BorderColor = 0
    Case False
      OutLine(0).BorderColor = &HFFFFFF
      OutLine(1).BorderColor = 0
      OutLine(2).BorderColor = 0
      OutLine(3).BorderColor = &HFFFFFF
  End Select
  'top
  OutLine(0).X1 = Thumb(Index).Left - Distance
  OutLine(0).Y1 = Thumb(Index).Top - Distance
  OutLine(0).X2 = OutLine(0).X1 + Thumb(Index).Width + (Distance * 2)
  OutLine(0).Y2 = OutLine(0).Y1
  'bottom
  OutLine(2).X1 = OutLine(0).X1
  OutLine(2).Y1 = Thumb(Index).Top + Thumb(Index).Height + Distance
  OutLine(2).X2 = OutLine(0).X2
  OutLine(2).Y2 = OutLine(2).Y1
  
 'left
  OutLine(1).X1 = OutLine(0).X2
  OutLine(1).Y1 = OutLine(0).Y1
  OutLine(1).X2 = OutLine(0).X2
  OutLine(1).Y2 = OutLine(0).Y1 + Thumb(Index).Height + (Distance * 3)
'right
  OutLine(3).X1 = OutLine(0).X1
  OutLine(3).Y1 = OutLine(1).Y1
  OutLine(3).X2 = OutLine(0).X1
  OutLine(3).Y2 = OutLine(1).Y2
  For i = 0 To 3
    OutLine(i).Visible = True
  Next i
End Sub
Private Sub Thumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  DrawBoarder Index, True
  LastThumbIndex = Index
  txt_Data(1) = Thumb(Index).Tag
  CurrentFilePathAndName = Thumb(Index).Tag
  RaiseEvent Thumbclick(CurrentFilePathAndName, Button)
End Sub
Private Sub Thumb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  DrawBoarder Index, False
End Sub


Private Function f_CaseExtension(strExtend As String) As Boolean
  Dim Extension As String
  Extension = Format(Right$(strExtend, 4), "<")
  Select Case Extension
    Case "jpeg", ".jpg", ".bmp", ".ico", ".gif", ".emf", ".wmf", ".rle"
      f_CaseExtension = True
    Case Else
    f_CaseExtension = False
  End Select
End Function
Private Function GetCorrectPath(strToCheck As String) As String
 If Right$(strToCheck, 1) = "\" Then
    GetCorrectPath = strToCheck
  Else
    GetCorrectPath = strToCheck & "\"
  End If
End Function

Sub MakeThumb(FilePath As String, FileNameAndExt As String)
  txt_Data(1).Text = "Loading file: " & FileNameAndExt
  DoEvents
  If UBound(Info) > 31999 Then Exit Sub
  If f_CaseExtension(FileNameAndExt) Then
    
    AddInfo FilePath, FileNameAndExt
    SetScrollBar
    ShowThumb FilePath, FileNameAndExt
  End If
  UpdateTxtData
End Sub
Private Sub ScrollBar_Change()
    ScrollBarTimer.Enabled = True
End Sub
Private Sub ScrollBarTimer_Timer()
  InvisibleOutlines
  ScrollBarTimer.Enabled = False
  ScrollBarEvents
End Sub
Private Sub ScrollBarEvents()
  Dim i As Integer, ActualCount As Integer
  ClearAllThumbs
  ActualCount = MaxThumbs * ScrollBar.Value
  For i = ActualCount To InfoCount - 1
    If ScrollBarTimer.Enabled Or ExitEvents Then
      Exit For
    End If
    DoEvents
    ShowThumb Info(i).FilePath, Info(i).FileName
    UpdateTxtData
    If ExitEvents Then Exit Sub
  Next i
End Sub
Private Sub UpdateTxtData()
  txt_Data(0) = "Thumbs " & (MaxThumbs * ScrollBar.Value) + 1 & " To " & ((Thumb.Count - 1) + ((MaxThumbs * ScrollBar.Value))) & " Of " & UBound(Info)
  txt_Data(2) = (ThumbSize / 15)
End Sub
Private Sub ClearTxtData()
  txt_Data(0) = "Thumbs 0 To 0 Of " & UBound(Info)
End Sub
Private Sub SetScrollBar()
  If (InfoCount) > MaxThumbs And MaxThumbs > 0 Then
    ScrollBar.Max = Int((InfoCount - 1) / MaxThumbs)
  Else
    ScrollBar.Max = 0
  End If
End Sub

Private Sub UserControl_Initialize()
  'CheckDate

  PicWindow.Top = 0
  PicWindow.Left = 0
  Spaceing = 90
  LastThumbIndex = 1000
  ClearInfo
  'If RegisteredCopy <> 0 Then
  '  Exit Sub
  'Else
  '  If GetVolName(Left$(App.Path, 2) & "\") <> "TGIC" Then
  '    On Error GoTo EndHere
  '    MsgBox "Not a Registered Copy, Please Contact Author. This program will fail to respond"
  '    RegisteredCopy = 1
      'frmAbout.Show , Me
  '  Else
      RegisteredCopy = 2
  '  End If
  'End If
EndHere:
End Sub


Private Sub ShowThumb(FilePath As String, FileNameAndExt As String)
  If RegisteredCopy <> 2 Then Exit Sub
  

  If (Thumb.Count) > MaxThumbs Then Exit Sub
  Dim tmpHeight As Long, tmpWidth As Long, NextThumb As Integer
  Dim DivisionX As Integer, DivisionY As Integer, i As Integer, j As Integer, dividitup As Integer
  NextThumb = Thumb.Count - 1
  
  Load Thumb(NextThumb)
  With Thumb(NextThumb)
    On Error Resume Next
    txt_Data(1).Text = "Loading file: " & FileNameAndExt
    DoEvents
    Original.Picture = LoadPicture((GetCorrectPath(FilePath) & FileNameAndExt))
    .Visible = False
    .Stretch = True
    .Tag = GetCorrectPath(FilePath) & FileNameAndExt
    tmpHeight = Original.Height
    tmpWidth = Original.Width
    Do While tmpHeight > ThumbSize Or tmpWidth > ThumbSize
      tmpHeight = Int((tmpHeight / DivideThumbs) * TimesThumbs)
      tmpWidth = Int((tmpWidth / DivideThumbs) * TimesThumbs)
      If ExitEvents Then Exit Sub
    Loop
'****************
 'Debug.Print "W:" & tmpWidth & "  " & "H:" & tmpHeight & " ThumbSize:" & ThumbSize
  DivisionX = Int(Original.Width / tmpWidth) + 1
  DivisionY = Int(Original.Height / tmpHeight) + 1
  CreateThumb.Width = (Original.Width / DivisionX)
  CreateThumb.Height = (Original.Height / DivisionY)
 ' Debug.Print "DivisionY:" & DivisionY & "  " & "DivisionX:" & DivisionX

  For i = 0 To Int(CreateThumb.Height / 15)
    For j = 0 To Int(CreateThumb.Width / 15)
      BitBlt CreateThumb.hDC, j, i, 1, 1, Original.hDC, j * DivisionX, i * DivisionY, SRCCOPY
    Next j
   Next i
  .Picture = CreateThumb.Image
  .Refresh
  CreateThumb = LoadPicture
'****************
    .ToolTipText = FileNameAndExt
    .Stretch = False
    .Top = ((Int(NextThumb / PicsAcross) * (ThumbSize + Spaceing)) + Spaceing)
    .Left = ((NextThumb Mod PicsAcross) * (ThumbSize + Spaceing)) + Spaceing
    .Visible = True
  End With
  DoEvents
End Sub
Public Sub Resize()
  'If CheckDate = True Then Exit Sub
  On Error Resume Next
  PicWindow.Width = UserControl.Width '/ 15
  PicWindow.Height = UserControl.Height ' / 15
  
  ScrollBar.Top = 0
  ScrollBar.Left = PicWindow.Width - (ScrollBar.Width + 1)
  ScrollBar.Height = PicWindow.Height - 255
  
  txt_Data(0).Height = 255
  txt_Data(0).Top = PicWindow.Height - txt_Data(0).Height
  txt_Data(0).Left = 0
  txt_Data(0).Width = 2000
  
  txt_Data(1).Height = 255
  txt_Data(1).Top = txt_Data(0).Top
  txt_Data(1).Left = txt_Data(0).Width + 45
  txt_Data(1).Width = Int(PicWindow.Width - (txt_Data(0).Width + 45)) - (255 * 6)
  
  txt_Data(2).Height = 255
  txt_Data(2).Top = txt_Data(0).Top
  txt_Data(2).Left = PicWindow.Width - (255 * 4)
  txt_Data(2).Width = 510
  
  OutLine(4).X1 = txt_Data(0).Left
  OutLine(4).X2 = PicWindow.Width - 15
  OutLine(4).Y1 = txt_Data(0).Top - 15
  OutLine(4).Y2 = OutLine(4).Y1
  
  cmdAbout.Height = 255
  cmdAbout.Top = PicWindow.Height - 255
  cmdAbout.Width = 255
  cmdAbout.Left = PicWindow.Width - 255
  
  cmdThumbSizeDown.Height = 255
  cmdThumbSizeDown.Top = PicWindow.Height - 255
  cmdThumbSizeDown.Width = 255
  cmdThumbSizeDown.Left = PicWindow.Width - (255 * 5)
  
  cmdThumbSizeUp.Height = 255
  cmdThumbSizeUp.Top = PicWindow.Height - 255
  cmdThumbSizeUp.Width = 255
  cmdThumbSizeUp.Left = PicWindow.Width - (255 * 2)
  
  cmdSlideShow.Height = 255
  cmdSlideShow.Top = PicWindow.Height - 255
  cmdSlideShow.Width = 255
  cmdSlideShow.Left = PicWindow.Width - (255 * 6)
  SortOutSizes
  UpdateAll

End Sub
Public Sub SortOutSizes()
  PicsAcross = Int(ScrollBar.Left / (ThumbSize + Spaceing))
  PicsDown = Int(txt_Data(0).Top / (ThumbSize + Spaceing))
  MaxThumbs = PicsAcross * PicsDown
End Sub
Public Sub UpdateAll()
  SortThumbs
  SetScrollBar
  UpdateThumbs
End Sub
Private Sub UpdateThumbs()
  Dim i As Integer, ActualCount As Integer
  ActualCount = MaxThumbs * ScrollBar.Value
  For i = (ActualCount + Thumb.Count - 1) To InfoCount - 1
    If ScrollBarTimer.Enabled Or ExitEvents Then
      Exit For
    End If
    ShowThumb Info(i).FilePath, Info(i).FileName
    UpdateTxtData
    If ExitEvents Then Exit Sub
  Next i
End Sub

Public Sub TerminateMe()
  ExitEvents = True
End Sub

Private Function CheckDate() As Boolean
    
    'DATENOW = "25/12/00"
    'If Date > DATENOW Then
    '  MsgBox "You need to register this OCX control, please register with derek.hall@virgin.net (Cost UK Pounds 5.00)"
    '  CheckDate = True
    'Else
      CheckDate = False
    'End If
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ThumbSize = PropBag.ReadProperty("ChangeThumbSize", New_ThumbSize)
  FilePathAndName = PropBag.ReadProperty("CurrentFilePathAndName", New_FilePathAndName)
End Sub

Private Sub UserControl_Terminate()
  ExitEvents = True
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "ChangeThumbSize", ThumbSize, New_ThumbSize
  PropBag.WriteProperty "CurrentFilePathAndName", FilePathAndName, New_FilePathAndName
End Sub
Private Property Get CurrentFilePathAndName() As String
  CurrentFilePathAndName = FilePathAndName
End Property
Private Property Let CurrentFilePathAndName(strPathAndName As String)
  FilePathAndName = strPathAndName
  PropertyChanged "CurrentFilePathAndName"
End Property

Public Property Get ChangeThumbSize() As Integer
  ChangeThumbSize = ThumbSize
End Property
Public Property Let ChangeThumbSize(Value As Integer)
  If Not Value = (ThumbSize / 15) Then
    If Value < 50 Then Value = 50
    If Value > 300 Then Value = 300
    ThumbSize = (Value * 15)
    PropertyChanged "ChangeThumbSize"
    InvisibleOutlines
    SortOutSizes
    txt_Data(2) = (ThumbSize / 15)
    SetScrollBar
    ScrollBarTimer.Enabled = True
  End If
End Property

Private Sub s_Wait(TimeToWait As Variant)
  Dim Start
  Start = Timer   ' Set start time.
  Do While Timer < Start + TimeToWait  ' Until added time.
    DoEvents    ' Yield to other processes.
    If ExitEvents Then Exit Sub
  Loop
End Sub
