VERSION 5.00
Object = "{58E0B815-4E70-11D3-8D40-B4841EB66730}#30.0#0"; "PicView.ocx"
Begin VB.Form FileManager 
   AutoRedraw      =   -1  'True
   Caption         =   "CD Quick View"
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   5400
      Width           =   1095
   End
   Begin VB.ListBox lstThumbSize 
      Height          =   1230
      ItemData        =   "FileManager.frx":0000
      Left            =   1200
      List            =   "FileManager.frx":0022
      TabIndex        =   5
      Top             =   5280
      Width           =   735
   End
   Begin PicView.PictureViewer PicView 
      Height          =   435
      Left            =   1980
      TabIndex        =   4
      Top             =   0
      Width           =   435
      _ExtentX        =   661
      _ExtentY        =   1508
   End
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   1935
   End
   Begin VB.CommandButton cmdClearAll 
      Caption         =   "Clear All"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   6000
      Width           =   1095
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   1935
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "FileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim exitnow As Boolean

Private Sub cmdClearAll_Click()
  PicView.DeleteAllThumbs
End Sub

Private Sub Dir1_Change()
  On Error Resume Next
  File1 = Dir1
End Sub

Private Sub Drive1_Change()
  On Error Resume Next
  Dir1 = Drive1
End Sub

Private Sub File1_DblClick()
  Dim i As Integer
  For i = 0 To File1.ListCount - 1
    PicView.MakeThumb File1.Path, File1.List(i)
    If exitnow Then Exit Sub
  Next i
End Sub

Private Sub Form_Load()
'  If GetVolName(Left$(App.Path, 2) & "\") <> "TGIC" Then End
'  Drive1 = Left$(App.Path, 2) & "\Pics"
 'Dir1.Path = Left$(App.Path, 2) & "\Pics"
  Drive1 = "C:"
  'Dir1.Path = "F:\A_Our House\Art\Pics"
  'Dir1.Path = "G:\aa-PHOTOSHOP Actions\Actions and Pictures"
  'Dir1_Change
  'File1_DblClick

End Sub

Private Sub Form_Resize()
On Error Resume Next
  PicView.Top = 0
  PicView.Left = File1.Width + 50 '0
  PicView.Width = FileManager.Width - (File1.Width + 165) '120
  PicView.Height = FileManager.Height - 400
  PicView.Resize
End Sub

Private Sub Form_Terminate()
  PicView.TerminateMe
  exitnow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
  PicView.TerminateMe
  exitnow = True
End Sub

Private Sub lstThumbSize_Click()
  PicView.ChangeThumbSize = lstThumbSize
  PicView.UpdateAll
End Sub

Private Sub PicView_ThumbClick(FilePathAndName As String, button As Integer)
 'If button = 2 Then MsgBox FilePathAndName

End Sub
