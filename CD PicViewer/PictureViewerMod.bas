Attribute VB_Name = "PictureViewerMod"
Type PicInfo
  FilePath As String
  FileName As String
End Type

Public Info() As PicInfo

Public Sub ClearInfo()
  On Error Resume Next
  ReDim Info(0)
End Sub

Public Sub AddInfo(FilePath As String, FileNameAndExt As String)
  On Error Resume Next
  Info(InfoCount).FilePath = FilePath
  Info(InfoCount).FileName = FileNameAndExt
  ReDim Preserve Info(InfoCount + 1)
End Sub
Public Function InfoCount() As Long
  On Error Resume Next
  InfoCount = UBound(Info)
End Function
