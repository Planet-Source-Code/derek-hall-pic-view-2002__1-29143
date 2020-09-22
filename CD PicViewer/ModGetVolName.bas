Attribute VB_Name = "ModGetVolName"
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Function GetVolName(Root As String) As String
Dim volume_name As String

Dim max_component_length As Long
Dim file_system_flags As Long
Dim file_system_name As String
Dim pos As Integer
    
    On Error Resume Next

    'Root = Combo1.Text
    volume_name = Space$(1024)
    'file_system_name = Space$(1024)

    If GetVolumeInformation(Root, volume_name, _
        Len(volume_name), serial_number, _
        max_component_length, file_system_flags, _
        file_system_name, Len(file_system_name)) = 0 _
    Then
        MsgBox "No Disk In Drive!", vbInformation, "Error Reading Disk"
        Exit Function
    End If

    pos = InStr(volume_name, Chr$(0))
    volume_name = Left$(volume_name, pos - 1)
    GetVolName = volume_name
End Function


