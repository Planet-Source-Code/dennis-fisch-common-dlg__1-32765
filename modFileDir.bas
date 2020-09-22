Attribute VB_Name = "modFileDir"
Public fs As New FileSystemObject
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Const HKEY_CURRENT_USER = &H80000001
Global Path As String

Sub ShowFolderList(FolderSpec, ByRef FolderList() As String)
    On Error Resume Next
    Dim f As Folder, f1 As Folder, fc, i As Long
    Set f = fs.GetFolder(FolderSpec)
    Set fc = f.SubFolders
    For Each f1 In fc
        i = i + 1
        ReDim Preserve FolderList(i)
        FolderList(i) = f1.Name
    Next
End Sub

Sub ShowFileList(FolderSpec, ByRef FileList() As String)
    On Error Resume Next
    Dim f As Folder, f1 As File, fc As Files, i As Long
    Set f = fs.GetFolder(FolderSpec)
    Set fc = f.Files
    For Each f1 In fc
        i = i + 1
        ReDim Preserve FileList(i)
        FileList(i) = f1.Name
    Next
End Sub

Public Function GetSettingString(hKey As Long, strPath As String, strValue As String) As String
    Dim hCurKey As Long
    Dim lngDataBufferSize As Long
    Dim lngValueType As Long
    Dim strBuffer As String
    apiError = RegOpenKeyEx(hKey, strPath, 0, &H1, hCurKey)
    If Not apiError = 0 Then
        Exit Function
    End If
    apiError = RegQueryValueEx(hCurKey, strValue & Chr$(0), 0&, lngValueType, ByVal 0&, lngDataBufferSize)
    If lngValueType = 1 Then
        strBuffer = Space$(lngDataBufferSize)
        apiError = RegQueryValueEx(hCurKey, strValue & Chr$(0), 0&, 0&, ByVal strBuffer, lngDataBufferSize)
        GetSettingString = Fix_NullTermStr(strBuffer)
    End If
    apiError = RegCloseKey(hCurKey)
End Function

Public Function Fix_NullTermStr(strData As String) As String
    If strData = "" Then Exit Function
    If InStr(1, strData, Chr$(0)) = 0 Then
        Exit Function
    Else
        Fix_NullTermStr = Left$(strData, InStr(1, strData, Chr$(0)) - 1) '-1 for removing null also
    End If
End Function

Public Function Word(ByVal sSource As String, n As Long, SP As String) As String
Dim pointer As Long
Dim pos As Long
Dim x As Long
Dim lEnd As Long
On Error Resume Next
x = 1
pointer = 1
Do
   Do While Mid$(sSource, pointer, 1) = SP
      pointer = pointer + 1
   Loop
   If x = n Then
      lEnd = InStr(pointer, sSource, SP)
      If lEnd = 0 Then lEnd = Len(sSource) + 1
      Word = Mid$(sSource, pointer, lEnd - pointer)
      Exit Do
   End If
   pos = InStr(pointer, sSource, SP)
   If pos = 0 Then Exit Do
   x = x + 1
   pointer = pos + 1
Loop
End Function
