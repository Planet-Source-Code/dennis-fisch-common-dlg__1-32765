VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.UserControl CommonDialog 
   ClientHeight    =   4185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   DefaultCancel   =   -1  'True
   ScaleHeight     =   4185
   ScaleWidth      =   5535
   Begin MSComctlLib.ImageCombo cboType 
      Height          =   330
      Left            =   1500
      TabIndex        =   14
      Top             =   3690
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
   Begin MSComctlLib.ImageList ILFiles16 
      Left            =   4830
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox PicFiles16 
      BackColor       =   &H80000009&
      Height          =   300
      Left            =   4860
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   4350
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox PicFiles32 
      BackColor       =   &H80000009&
      Height          =   600
      Left            =   4860
      ScaleHeight     =   540
      ScaleWidth      =   540
      TabIndex        =   11
      Top             =   4320
      Visible         =   0   'False
      Width           =   600
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4830
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":0000
            Key             =   "CD"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":0352
            Key             =   "Default"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":0A64
            Key             =   "Desktop"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":0DB6
            Key             =   "Floppy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":1108
            Key             =   "Folder"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":145A
            Key             =   "HD"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":17AC
            Key             =   "MyComp"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":1AFE
            Key             =   "NetHood"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":1E50
            Key             =   "Personal"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":21A2
            Key             =   "Remote"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   1500
      TabIndex        =   9
      Top             =   3360
      Width           =   2625
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   345
      Left            =   4230
      TabIndex        =   8
      Top             =   3750
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4230
      TabIndex        =   7
      Top             =   3360
      Width           =   1245
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2655
      Left            =   30
      TabIndex        =   5
      Top             =   570
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4683
      Arrange         =   2
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      Icons           =   "ILFiles32"
      SmallIcons      =   "ILFiles16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageCombo imcExtra 
      Height          =   330
      Left            =   1020
      TabIndex        =   4
      Top             =   120
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Locked          =   -1  'True
      Text            =   "My Computer"
      ImageList       =   "ImageList1"
   End
   Begin VB.CommandButton cmdExtra 
      Enabled         =   0   'False
      Height          =   405
      Index           =   0
      Left            =   3450
      Picture         =   "CommonDialog.ctx":24F4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdExtra 
      Height          =   405
      Index           =   2
      Left            =   4320
      Picture         =   "CommonDialog.ctx":2836
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdExtra 
      Height          =   405
      Index           =   3
      Left            =   4770
      Picture         =   "CommonDialog.ctx":2EF0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   60
      Width           =   375
   End
   Begin VB.CommandButton cmdExtra 
      Height          =   405
      Index           =   1
      Left            =   3870
      Picture         =   "CommonDialog.ctx":35F2
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   375
   End
   Begin MSComctlLib.ImageList ILFiles32 
      Left            =   4830
      Top             =   4320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CommonDialog.ctx":3B74
            Key             =   "Folder"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "File Type:"
      Height          =   255
      Left            =   210
      TabIndex        =   13
      Top             =   3690
      Width           =   1245
   End
   Begin VB.Label Label2 
      Caption         =   "File Name:"
      Height          =   255
      Left            =   210
      TabIndex        =   10
      Top             =   3360
      Width           =   1245
   End
   Begin VB.Label Label1 
      Caption         =   "Search at:"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   180
      Width           =   795
   End
   Begin VB.Menu mnuView 
      Caption         =   "mnuView"
      Visible         =   0   'False
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Icons"
         Index           =   0
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Small Icons"
         Index           =   1
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "List"
         Index           =   2
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "Report"
         Index           =   3
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "mnuFile"
      Visible         =   0   'False
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendTo 
         Caption         =   "Send To"
         Begin VB.Menu mnuSendToNotePad 
            Caption         =   "NotePad"
         End
         Begin VB.Menu mnuSendToDisk 
            Caption         =   "Disk Device"
         End
         Begin VB.Menu mnuSendToDesktop 
            Caption         =   "Desktop (Create Link)"
         End
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuFileDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuFileCreateLink 
         Caption         =   "Create Link..."
      End
      Begin VB.Menu mnuFileLine3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Properties..."
      End
   End
End
Attribute VB_Name = "CommonDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal pszPath As String) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1
Dim Sort(1 To 4) As Boolean
Dim Selitem As Long
Dim History() As String
Const FileTypeSeperator = "$"
Const FileTypeDescSep = "§"
'Default Property Values:
Const m_def_Path = "C:\"
Const m_def_InitDir = "C:\"
Const m_def_Filter = ""
Const m_def_FileName = ""
'Property Variables:
Dim m_Path As String
Dim m_InitDir As String
Dim m_Filter As String
Dim m_FileName As String
Public Type FilterType
    Filtername As String
    Filter As String
End Type
'Event Declarations:
Event OperationDone(FileName As String, FileType As String)

Private Sub cboType_Click()
On Error Resume Next
Dim c1 As String
Dim c2 As String
FillList Path
With lvwItems.ListItems
    For i = 1 To .Count
        If .Item(i).Tag <> "Folder" Then
            c1 = UCase(.Item(i).Key)
            c2 = UCase(cboType.SelectedItem.Key)
            If Not (c1) Like (c2) Then
                .Remove i
                i = i - 1
                If Err.Number <> 0 Then Exit Sub
            End If
        End If
        Err.Clear
    Next i
End With
End Sub

Private Sub cmdExtra_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Select Case Index
    Case 0
        'Path = History(UBound(History))
        'ReDim Preserve History(UBound(History) - 1)
    Case 1
        Path = fs.GetFolder(Path).ParentFolder
        FillList Path
    Case 2
        newpath = InputBox("Please enter the name for the new Folder." & vbCrLf & vbCrLf & _
                            "It will be created under: " & Path & ".", , "New Folder")
        fs.CreateFolder Path & "\" & newpath
        Path = Path & "\" & newpath
        FillList Path
    Case 3
        UserControl.PopupMenu mnuView, , cmdExtra(Index).Left + x, cmdExtra(Index).Height + y
End Select
End Sub

Private Sub cmdOK_Click()
FileName = txtFileName.Text
RaiseEvent OperationDone(txtFileName.Text, cboType.SelectedItem.Text)
End Sub

Private Sub imcExtra_Click()
On Error Resume Next
Path = imcExtra.SelectedItem.Key
lvwItems.ListItems.Clear
FillList Path
End Sub

Private Sub lvwItems_AfterLabelEdit(Cancel As Integer, NewString As String)
Select Case lvwItems.ListItems(Selitem).Tag
    Case "File"
        fs.GetFile(lvwItems.ListItems(Selitem).Key).Name = NewString
    Case "Folder"
        fs.GetFolder(lvwItems.ListItems(Selitem).Key).Name = NewString
End Select
FillList Path
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Sort(ColumnHeader.Index) = Not Sort(ColumnHeader.Index)
Select Case ColumnHeader.Index
    Case 1
        SortListView lvwItems, ColumnHeader.Index, ldtString, Sort(ColumnHeader.Index)
    Case 2
        For i = 1 To lvwItems.ListItems.Count
            lvwItems.ListItems(i).SubItems(1) = Replace(lvwItems.ListItems(i).SubItems(1), " KB", "")
        Next i
        SortListView lvwItems, ColumnHeader.Index, ldtNumber, Sort(ColumnHeader.Index)
        For i = 1 To lvwItems.ListItems.Count
            If Not lvwItems.ListItems(i).SubItems(1) = "" Then lvwItems.ListItems(i).SubItems(1) = lvwItems.ListItems(i).SubItems(1) & " KB"
        Next i
    Case 3
        SortListView lvwItems, ColumnHeader.Index, ldtString, Sort(ColumnHeader.Index)
    Case 4
        SortListView lvwItems, ColumnHeader.Index, ldtDateTime, Sort(ColumnHeader.Index)
End Select
End Sub

Private Sub lvwItems_DblClick()
On Error Resume Next
With lvwItems.ListItems(Selitem)
Select Case .Tag
    Case "File"
        ShellExecute UserControl.hwnd, "Open", .Key, vbNullString, "C:\", SW_SHOWNORMAL
    Case "Folder"
        Path = .Key
        FillList Path
End Select
End With
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
Selitem = Item.Index
txtFileName.Text = fs.GetFileName(Item.Key)
FileName = Item.Key
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    UserControl.PopupMenu mnuFile, , x, y, mnuFileOpen
End If
End Sub

Private Sub mnuFileCreateLink_Click()
'todo Creatlink Property must be added!
End Sub

Private Sub mnuFileDelete_Click()
Select Case lvwItems.ListItems(Selitem).Tag
Case "Folder"
fs.DeleteFolder lvwItems.ListItems(Selitem).Key, True
Case "File"
fs.DeleteFile lvwItems.ListItems(Selitem).Key, True
End Select
FillList Path
End Sub

Private Sub mnuFileOpen_Click()
    ShellExecute UserControl.hwnd, "Open", lvwItems.ListItems(Selitem).Key, vbNullString, "C:\", SW_SHOWNORMAL
End Sub

Private Sub mnuFileProperties_Click()
On Error Resume Next
If Not Selitem = 0 Then
i = lvwItems.ListItems(Selitem).Tag
If Err.Number <> 0 Then
    frmFolderProperties.Show
    frmFolderProperties.SetProps Path
    Exit Sub
End If
Select Case lvwItems.ListItems(Selitem).Tag
    Case "File"
        frmProperties.Show
        frmProperties.SetProps lvwItems.ListItems(Selitem).Key
    Case "Folder"
        frmFolderProperties.Show
        frmFolderProperties.SetProps lvwItems.ListItems(Selitem).Key
End Select
Else
frmFolderProperties.Show
frmFolderProperties.SetProps Path
End If
End Sub

Private Sub mnuFileRename_Click()
lvwItems.StartLabelEdit
End Sub

Private Sub mnuViewOptions_Click(Index As Integer)
lvwItems.View = Index
End Sub

Private Sub UserControl_Initialize()
Dim Folders() As String
Dim Files() As String
Dim DeskPath As String
Dim PersPath As String
Dim d As Drive
Dim Vol As String
On Error Resume Next
ReDim History(0)
With lvwItems
    .ColumnHeaders.Add , , "Name"
    .ColumnHeaders.Add , , "Size"
    .ColumnHeaders.Add , , "Type"
    .ColumnHeaders.Add , , "Last Change"
End With
DeskPath = GetSettingString(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Desktop")
PersPath = GetSettingString(&H80000001, "Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders", "Personal")
With imcExtra.ComboItems
    .Add , DeskPath, "Desktop", "Desktop", "Desktop"
    .Add , PersPath, "Personal", "Personal", "Personal"
    .Add , , "My Computer", "MyComp", "MyComp", 1
    For Each d In fs.Drives
        Vol = d.VolumeName
        If Vol = "" Then Vol = DriveType(d.DriveLetter)
        .Add , d.DriveLetter & ":\", d.DriveLetter & ":\ (" & Vol & ")", GetDriveIcon(d.Path), , 2
    Next
    ShowFolderList DeskPath, Folders
    For i = 1 To UBound(Folders)
        .Add , DeskPath & "\" & Folders(i), Folders(i), "Folder", "Folder", 1
    Next i
End With
Path = InitDir
FillList Path
End Sub

Private Function DriveType(Drive As String)
    Select Case GetDriveType(Drive)
        Case 1
            DriveType = "Disk or Absent"
        Case 2
            DriveType = "Removable"
        Case 3
            DriveType = "Drive Fixed"
        Case Is = 4
            DriveType = "Remote"
        Case Is = 5
            DriveType = "Cd-Rom"
        Case Is = 6
            DriveType = "Ram disk"
        Case Else
            DriveType = "Unrecognized"
    End Select
End Function

'Private Function GetIcon(Folder As Boolean) As String
'If Folder = True Then
'    GetIcon = "Folder"
'    Exit Function
'    GetIcon = "Default"
'End If
'End Function

Private Function GetDriveIcon(Drive As String)
    Select Case GetDriveType(Drive)
        Case 1
            GetDriveIcon = "Disk"
        Case 2
            GetDriveIcon = "Disk"
        Case 3
            GetDriveIcon = "HD"
        Case Is = 4
            GetDriveIcon = "Remote"
        Case Is = 5
            GetDriveIcon = "CD"
    End Select
End Function

Private Function BuildPath(Path As String, FileName As String) As String
PathAddBackslash Path
Path = Fix_NullTermStr(Path)
Path = Path & FileName
BuildPath = Path
End Function

Private Sub FillList(sPath)
Dim hSIcon As Long 'SmallIcon
Dim hLIcon As Long 'LargeIcon
Dim c1 As String
Dim c2 As String
txtFileName.Text = ""
On Error Resume Next
Dim Folders() As String
Dim Files() As String
Dim fpath As String
lvwItems.ListItems.Clear
ShowFolderList sPath, Folders
For i = 1 To UBound(Folders)
    fpath = fs.BuildPath(sPath, Folders(i))
    hSIcon = SHGetFileInfo(fpath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    hLIcon = SHGetFileInfo(fpath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    With PicFiles16
      Set .Picture = LoadPicture("")
      .AutoRedraw = True
      r = ImageList_Draw(hSIcon, SHInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
    With PicFiles32
      Set .Picture = LoadPicture("")
      .AutoRedraw = True
      r = ImageList_Draw(hLIcon, SHInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
    'ILFiles16.ListImages.Clear
    'MsgBox ILFiles16.ListImages.Count
    ILFiles16.ListImages.Add , , PicFiles16.Image
    'MsgBox ILFiles16.ListImages.Count
    'ILFiles32.ListImages.Clear
    ILFiles32.ListImages.Add , , PicFiles32.Image
    With lvwItems.ListItems.Add(, fpath, Folders(i), ILFiles32.ListImages(1).Index, ILFiles16.ListImages(1).Index)
        '.SubItems(1) = fs.GetFolder(fPath).Size
        .SubItems(2) = fs.GetFolder(fpath).Type
        .SubItems(3) = fs.GetFolder(fpath).DateLastModified
        .Tag = "Folder"
    End With
Next i
ShowFileList sPath, Files
For i = 1 To UBound(Files)
    fpath = fs.BuildPath(sPath, Files(i))
    hSIcon = SHGetFileInfo(fpath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
    hLIcon = SHGetFileInfo(fpath, 0&, SHInfo, Len(SHInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    With PicFiles16
      Set .Picture = LoadPicture("")
      .AutoRedraw = True
      r = ImageList_Draw(hSIcon, SHInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
    With PicFiles32
      Set .Picture = LoadPicture("")
      .AutoRedraw = True
      r = ImageList_Draw(hLIcon, SHInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT)
      .Refresh
    End With
    'ILFiles16.ListImages.Clear
    ILFiles16.ListImages.Add , fs.GetExtensionName(fpath), PicFiles16.Image
    'ILFiles32.ListImages.Clear
    ILFiles32.ListImages.Add , fs.GetExtensionName(fpath), PicFiles32.Image
    With lvwItems.ListItems.Add(, fpath, Files(i), ILFiles32.ListImages(fs.GetExtensionName(fpath)).Index, ILFiles16.ListImages(fs.GetExtensionName(fpath)).Index)
        .SubItems(1) = FormatNumber(fs.GetFile(fpath).Size \ 1024, 0, vbUseDefault, , vbTrue) & " KB"
        .SubItems(2) = fs.GetFile(fpath).Type
        .SubItems(3) = fs.GetFile(fpath).DateLastModified
        .Tag = "File"
    End With
Next i
With lvwItems.ListItems
    For i = 1 To .Count
        If .Item(i).Tag <> "Folder" Then
            c1 = UCase(.Item(i).Key)
            c2 = UCase(cboType.SelectedItem.Key)
            If Not (c1) Like (c2) Then
                .Remove i
                i = i - 1
                If Err.Number <> 0 Then Exit Sub
            End If
        End If
        Err.Clear
    Next i
End With
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal New_Path As String)
    m_Path = New_Path
    ReDim Preserve History(UBound(History) + 1)
    History(UBound(History)) = New_Path
    PropertyChanged "Path"
    FillList newpath
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get InitDir() As String
    InitDir = m_InitDir
End Property

Public Property Let InitDir(ByVal New_InitDir As String)
    m_InitDir = New_InitDir
    PropertyChanged "InitDir"
End Property

Public Sub SetFilter(NewFilter() As FilterType)
Dim i As Long
    For i = LBound(NewFilter) To UBound(NewFilter)
        cboType.ComboItems.Add , NewFilter(i).Filter, NewFilter(i).Filtername
    Next i
    PropertyChanged "Filter"
End Sub

Public Function Filter(ByRef OldFilter() As FilterType)
For i = 1 To cboType.ComboItems.Count
    ReDim Preserve OldFilter(i)
    OldFilter(i).Filter = cboType.ComboItems(i).Key
    OldFilter(i).Filter = cboType.ComboItems(i).Text
Next i
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,0
Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal New_FileName As String)
    m_FileName = New_FileName
    PropertyChanged "FileName"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Path = m_def_Path
    m_InitDir = m_def_InitDir
    m_Filter = m_def_Filter
    m_FileName = m_def_FileName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Path = PropBag.ReadProperty("Path", m_def_Path)
    m_InitDir = PropBag.ReadProperty("InitDir", m_def_InitDir)
    m_Filter = PropBag.ReadProperty("Filter", m_def_Filter)
    m_FileName = PropBag.ReadProperty("FileName", m_def_FileName)
    Path = m_Path
    InitDir = m_InitDir
    FileName = m_FileName
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Path", m_Path, m_def_Path)
    Call PropBag.WriteProperty("InitDir", m_InitDir, m_def_InitDir)
    Call PropBag.WriteProperty("FileName", m_FileName, m_def_FileName)
End Sub

