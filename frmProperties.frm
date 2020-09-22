VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   2670
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1065
      Left            =   60
      TabIndex        =   8
      Top             =   1080
      Width           =   4995
      Begin VB.TextBox txtLoc 
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   600
         Width           =   4095
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   210
         Width           =   4095
      End
      Begin VB.Label Label4 
         Caption         =   "Location"
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "File Size"
         Height          =   255
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4995
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   570
         Width           =   2925
      End
      Begin VB.TextBox txtType 
         Height          =   315
         Left            =   1770
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   180
         Width           =   2925
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   180
         ScaleHeight     =   39
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   37
         TabIndex        =   3
         Top             =   270
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "File Name"
         Height          =   255
         Left            =   870
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "File Type"
         Height          =   255
         Left            =   870
         TabIndex        =   4
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   1
      Top             =   2220
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   2460
      TabIndex        =   0
      Top             =   2220
      Width           =   1215
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const DI_MASK = &H1
Const DI_IMAGE = &H2
Const DI_NORMAL = DI_MASK Or DI_IMAGE
Private Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Public Sub SetProps(FileName As String)
    Dim mIcon As Long
    mIcon = ExtractAssociatedIcon(App.hInstance, FileName, 2)
    DrawIconEx Picture1.hdc, 0, 0, mIcon, 0, 0, 0, 0, DI_NORMAL
    DestroyIcon mIcon
    txtSize.Text = FormatNumber(fs.GetFile(FileName).Size \ 1024, 0, , , vbTrue) & " KB"
    txtName.Text = fs.GetFileName(FileName)
    txtType.Text = fs.GetFile(FileName).Type
    txtLoc.Text = fs.GetFile(FileName).ParentFolder.path
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
Unload Me
End Sub
