VERSION 5.00
Begin VB.Form frmFolderProperties 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Properties"
   ClientHeight    =   2895
   ClientLeft      =   2760
   ClientTop       =   3630
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   60
      TabIndex        =   7
      Top             =   1020
      Width           =   4995
      Begin VB.TextBox txtFiles 
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   3855
      End
      Begin VB.TextBox txtSize 
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   210
         Width           =   3855
      End
      Begin VB.TextBox txtLoc 
         Height          =   315
         Left            =   1050
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Total Files"
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   990
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Folder Size"
         Height          =   255
         Left            =   90
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Location"
         Height          =   255
         Left            =   90
         TabIndex        =   10
         Top             =   630
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   4995
      Begin VB.TextBox txtType 
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   180
         Width           =   2925
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   570
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   360
         Left            =   180
         Picture         =   "frmFolderProperties.frx":0000
         Stretch         =   -1  'True
         Top             =   270
         Width           =   330
      End
      Begin VB.Label Label1 
         Caption         =   "Folder:"
         Height          =   255
         Left            =   870
         TabIndex        =   6
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "Folder Name"
         Height          =   255
         Left            =   870
         TabIndex        =   5
         Top             =   600
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2460
      TabIndex        =   1
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3810
      TabIndex        =   0
      Top             =   2460
      Width           =   1215
   End
End
Attribute VB_Name = "frmFolderProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub SetProps(FileName As String)
On Error Resume Next
    i = fs.GetFolder(FileName).Size \ 1024 \ 1024
    txtSize.Text = IIf((i = 0), "Too High, Unable to show Size", i & " MB")
    txtName.Text = fs.GetFolder(FileName).Name
    txtType.Text = FileName
    txtLoc.Text = fs.GetFolder(FileName).ParentFolder.path
    txtFiles.Text = fs.GetFolder(FileName).Files.Count
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Unload Me
End Sub
