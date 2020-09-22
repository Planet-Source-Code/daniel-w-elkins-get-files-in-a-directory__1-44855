VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Files in Directory"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8205
   StartUpPosition =   2  'CenterScreen
   Begin prjFileDir.Button cmdListFiles 
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   5760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "List Files"
   End
   Begin MSComctlLib.ImageList ILFiles 
      Left            =   1560
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E82
            Key             =   "FileName"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":32D4
            Key             =   "FileSize"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3726
            Key             =   "FilePath"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView LVFiles 
      Height          =   5055
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   8916
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ILFiles"
      SmallIcons      =   "ILFiles"
      ColHdrIcons     =   "ILFiles"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   3440
         ImageIndex      =   1
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "File Size"
         Object.Width           =   3175
         ImageIndex      =   2
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path (Directory)"
         Object.Width           =   7223
         ImageIndex      =   3
      EndProperty
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin prjFileDir.Button cmdFileNum 
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
      Caption         =   "Number of Files in Dir."
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter a Directory :"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   1605
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   120
      Picture         =   "frmMain.frx":3B78
      Stretch         =   -1  'True
      Top             =   120
      Width           =   360
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdFileNum_Click()
If Len(txtPath.Text) = 0 Then
MsgBox "Please enter a directory to count", vbCritical, "Directory Required"
txtPath.SetFocus
Else
MsgBox "There are " & DirFileNum(txtPath.Text) & " files in this directory.", vbInformation, "Number of Files in Directory"
End If
End Sub

Private Sub cmdListFiles_Click()
If Len(txtPath.Text) = 0 Then
MsgBox "Please enter a directory to list", vbCritical, "Directory Required"
txtPath.SetFocus
Else
Call AddDirFiles(txtPath.Text, LVFiles)
End If
End Sub

Private Sub Form_Load()
txtPath.Text = WinDir
End Sub
