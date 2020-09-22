VERSION 5.00
Begin VB.Form frmFolders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Folders"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFolders.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   15060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   """Copy To Clipboard"""
      Height          =   495
      Left            =   4620
      TabIndex        =   5
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy To Clipboard"
      Height          =   495
      Left            =   4620
      TabIndex        =   4
      Top             =   840
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   6420
      TabIndex        =   3
      Top             =   840
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Explore"
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Top             =   840
      Width           =   1155
   End
   Begin VB.TextBox txtFolder 
      Height          =   375
      Left            =   3300
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   180
      Width           =   11595
   End
   Begin VB.ListBox lstFolder 
      Height          =   4935
      Left            =   180
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   2955
   End
End
Attribute VB_Name = "frmFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

Shell "explorer.exe " & Chr$(34) & txtFolder.Text & Chr$(34), vbNormalFocus

End Sub


Private Sub Command2_Click()

Unload Me
End

End Sub


Private Sub Command3_Click()

Clipboard.Clear
Clipboard.SetText Chr$(34) & txtFolder.Text & Chr$(34), vbCFText

End Sub

Private Sub Command4_Click()

Clipboard.Clear
Clipboard.SetText txtFolder.Text, vbCFText

End Sub

Private Sub Form_Load()

With lstFolder
.AddItem "CD Burning Cache"
.ItemData(.NewIndex) = 59&
.AddItem "Common Admin Tools"
.ItemData(.NewIndex) = 47&
.AddItem "Common Application Data"
.ItemData(.NewIndex) = 35&
.AddItem "Common Desktop"
.ItemData(.NewIndex) = 25&
.AddItem "Common Document Templates"
.ItemData(.NewIndex) = 45&
.AddItem "Common Favorites"
.ItemData(.NewIndex) = 31&
.AddItem "Common My Documents"
.ItemData(.NewIndex) = 46&
.AddItem "Common My Pictures"
.ItemData(.NewIndex) = 54&
.AddItem "Common Program Files"
.ItemData(.NewIndex) = 43&
.AddItem "Common Start Menu"
.ItemData(.NewIndex) = 22&
.AddItem "Common Start Menu Programs"
.ItemData(.NewIndex) = 23&
.AddItem "Common Startup"
.ItemData(.NewIndex) = 24&
.AddItem "Fonts"
.ItemData(.NewIndex) = 20&
.AddItem "Program Files"
.ItemData(.NewIndex) = 38&
.AddItem "System32 Folder"
.ItemData(.NewIndex) = 41&
.AddItem "System Folder"
.ItemData(.NewIndex) = 37&
.AddItem "Themes"
.ItemData(.NewIndex) = 56&
.AddItem "User Admin Tools"
.ItemData(.NewIndex) = 48&
.AddItem "User Application Data"
.ItemData(.NewIndex) = 26&
.AddItem "User Cookies"
.ItemData(.NewIndex) = 33&
.AddItem "User Desktop"
.ItemData(.NewIndex) = 16&
.AddItem "User Document Templates"
.ItemData(.NewIndex) = 21&
.AddItem "User Favorites"
.ItemData(.NewIndex) = 6&
.AddItem "User History"
.ItemData(.NewIndex) = 34&
.AddItem "User Local Application Data"
.ItemData(.NewIndex) = 28&
.AddItem "User My Documents"
.ItemData(.NewIndex) = 5&
.AddItem "User My Music"
.ItemData(.NewIndex) = 13&
.AddItem "User My Pictures"
.ItemData(.NewIndex) = 39&
.AddItem "User Net Hood"
.ItemData(.NewIndex) = 19&
.AddItem "User Print Hood"
.ItemData(.NewIndex) = 27&
.AddItem "User Profile Folder"
.ItemData(.NewIndex) = 40&
.AddItem "User Recent Documents"
.ItemData(.NewIndex) = 8&
.AddItem "User SendTo"
.ItemData(.NewIndex) = 9&
.AddItem "User Start Menu"
.ItemData(.NewIndex) = 11&
.AddItem "UserStartMenuPrograms"
.ItemData(.NewIndex) = 2&
.AddItem "User Startup"
.ItemData(.NewIndex) = 7&
.AddItem "UserTempInternetFiles"
.ItemData(.NewIndex) = 32&
.AddItem "Windows Folder"
.ItemData(.NewIndex) = 36&
End With

End Sub


Private Sub Form_Unload(Cancel As Integer)

End

End Sub


Private Sub lstFolder_Click()

Dim l As Long

l = CLng(lstFolder.ItemData(lstFolder.ListIndex))
txtFolder = SpecialFolderPath(l)

End Sub


