VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About hammer launcher"
   ClientHeight    =   2760
   ClientLeft      =   6390
   ClientTop       =   4440
   ClientWidth     =   4215
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4215
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Ok 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   720
         Left            =   120
         Picture         =   "About.frx":030A
         ScaleHeight     =   720
         ScaleWidth      =   720
         TabIndex        =   2
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Link 
         Caption         =   "CubeD.dk"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         MouseIcon       =   "About.frx":1FD4
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Lines: 2'722"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Language: Visual Basic"
         Height          =   255
         Left            =   1200
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Date: 26. Jul 2005"
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Author: Anders Jenbo"
         Height          =   255
         Left            =   1200
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Version v0.8.1"
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib _
              "shell32.dll" Alias "ShellExecuteA" _
              (ByVal hwnd As Long, _
               ByVal lpOperation As String, _
               ByVal lpFile As String, _
               ByVal lpParameters As String, _
               ByVal lpDirectory As String, _
               ByVal nShowCmd As Long) As Long
                
Private Const SW_SHOW = 1

Public Sub Navigate(ByVal NavTo As String)
  Dim hBrowse As Long
  hBrowse = ShellExecute(0&, "open", NavTo, "", "", SW_SHOW)
End Sub

Private Sub Ok_Click()
About.Hide
End Sub

Private Sub Link_Click()
Navigate "http://cubed.dk"
End Sub
