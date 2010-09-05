VERSION 5.00
Begin VB.Form secUsers 
   BorderStyle     =   0  'None
   Caption         =   "Section Users"
   ClientHeight    =   5400
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   LinkTopic       =   "Form1"
   Picture         =   "secUsers.frx":0000
   ScaleHeight     =   5400
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ListBox allList 
      Height          =   3375
      Left            =   4730
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ListBox secList 
      Height          =   3375
      Left            =   720
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label all_users 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ALL USERS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   5650
      TabIndex        =   1
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label curren_sec_users 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "CURRENT SECTION USERS"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   255
      Left            =   1060
      TabIndex        =   0
      Top             =   675
      Width           =   2535
   End
   Begin VB.Image exit_sec_user 
      Height          =   465
      Left            =   7600
      Picture         =   "secUsers.frx":11380
      Top             =   130
      Width           =   540
   End
   Begin VB.Image right_arrow 
      Height          =   480
      Left            =   3840
      Picture         =   "secUsers.frx":11A25
      Top             =   3120
      Width           =   510
   End
   Begin VB.Image left_arrow 
      Height          =   435
      Left            =   3840
      Picture         =   "secUsers.frx":11ECD
      Top             =   2400
      Width           =   510
   End
End
Attribute VB_Name = "secUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub exit_sec_user_Click()
    Unload Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set left_arrow.Picture = MainForm.winButtonsImg.ListImages(9).Picture
    Set right_arrow.Picture = MainForm.winButtonsImg.ListImages(11).Picture
    Set exit_sec_user.Picture = MainForm.winButtonsImg.ListImages(13).Picture
End Sub

Private Sub left_arrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set left_arrow.Picture = MainForm.winButtonsImg.ListImages(10).Picture
End Sub

Private Sub right_arrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set right_arrow.Picture = MainForm.winButtonsImg.ListImages(12).Picture
End Sub
Private Sub exit_sec_user_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set exit_sec_user.Picture = MainForm.winButtonsImg.ListImages(14).Picture
End Sub
