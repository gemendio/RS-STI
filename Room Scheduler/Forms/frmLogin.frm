VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login to Room Scheduler"
   ClientHeight    =   5730
   ClientLeft      =   -60
   ClientTop       =   -105
   ClientWidth     =   4980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Palette         =   "frmLogin.frx":0000
   Picture         =   "frmLogin.frx":F0BB
   ScaleHeight     =   5730
   ScaleWidth      =   4980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPwd 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   320
      IMEMode         =   3  'DISABLE
      Left            =   1120
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "p@ssword"
      Top             =   3670
      Width           =   2655
   End
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   320
      Left            =   1120
      TabIndex        =   2
      Text            =   "root"
      Top             =   2830
      Width           =   2655
   End
   Begin VB.TextBox txtDB 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   300
      Left            =   1120
      TabIndex        =   1
      Text            =   "room_scheduler"
      Top             =   2010
      Width           =   2655
   End
   Begin VB.TextBox txtServer 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1120
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   1190
      Width           =   2655
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   840
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   38
      ImageHeight     =   37
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin.frx":24BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin.frx":25616
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin.frx":26030
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmLogin.frx":2696E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image cmdCancel 
      Height          =   525
      Left            =   2510
      Picture         =   "frmLogin.frx":272AC
      Top             =   4320
      Width           =   525
   End
   Begin VB.Image cmdLogin 
      Height          =   555
      Left            =   1840
      Picture         =   "frmLogin.frx":279C5
      Top             =   4320
      Width           =   570
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    Dim conn As New Model.Db

    conn.Database = txtDB.Text
    conn.Password = txtPwd.Text
    conn.Server = txtServer.Text
    conn.UserName = txtUserName.Text
    
    Dim connstr As String
    
    connstr = conn.ToString
    
    Open App.Path & "/conf.ini" For Output As #1
    Print #1, connstr
    
    Close #1
    
    Unload Me
    MainForm.Show
End Sub

Private Sub cmdLogin_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set cmdLogin.Picture = ImageList1.ListImages(2).Picture
End Sub
Private Sub cmdCancel_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set cmdCancel.Picture = ImageList1.ListImages(4).Picture
End Sub
Private Sub Form_Load()
    frmLogin.Left = (Screen.Width / 2) - (frmLogin.Width / 2)
    frmLogin.Top = (Screen.Height / 2) - (frmLogin.Height / 2)
End Sub
Private Function buttonsOut()
    Set cmdLogin.Picture = ImageList1.ListImages(1).Picture
    Set cmdCancel.Picture = ImageList1.ListImages(3).Picture
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
End Sub
