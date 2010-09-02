VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Login to Room Scheduler"
   ClientHeight    =   2775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox txtPwd 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "p@ssword"
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Text            =   "root"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtDB 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Text            =   "room_scheduler"
      Top             =   840
      Width           =   2535
   End
   Begin VB.TextBox txtServer 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "localhost"
      Top             =   480
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Connection Values"
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3855
      Begin VB.Label lblPassword 
         Caption         =   "Password:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblUsername 
         Caption         =   "Username:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblDatabase 
         Caption         =   "Database:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblServer 
         Caption         =   "Server:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
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

Private Sub Form_Load()
    frmLogin.Left = (Screen.Width / 2) - (frmLogin.Width / 2)
    frmLogin.Top = (Screen.Height / 2) - (frmLogin.Height / 2)
End Sub
