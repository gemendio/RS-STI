VERSION 5.00
Begin VB.Form User 
   Caption         =   "User"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "User.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox user_type 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "User.frx":1AFFD
      Left            =   2760
      List            =   "User.frx":1B007
      TabIndex        =   4
      Top             =   3640
      Width           =   3625
   End
   Begin VB.TextBox last_name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   370
      Left            =   2790
      TabIndex        =   3
      Top             =   2930
      Width           =   3490
   End
   Begin VB.TextBox middle_name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   370
      Left            =   2790
      TabIndex        =   2
      Top             =   2200
      Width           =   3490
   End
   Begin VB.TextBox first_name 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   370
      Left            =   2790
      TabIndex        =   1
      Top             =   1510
      Width           =   3490
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   280
      Width           =   2775
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2520
      Picture         =   "User.frx":1B020
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "User.frx":1BD72
      Top             =   6600
      Width           =   1605
   End
End
Attribute VB_Name = "User"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private state As String
Private user_grid As DataGrid
Private Id As Integer
Private user As New Model.user

Private Sub Form_Load()
    state = MainForm.Label2.Caption
    Label1.Caption = state + " User"
    Me.Caption = state + " User"
    Set user_grid = MainForm.user_grid
    
    If state = "Edit" Then
         Id = user_grid.Columns("ID")
         user.Load (Id)
         first_name.Text = user.FirstName
         middle_name.Text = user.MiddleName
         last_name.Text = user.LastName
         user_type.Text = user.UserType
    Else
        user.Id = 0
    End If
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set Image1.Picture = MainForm.winButtonsImg.ListImages(2).Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set Image1.Picture = MainForm.winButtonsImg.ListImages(1).Picture
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set Image2.Picture = MainForm.winButtonsImg.ListImages(4).Picture
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Set Image2.Picture = MainForm.winButtonsImg.ListImages(3).Picture
End Sub
Private Sub Image2_Click()
    MainForm.Label2.Caption = state + "ing user was canceled."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    user.FirstName = first_name.Text
    user.MiddleName = middle_name.Text
    user.LastName = last_name.Text
    user.UserType = user_type.Text
    user.Upsert
    
    MainForm.Label2.Caption = state + "ing user was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

