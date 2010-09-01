VERSION 5.00
Begin VB.Form Room 
   BackColor       =   &H8000000B&
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7020
   FillColor       =   &H00FFFFFF&
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Room.frx":0000
   ScaleHeight     =   7635
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox room_name 
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
      Left            =   2800
      TabIndex        =   1
      Top             =   1510
      Width           =   3480
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
      Left            =   4560
      TabIndex        =   0
      Top             =   315
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2520
      Picture         =   "Room.frx":16655
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "Room.frx":173A7
      Top             =   6600
      Width           =   1605
   End
End
Attribute VB_Name = "Room"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private state As String
Private room_grid As DataGrid
Private Id As Integer
Private room As New ModelRoom

Private Sub Form_Load()
    state = MainForm.Label2.Caption
    Label1.Caption = state + " Room"
    Me.Caption = state + " Room"
    Set room_grid = MainForm.room_grid
    
    If state = "Edit" Then
         Id = room_grid.Columns("ID")
         room.Load (Id)
         room_name.Text = room.Name
    Else
        room.Id = 0
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
    MainForm.Label2.Caption = state + "ing room was canceled."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    room.Name = room_name.Text
    room.Upsert
    
    MainForm.Label2.Caption = state + "ing room was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

