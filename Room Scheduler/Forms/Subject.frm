VERSION 5.00
Begin VB.Form Subject 
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   Picture         =   "Subject.frx":0000
   ScaleHeight     =   7620
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox subject_name 
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
      Height          =   380
      Left            =   2780
      TabIndex        =   1
      Top             =   1500
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
      Left            =   4320
      TabIndex        =   0
      Top             =   315
      Width           =   2535
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2520
      Picture         =   "Subject.frx":16607
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "Subject.frx":17359
      Top             =   6600
      Width           =   1605
   End
End
Attribute VB_Name = "Subject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private state As String
Private subject_grid As DataGrid
Private Id As Integer
Private subject As New ModelSubject

Public Function goEdit()
    subject.Name = subject_name.Text
    subject.Upsert
End Function
Public Function goAdd()
    Set subject_grid = MainForm.subject_grid
End Function
Private Sub Form_Load()
    state = MainForm.Label2.Caption
    Label1.Caption = state + " Subject"
    Me.Caption = state + " Subject"
    
    If state = "Edit" Then
         Set subject_grid = MainForm.subject_grid
         Id = subject_grid.Bookmark
         subject.Load (Id)
         subject_name.Text = subject.Name
    End If
End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Image1.Picture = MainForm.winButtonsImg.ListImages(2).Picture
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Image1.Picture = MainForm.winButtonsImg.ListImages(1).Picture
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Image2.Picture = MainForm.winButtonsImg.ListImages(4).Picture
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Image2.Picture = MainForm.winButtonsImg.ListImages(3).Picture
End Sub
Private Sub Image2_Click()
    MainForm.Label2.Caption = state + "ing subject was canceled."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    If state = "Edit" Then
        goEdit
    Else
        goAdd
    End If
    
    MainForm.Label2.Caption = state + "ing subject was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub


