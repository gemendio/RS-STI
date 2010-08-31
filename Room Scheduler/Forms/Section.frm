VERSION 5.00
Begin VB.Form Section 
   Caption         =   "Form1"
   ClientHeight    =   7665
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Section.frx":0000
   ScaleHeight     =   7665
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox sec_user_name 
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
      ItemData        =   "Section.frx":1790C
      Left            =   2760
      List            =   "Section.frx":1790E
      TabIndex        =   2
      Top             =   2200
      Width           =   3625
   End
   Begin VB.TextBox section_name 
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
      Left            =   2880
      TabIndex        =   1
      Top             =   1520
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
      Width           =   2415
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2520
      Picture         =   "Section.frx":17910
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "Section.frx":18662
      Top             =   6600
      Width           =   1605
   End
End
Attribute VB_Name = "Section"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private state As String
Private section_grid As DataGrid
Private Id As Integer
Private section As New ModelSection
Private user As New ModelUser

Public Function goEdit()
    section.Name = section_name.Text
    section.UserId = sec_user_name.ListIndex
    section.Upsert
End Function

Public Function goAdd()
    Set section_grid = MainForm.section_grid
End Function
Private Function LoadUsers()
    Dim rs_users As New ADODB.Recordset
    
    Set rs_users = user.GetAll
    sec_user_name.Clear
    Do While Not rs_users.EOF
        sec_user_name.AddItem rs_users.fields("Last Name") & ", " & rs_users.fields("First Name")
        sec_user_name.ItemData(sec_user_name.NewIndex) = rs_users.fields("ID")
        rs_users.MoveNext
    Loop
End Function


Private Sub Form_Load()
    state = MainForm.Label2.Caption
    Label1.Caption = state + " Section"
    Me.Caption = state + " Schedule"

    LoadUsers
    
    'goValidate
    
    If state = "Edit" Then
         Set section_grid = MainForm.section_grid
         
         Id = section_grid.Bookmark
         section.Load (Id)

         Me.section_name.Text = section.Name
         Me.sec_user_name.ListIndex = section.UserId - 1

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
    MainForm.Label2.Caption = state + "ing section was canceled."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    If state = "Edit" Then
        goEdit
    Else
        goAdd
    End If
    MainForm.Label2.Caption = state + "ing section was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub


