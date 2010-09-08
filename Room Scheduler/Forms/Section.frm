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
   Begin VB.TextBox sectionIdTxt 
      Height          =   375
      Left            =   5280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox secStudentList 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   1320
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   2760
      Width           =   4455
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
      ForeColor       =   &H80000007&
      Height          =   370
      Left            =   2880
      TabIndex        =   1
      Top             =   1520
      Width           =   3480
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
      Left            =   2520
      TabIndex        =   3
      Top             =   2070
      Width           =   2535
   End
   Begin VB.Image Image3 
      Height          =   1005
      Left            =   530
      Picture         =   "Section.frx":187E7
      Top             =   1860
      Width           =   945
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
      Picture         =   "Section.frx":19083
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "Section.frx":19DD5
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
Private m_section As New ModelSection
Private section_user As New ModelSectionUser

Private Sub Form_Load()
    state = MainForm.Label2.Caption
    Label1.Caption = state + " Section"
    Me.Caption = state + " Section"
    Set section_grid = MainForm.section_grid
    
    'goValidate
    If state = "Edit" Then
         Id = section_grid.Columns("ID")
         m_section.Load (Id)
         Me.section_name.Text = m_section.Name
         LoadSectionStudents
    Else
        m_section.Id = 0
    End If
    
    Me.sectionIdTxt.Text = m_section.Id

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Image3.Picture = MainForm.winButtonsImg.ListImages(5).Picture
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

Private Sub Image3_Click()
    Load secUsers
    secUsers.Show vbModal
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set Image3.Picture = MainForm.winButtonsImg.ListImages(6).Picture
End Sub

Private Sub Image2_Click()
    MainForm.Label2.Caption = state + "ing section was canceled."
    MainForm.Label2.ForeColor = &HC0&
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    Dim i As Integer
    Dim rs_secUsr As New ADODB.Recordset
    
    m_section.Name = section_name.Text
    m_section.Upsert
     
    section_user.clearSectionUser m_section.Id
    
    For i = 0 To secStudentList.ListCount - 1
            section_user.SectionId = m_section.Id
            section_user.UserId = secStudentList.ItemData(i)
            section_user.Upsert
    Next i
    
    MainForm.Label2.Caption = state + "ing section was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub
Public Function LoadSectionStudents()
    Dim rs_section As New ADODB.Recordset
    
    Set rs_section = m_section.getStudents()
    secStudentList.Clear
    Do While Not rs_section.EOF
        secStudentList.AddItem rs_section.fields("Student Name")
        secStudentList.ItemData(secStudentList.NewIndex) = rs_section.fields("ID")
        rs_section.MoveNext
    Loop
End Function

