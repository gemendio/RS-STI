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
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   4730
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   1200
      Width           =   2775
   End
   Begin VB.ListBox secStudList 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "secUsers.frx":11380
      Left            =   720
      List            =   "secUsers.frx":11382
      MultiSelect     =   2  'Extended
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
      Picture         =   "secUsers.frx":11384
      Top             =   130
      Width           =   540
   End
   Begin VB.Image right_arrow 
      Height          =   480
      Left            =   3840
      Picture         =   "secUsers.frx":11A29
      Top             =   3120
      Width           =   510
   End
   Begin VB.Image left_arrow 
      Height          =   435
      Left            =   3840
      Picture         =   "secUsers.frx":11ED1
      Top             =   2400
      Width           =   510
   End
End
Attribute VB_Name = "secUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private section_user As New ModelSectionUser
Private m_section As New ModelSection
Private m_user As New ModelUser


Private Sub Form_Load()
    state = MainForm.Label2.Caption
        
    If state = "Edit" Then
         m_section.Load (section.sectionIdTxt.Text)
         LoadSectionStudents
    End If
         LoadUsers
End Sub

Private Sub exit_sec_user_Click()
    section.secStudentList.Clear
    Dim i As Integer

    For i = 0 To secStudList.ListCount - 1
     section.secStudentList.AddItem secStudList.List(i)
     section.secStudentList.ItemData(section.secStudentList.NewIndex) = secStudList.ItemData(i)
    Next i
    
    Unload Me
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set left_arrow.Picture = MainForm.winButtonsImg.ListImages(9).Picture
    Set right_arrow.Picture = MainForm.winButtonsImg.ListImages(11).Picture
    Set exit_sec_user.Picture = MainForm.winButtonsImg.ListImages(13).Picture
End Sub

Private Sub left_arrow_Click()
    Dim i As Integer

    If allList.ListIndex = -1 Then Exit Sub
        For i = allList.ListCount - 1 To 0 Step -1
            If allList.Selected(i) = True Then
               secStudList.AddItem allList.List(i)
               secStudList.ItemData(secStudList.NewIndex) = allList.ItemData(i)
               allList.RemoveItem i
        End If
    Next i
End Sub

Private Sub left_arrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set left_arrow.Picture = MainForm.winButtonsImg.ListImages(10).Picture
End Sub

Private Sub right_arrow_Click()
    Dim i As Integer

    If secStudList.ListIndex = -1 Then Exit Sub
        For i = secStudList.ListCount - 1 To 0 Step -1
            If secStudList.Selected(i) = True Then
               allList.AddItem secStudList.List(i)
               allList.ItemData(allList.NewIndex) = secStudList.ItemData(i)
               secStudList.RemoveItem i
        End If
    Next i
End Sub

Private Sub right_arrow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set right_arrow.Picture = MainForm.winButtonsImg.ListImages(12).Picture
End Sub
Private Sub exit_sec_user_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Set exit_sec_user.Picture = MainForm.winButtonsImg.ListImages(14).Picture
End Sub
Public Function LoadSectionStudents()
    Dim rs_section As New ADODB.Recordset

    Set rs_section = m_section.getStudents
    secStudList.Clear
    Do While Not rs_section.EOF
        secStudList.AddItem rs_section.fields("Student Name")
        secStudList.ItemData(secStudList.NewIndex) = rs_section.fields("ID")
        rs_section.MoveNext
    Loop
End Function
Public Function LoadUsers()
    Dim rs_user As New ADODB.Recordset
    
    Set rs_user = m_user.GetAll
    allList.Clear
    Do While Not rs_user.EOF
        allList.AddItem (rs_user.fields("Last Name") + ", " + rs_user.fields("First Name"))
        allList.ItemData(allList.NewIndex) = rs_user.fields("ID")
        rs_user.MoveNext
    Loop
End Function
