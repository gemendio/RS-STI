VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Schedule 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "Schedule.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox sec_sched 
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
      Left            =   2770
      TabIndex        =   4
      Top             =   5050
      Width           =   3625
   End
   Begin VB.ComboBox subj_sched 
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
      Left            =   2760
      TabIndex        =   3
      Top             =   4350
      Width           =   3625
   End
   Begin VB.ComboBox room_sched 
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
      ItemData        =   "Schedule.frx":1D498
      Left            =   2760
      List            =   "Schedule.frx":1D49A
      TabIndex        =   2
      Top             =   3650
      Width           =   3625
   End
   Begin MSComCtl2.DTPicker sched_date 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "yyyy/MM/dd"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   3
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1480
      Width           =   3625
      _ExtentX        =   6403
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTrailingForeColor=   7171437
      CustomFormat    =   "yyyy/MM/dd"
      Format          =   16252931
      CurrentDate     =   40418
   End
   Begin MSComCtl2.DTPicker start_time 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   2190
      Width           =   3625
      _ExtentX        =   6403
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTrailingForeColor=   7171437
      CustomFormat    =   "HH:mm:ss"
      Format          =   16252931
      UpDown          =   -1  'True
      CurrentDate     =   40421
   End
   Begin MSComCtl2.DTPicker end_time 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm:ss"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   4
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2925
      Width           =   3640
      _ExtentX        =   6429
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTrailingForeColor=   7171437
      CustomFormat    =   "HH:mm:ss"
      Format          =   16252931
      UpDown          =   -1  'True
      CurrentDate     =   40422
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   285
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2520
      Picture         =   "Schedule.frx":1D49C
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "Schedule.frx":1E1EE
      Top             =   6600
      Width           =   1605
   End
End
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private state As String
Private room As New Model.room
Private section As New Model.section
Private subject As New Model.subject
Private schedule As New Model.schedule

Public Function LoadRooms()
    Dim rs_room As New ADODB.Recordset
    
    Set rs_room = room.GetAll
    room_sched.Clear
    Do While Not rs_room.EOF
        room_sched.AddItem rs_room.fields("Room Name")
        room_sched.ItemData(room_sched.NewIndex) = rs_room.fields("ID")
        rs_room.MoveNext
    Loop
End Function
Public Function LoadSubjects()
    Dim rs_subject As New ADODB.Recordset
    
    Set rs_subject = subject.GetAll
    subj_sched.Clear
    Do While Not rs_subject.EOF
        subj_sched.AddItem rs_subject.fields("Subject Name")
        subj_sched.ItemData(subj_sched.NewIndex) = rs_subject.fields("ID")
        rs_subject.MoveNext
    Loop

End Function
Public Function LoadSections()
    Dim rs_section As New ADODB.Recordset
    
    Set rs_section = section.GetAll
    sec_sched.Clear
    Do While Not rs_section.EOF
        sec_sched.AddItem rs_section.fields("Section Name")
        sec_sched.ItemData(sec_sched.NewIndex) = rs_section.fields("ID")
        rs_section.MoveNext
    Loop
    
End Function

Private Sub Form_Load()
    state = MainForm.Label2.Caption
    
    Label1.Caption = state + " Schedule"
    Me.Caption = state + " Schedule"

    Set schedule_grid = MainForm.schedule_grid
    LoadRooms
    LoadSubjects
    LoadSections
    sched_date.value = Format$(Now(), "yyyy/MM/dd")
    'goValidate
    
    If state = "Edit" Then
         Id = schedule_grid.Columns("ID")
         schedule.Load (Id)
         
         Me.sched_date.value = schedule.Day
         Me.start_time = schedule.StartTime
         Me.end_time = schedule.EndTime
         Me.sec_sched.ListIndex = schedule.SectionId
         Me.subj_sched.ListIndex = schedule.SubjectId - 1
         Me.room_sched.ListIndex = schedule.RoomId - 1
    Else
        schedule.Id = 0
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
    MainForm.Label2.Caption = state + "ing schedule was canceled."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    schedule.Day = Format$(sched_date.value, "yyyy/MM/dd")
    schedule.StartTime = Format$(start_time.value, "h:mm:ss")
    schedule.EndTime = Format$(end_time.value, "h:mm:ss")
    schedule.RoomId = Me.room_sched.ListIndex + 1
    schedule.SubjectId = Me.subj_sched.ListIndex + 1
    schedule.SectionId = Me.sec_sched.ListIndex + 1
    schedule.Upsert
    
    MainForm.Label2.Caption = state + "ing schedule was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

