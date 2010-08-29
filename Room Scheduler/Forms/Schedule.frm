VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Schedule 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   Picture         =   "Schedule.frx":0000
   ScaleHeight     =   7650
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox day_sched 
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
      TabIndex        =   6
      Top             =   1520
      Width           =   3615
   End
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
      TabIndex        =   5
      Top             =   5050
      Width           =   3610
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
      TabIndex        =   4
      Top             =   4350
      Width           =   3615
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
      Left            =   2760
      TabIndex        =   3
      Top             =   3650
      Width           =   3615
   End
   Begin MSComCtl2.DTPicker end_time 
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   2925
      Width           =   3645
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
      CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
      Format          =   126943235
      CurrentDate     =   40418.5
   End
   Begin MSComCtl2.DTPicker start_time 
      Height          =   375
      Left            =   2780
      TabIndex        =   1
      Top             =   2205
      Width           =   3600
      _ExtentX        =   6350
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
      CustomFormat    =   "MM/dd/yyyy hh:mm:ss tt"
      Format          =   126943235
      CurrentDate     =   40418
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
      Height          =   495
      Left            =   4080
      TabIndex        =   0
      Top             =   285
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   2520
      Picture         =   "Schedule.frx":1D31D
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   720
      Picture         =   "Schedule.frx":1E06F
      Top             =   6600
      Width           =   1605
   End
End
Attribute VB_Name = "Schedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim state As String

Private Sub Form_Load()
    state = MainForm.Label2.Caption
    Label1.Caption = state + " Schedule"
    Schedule.Caption = state + " Schedule"
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
    MainForm.Label2.Caption = state + "ing schedule was canceled."
    MainForm.Label2.Visible = True
    Unload Me
End Sub

Private Sub Image1_Click()
    MainForm.Label2.Caption = state + "ing schedule was successful."
    MainForm.Label2.Visible = True
    Unload Me
End Sub


