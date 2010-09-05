VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form MainForm 
   BackColor       =   &H80000010&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   10380
   ClientLeft      =   90
   ClientTop       =   375
   ClientWidth     =   17310
   BeginProperty Font 
      Name            =   "Bernard MT Condensed"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   17310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid schedule_grid 
      Height          =   5850
      Left            =   5040
      TabIndex        =   5
      Top             =   2920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10319
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox searchStr 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12135
      TabIndex        =   7
      Top             =   2050
      Width           =   3720
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   1560
   End
   Begin MSDataGridLib.DataGrid section_grid 
      Height          =   5850
      Left            =   5040
      TabIndex        =   4
      Top             =   2920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10319
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid user_grid 
      Height          =   5850
      Left            =   5040
      TabIndex        =   3
      Top             =   2920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10319
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid subject_grid 
      Height          =   5850
      Left            =   5040
      TabIndex        =   2
      Top             =   2920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10319
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid room_grid 
      Height          =   5850
      Left            =   5040
      TabIndex        =   0
      Top             =   2920
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10319
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Image section_user 
      Height          =   750
      Left            =   7320
      Picture         =   "MainForm.frx":442C7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1005
   End
   Begin ComctlLib.ImageList winButtonsImg 
      Left            =   12000
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   107
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   14
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":44A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":45FC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":474F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":48BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":4A2BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":4B7D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":4CCE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":4DE7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":4F016
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":4F87C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":500E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":509B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":51286
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":51B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   495
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Schedules"
      BeginProperty Font 
         Name            =   "Maiandra GD"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   675
      Left            =   5160
      TabIndex        =   1
      Top             =   1995
      Width           =   2415
   End
   Begin VB.Image Image13 
      Height          =   945
      Left            =   120
      Picture         =   "MainForm.frx":523E2
      Top             =   9345
      Width           =   975
   End
   Begin VB.Image Image12 
      Height          =   1245
      Left            =   0
      Picture         =   "MainForm.frx":532AB
      Top             =   7150
      Width           =   4470
   End
   Begin VB.Image Image10 
      Height          =   720
      Left            =   11880
      Picture         =   "MainForm.frx":5682B
      Top             =   1920
      Width           =   4740
   End
   Begin VB.Image Image9 
      Height          =   750
      Left            =   11280
      Picture         =   "MainForm.frx":5712A
      Top             =   1905
      Width           =   510
   End
   Begin VB.Image Image8 
      Height          =   750
      Left            =   10560
      Picture         =   "MainForm.frx":578E9
      Top             =   1905
      Width           =   645
   End
   Begin VB.Image Image7 
      Height          =   750
      Left            =   9840
      Picture         =   "MainForm.frx":580FD
      Top             =   1920
      Width           =   570
   End
   Begin VB.Image Image6 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":588F9
      Top             =   6170
      Width           =   4470
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   0
      Picture         =   "MainForm.frx":5B6F2
      Top             =   5145
      Width           =   4470
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   0
      Picture         =   "MainForm.frx":5E8F3
      Top             =   4200
      Width           =   4470
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   960
      Top             =   8760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   299
      ImageHeight     =   68
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   22
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":615BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":671F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6CE26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":72228
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7762A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7C7D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":8197E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":86FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":8C632
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":91A34
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":96E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9D3CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A3962
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A4E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A637E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A6FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A7BC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A88AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A9596
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":AA0F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":AAC4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":ABA90
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":AC8D6
      Top             =   3240
      Width           =   4470
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   0
      Picture         =   "MainForm.frx":AF54F
      Top             =   2160
      Width           =   4470
   End
   Begin VB.Image Image1 
      Height          =   15000
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function msgFadeout()
    Timer1.Interval = 5000
    Timer1.Enabled = True
End Function

Private Function buttonsOut()
    Set Image2.Picture = ImageList1.ListImages(1).Picture
    Set Image3.Picture = ImageList1.ListImages(3).Picture
    Set Image4.Picture = ImageList1.ListImages(5).Picture
    Set Image5.Picture = ImageList1.ListImages(7).Picture
    Set Image6.Picture = ImageList1.ListImages(9).Picture
    Set Image7.Picture = ImageList1.ListImages(15).Picture
    Set Image8.Picture = ImageList1.ListImages(17).Picture
    Set Image9.Picture = ImageList1.ListImages(19).Picture
    Set Image12.Picture = ImageList1.ListImages(11).Picture
    Set Image13.Picture = ImageList1.ListImages(13).Picture
    Set section_user.Picture = winButtonsImg.ListImages(7).Picture
End Function

Private Function deployTable()
    Dim currenTab As String

    currentTab = Label1.Caption
    
    Select Case currentTab
            Case "Rooms":
                            Dim room As New Model.room
                            Set room_grid.DataSource = room.GetAll
            Case "Subjects":
                            Dim subject As New Model.subject
                            Set subject_grid.DataSource = subject.GetAll

            Case "Sections":
                            Dim section As New Model.section
                            Set section_grid.DataSource = section.GetAll

            Case "Users":
                            Dim user As New Model.user
                            Set user_grid.DataSource = user.GetAll

            Case "Schedules":
                            Dim schedule As New Model.schedule
                            Set schedule_grid.DataSource = schedule.GetAll
            
            Case Else: Label2.Caption = ""
        End Select
End Function

Private Sub Form_Load()
    msgFadeout
    deployTable
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
End Sub

Private Sub Image11_Click()
    goSearch
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image11.Picture = ImageList1.ListImages(22).Picture
End Sub

Private Sub Image12_Click()
    Label1.Caption = "Reports"
    
    room_grid.Visible = False
    subject_grid.Visible = False
    user_grid.Visible = False
    section_grid.Visible = False
    schedule_grid.Visible = False
End Sub

Private Sub Image13_Click()
    If MsgBox("Are you sure you want to quit?", _
        vbYesNo + vbQuestion, _
        "Exit Room Scheduler") = vbNo Then
        Cancel = 1
    Else
        Unload Me
    End If
    
End Sub

Private Sub Image2_Click()
    Label1.Caption = "Rooms"
    
    room_grid.Visible = True
    
    subject_grid.Visible = False
    user_grid.Visible = False
    section_grid.Visible = False
    schedule_grid.Visible = False
    deployTable

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image2.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub Image3_Click()
    Label1.Caption = "Subjects"
    
    subject_grid.Visible = True
    
    room_grid.Visible = False
    user_grid.Visible = False
    section_grid.Visible = False
    schedule_grid.Visible = False

    deployTable
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image3.Picture = ImageList1.ListImages(4).Picture
End Sub

Private Sub Image4_Click()
    Label1.Caption = "Users"
    
    user_grid.Visible = True
    
    subject_grid.Visible = False
    room_grid.Visible = False
    section_grid.Visible = False
    schedule_grid.Visible = False

    deployTable
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image4.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub Image5_Click()
    Dim sqlScript As String
    Label1.Caption = "Sections"
    section_user.Visible = True
    
    section_grid.Visible = True
    
    subject_grid.Visible = False
    user_grid.Visible = False
    room_grid.Visible = False
    schedule_grid.Visible = False
    
    deployTable
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image5.Picture = ImageList1.ListImages(8).Picture
End Sub

Private Sub Image6_Click()
    Dim sqlScript As String
    Label1.Caption = "Schedules"
    
    schedule_grid.Visible = True
    
    subject_grid.Visible = False
    user_grid.Visible = False
    section_grid.Visible = False
    room_grid.Visible = False
    
    deployTable
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image6.Picture = ImageList1.ListImages(10).Picture
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image12.Picture = ImageList1.ListImages(12).Picture
End Sub
Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image13.Picture = ImageList1.ListImages(14).Picture
End Sub

Private Sub Image7_Click()
    goAdd
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
     buttonsOut
    Set Image7.Picture = ImageList1.ListImages(16).Picture
End Sub

Private Sub Image8_Click()
    goEdit
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image8.Picture = ImageList1.ListImages(18).Picture
End Sub

Private Sub Image9_Click()
    goDel
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image9.Picture = ImageList1.ListImages(20).Picture
End Sub

Private Sub SearchStr_Change()
    goSearch
End Sub

Private Sub SearchStr_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then KeyAscii = 0: goSearch
End Sub

Private Sub section_user_Click()
    Load secUsers
    secUsers.Show vbModal
End Sub

Private Sub section_user_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set section_user.Picture = winButtonsImg.ListImages(8).Picture
End Sub

Private Sub Timer1_Timer()
    Label2.Visible = False
End Sub
Private Function goAdd()
    Dim currentTab As String
    currentTab = Label1.Caption
    Label2.Caption = "Add"
    
    Select Case currentTab
            Case "Rooms":
                            Load room
                            room.Show vbModal
            Case "Subjects":
                            Load subject
                            subject.Show vbModal
            Case "Sections":
                            Load section
                            section.Show vbModal
            Case "Users":
                            Load user
                            user.Show vbModal
            Case "Schedules":
                            Load schedule
                            schedule.Show vbModal
            Case Else: Label2.Caption = ""
        End Select
        deployTable
End Function
Private Function goEdit()
    Dim currentTab As String
    currentTab = Label1.Caption
    Label2.Caption = "Edit"
        
    Select Case currentTab
            Case "Rooms":
                            Load room
                            room.Show vbModal
            Case "Subjects":
                            Load subject
                            subject.Show vbModal
            Case "Sections":
                            Load section
                            section.Show vbModal
            Case "Users":
                            Load user
                            user.Show vbModal
            Case "Schedules":
                            Load schedule
                            schedule.Show vbModal
            Case Else: Label2.Caption = ""
        End Select
        deployTable
End Function

Private Function goDel()
Dim X As String
    On Error GoTo ErrFound
    
    Dim currentTab As String
    Dim tmp As Integer
    
    currentTab = Label1.Caption
    
    tmp = MsgBox("Are you sure to delete this record?", vbYesNo, "Delete Record?")
    
    If tmp = 7 Then
        Label2.Caption = "Deleting record was suspended."
        Label2.Visible = True
        Exit Function
    End If
    
    Select Case currentTab
            Case "Rooms":
                           Dim room As New Model.room
                           room.Load (room_grid.Columns("ID"))
                           room.Delete
                           
                           Label2.Caption = "A room record was deleted."
                           Label2.Visible = True
                            
            Case "Subjects":
                           Dim subject As New Model.subject
                           subject.Load (subject_grid.Columns("ID"))
                           subject.Delete
                           
                           Label2.Caption = "A subject record was deleted."
                           Label2.Visible = True
                           
            Case "Sections":
                           Dim section As New Model.section
                           section.Load (section_grid.Columns("ID"))
                           section.Delete
                           
                           Label2.Caption = "A section record was deleted."
                           Label2.Visible = True
                           
            Case "Users":
                           Dim user As New Model.user
                           user.Load (user_grid.Columns("ID"))
                           user.Delete
                           
                           Label2.Caption = "A user record was deleted."
                           Label2.Visible = True
                           
            Case "Schedules":
                           Dim schedule As New Model.schedule
                           schedule.Load (schedule_grid.Columns("ID"))
                           schedule.Delete
                           
                           Label2.Caption = "A schedule record was deleted."
                           Label2.Visible = True
            Case Else: Label2.Caption = ""
        End Select
        deployTable
    Exit Function

ErrFound:

    MsgBox "Error Number : " & Err.Number & _
    "Error Description : " & _
    Err.Description, vbInformation, Err.Source

Resume

End Function
Private Function goSearch()
    Dim currentTab As String
    Dim strSeek As String

    strSeek = searchStr.Text
    currentTab = Label1.Caption
        
        Select Case currentTab
            Case "Rooms":
                         Dim room As New Model.room
                         Set room_grid.DataSource = room.search(strSeek)
                         room_grid.Refresh
            Case "Sections":
                         Dim section As New Model.section
                         Set section_grid.DataSource = section.search(strSeek)
                         section_grid.Refresh
            Case "Subjects":
                         Dim subject As New Model.subject
                         Set subject_grid.DataSource = subject.search(strSeek)
                         subject_grid.Refresh
                         
            Case "Users":
                         Dim user As New Model.user
                         Set user_grid.DataSource = user.search(strSeek)
                         user_grid.Refresh
                         
            Case "Schedules":
                         Dim schedule As New Model.schedule
                         Set schedule_grid.DataSource = schedule.search(strSeek)
                         schedule_grid.Refresh

            Case Else: Label2.Caption = ""
        End Select

End Function

