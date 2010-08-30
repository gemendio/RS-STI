VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form MainForm 
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
   Picture         =   "MainForm.frx":0000
   ScaleHeight     =   10380
   ScaleWidth      =   17310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   420
      Left            =   11400
      TabIndex        =   7
      Top             =   2040
      Width           =   4335
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   1560
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Height          =   6015
      Left            =   5040
      TabIndex        =   5
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
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
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   6015
      Left            =   5040
      TabIndex        =   4
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
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
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   6015
      Left            =   5040
      TabIndex        =   3
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
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
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   6015
      Left            =   5040
      TabIndex        =   2
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6015
      Left            =   5040
      TabIndex        =   0
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
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
   Begin ComctlLib.ImageList winButtonsImg 
      Left            =   7920
      Top             =   1920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   107
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":3F5FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":40B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":42060
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":43746
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H000000FF&
      Height          =   615
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
      Height          =   825
      Left            =   240
      Picture         =   "MainForm.frx":44E2C
      Top             =   9360
      Width           =   2190
   End
   Begin VB.Image Image12 
      Height          =   1245
      Left            =   0
      Picture         =   "MainForm.frx":46650
      Top             =   7080
      Width           =   4455
   End
   Begin VB.Image Image11 
      Height          =   735
      Left            =   15960
      Picture         =   "MainForm.frx":49D07
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image10 
      Height          =   735
      Left            =   11280
      Picture         =   "MainForm.frx":4A4F6
      Top             =   1920
      Width           =   4635
   End
   Begin VB.Image Image9 
      Height          =   735
      Left            =   10080
      Picture         =   "MainForm.frx":4A86E
      Top             =   1905
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   735
      Left            =   9360
      Picture         =   "MainForm.frx":4AF3E
      Top             =   1900
      Width           =   600
   End
   Begin VB.Image Image7 
      Height          =   690
      Left            =   8640
      Picture         =   "MainForm.frx":4B6AA
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image Image6 
      Height          =   1305
      Left            =   0
      Picture         =   "MainForm.frx":4BD5E
      Top             =   6120
      Width           =   4485
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   0
      Picture         =   "MainForm.frx":4F5FA
      Top             =   5160
      Width           =   4485
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   0
      Picture         =   "MainForm.frx":528B4
      Top             =   4200
      Width           =   4485
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   8160
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
            Picture         =   "MainForm.frx":557A1
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":5ABA3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":5FFA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":653A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6A7A9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6F953
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":74AFD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7A157
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7F7B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":861F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":8CC3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":931D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":99769
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9BB87
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9DFA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9EBDF
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9F819
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A0413
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A100D
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A1CCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A2989
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A37CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":A4615
      Top             =   3240
      Width           =   4485
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":A73E5
      Top             =   2280
      Width           =   4485
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
    Set Image11.Picture = ImageList1.ListImages(21).Picture
    Set Image12.Picture = ImageList1.ListImages(11).Picture
    Set Image13.Picture = ImageList1.ListImages(13).Picture
End Function

Private Function deployTable()
    Dim currenTab As String

    currentTab = Label1.Caption
    
    Select Case currentTab
            Case "Rooms":
                            Dim room As New ModelRoom
                            Set DataGrid1.DataSource = room.GetAll
            Case "Subjects":
                            Dim subject As New ModelSubject
                            Set DataGrid2.DataSource = subject.GetAll

            Case "Sections":
                            Dim section As New ModelSection
                            Set DataGrid4.DataSource = section.GetAll

            Case "Users":
                            Dim user As New ModelUser
                            Set DataGrid3.DataSource = user.GetAll

            Case "Schedules":
                            Dim schedule As New ModelSchedule
                            Set DataGrid5.DataSource = schedule.GetAll
            
            Case Else: Label2.Caption = ""
        End Select
    'DataGrid1
End Function

Private Sub Form_Load()
    msgFadeout
    deployTable
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
End Sub

Private Sub Image11_Click()
    goSearch
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image11.Picture = ImageList1.ListImages(22).Picture
End Sub

Private Sub Image12_Click()
    Label1.Caption = "Reports"
    
    DataGrid1.Visible = False
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False
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
    
    DataGrid1.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False
    deployTable
    'Adodc1.RecordSource = "SELECT id as ID, name as 'Room Name' FROM rooms"
   ' Adodc1.Refresh

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image2.Picture = ImageList1.ListImages(2).Picture
End Sub

Private Sub Image3_Click()
    Label1.Caption = "Subjects"
    
    DataGrid2.Visible = True
    
    DataGrid1.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False

    deployTable
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image3.Picture = ImageList1.ListImages(4).Picture
End Sub

Private Sub Image4_Click()
    Label1.Caption = "Users"
    
    DataGrid3.Visible = True
    
    DataGrid2.Visible = False
    DataGrid1.Visible = False
    DataGrid4.Visible = False
    DataGrid5.Visible = False

    deployTable
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image4.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub Image5_Click()
    Dim sqlScript As String
    Label1.Caption = "Sections"
    
    DataGrid4.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid1.Visible = False
    DataGrid5.Visible = False
    
    deployTable
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image5.Picture = ImageList1.ListImages(8).Picture
End Sub

Private Sub Image6_Click()
    Dim sqlScript As String
    Label1.Caption = "Schedules"
    
    DataGrid5.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid1.Visible = False
    
    sqlScript = "SELECT scd.id as ID, scd.day as Day,scd.start_time as 'Start Time',scd.end_time as 'End Time', "
    sqlScript = sqlScript + "r.name as Room,sbj.name as Subject,sec.name as Section "
    sqlScript = sqlScript + "FROM schedules scd JOIN rooms r ON r.id=scd.room_id "
    sqlScript = sqlScript + "JOIN sections sec ON sec.id=scd.section_id "
    sqlScript = sqlScript + "JOIN subjects sbj ON sbj.id=scd.subject_id "
    
    deployTable
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image6.Picture = ImageList1.ListImages(10).Picture
End Sub

Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image12.Picture = ImageList1.ListImages(12).Picture
End Sub
Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image13.Picture = ImageList1.ListImages(14).Picture
End Sub

Private Sub Image7_Click()
    goAdd
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
     buttonsOut
    Set Image7.Picture = ImageList1.ListImages(16).Picture
End Sub

Private Sub Image8_Click()
    goEdit
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image8.Picture = ImageList1.ListImages(18).Picture
End Sub

Private Sub Image9_Click()
    goDel
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    buttonsOut
    Set Image9.Picture = ImageList1.ListImages(20).Picture
End Sub

Private Sub SearchStr_Change()
    goSearch
End Sub

Private Sub SearchStr_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then KeyAscii = 0: goSearch
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
                            'Set Adodc5.Recordset = Adodc1.Recordset
                            LoadRoom ("Add")
            Case "Subjects":
                            'Set Adodc5.Recordset = Adodc2.Recordset
                            LoadSubject ("Add")
            Case "Sections":
                            'Set Adodc5.Recordset = Adodc3.Recordset
                            LoadSection ("Add")
            Case "Users":
                            'Set Adodc5.Recordset = Adodc4.Recordset
                            LoadUser ("Add")
            Case "Schedules":
                            'Set Adodc5.Recordset = Adodc6.Recordset
                            LoadSchedule ("Add")
            Case Else: Label2.Caption = ""
        End Select
End Function
Private Function goEdit()
    Dim currentTab As String
    currentTab = Label1.Caption
    Label2.Caption = "Edit"
        
    Select Case currentTab
            Case "Rooms":
                            'Set Adodc5.Recordset = Adodc1.Recordset
                            LoadRoom ("Edit")
            Case "Subjects":
                            'Set Adodc5.Recordset = Adodc2.Recordset
                            LoadSubject ("Edit")
            Case "Sections":
                            'Set Adodc5.Recordset = Adodc3.Recordset
                            LoadSection ("Edit")
            Case "Users":
                            'Set Adodc5.Recordset = Adodc4.Recordset
                            LoadUser ("Edit")
            Case "Schedules":
                            'Set Adodc5.Recordset = Adodc6.Recordset
                            LoadSchedule ("Edit")
            Case Else: Label2.Caption = ""
        End Select
End Function

Private Function goDel()
    On Error GoTo ErrFound
    
    Dim currentTab As String
    currentTab = Label1.Caption
         
    Dim tmp As Integer
    
    tmp = MsgBox("Are you sure to delete this record?", vbYesNo, "Delete Record?")
    
    If tmp = 7 Then
        Label2.Caption = "Deleting record was suspended."
        Label2.Visible = True
        Exit Function
    End If
    
    Select Case currentTab
            Case "Rooms":
                           ' Set Adodc5.Recordset = Adodc1.Recordset
                           ' Adodc5.Recordset!LName = "null"
                           ' Adodc5.Recordset.Update
                           ' Adodc1.Refresh
                           Label2.Caption = "A room record was deleted."
                           Label2.Visible = True
                            
            Case "Subjects":
                           ' Set Adodc5.Recordset = Adodc2.Recordset
                           ' Adodc5.Recordset!LName = "null"
                           ' Adodc5.Recordset.Update
                           ' Adodc2.Refresh
                           Label2.Caption = "A subject record was deleted."
                           Label2.Visible = True
                           
            Case "Sections":
                           ' Set Adodc5.Recordset = Adodc3.Recordset
                           ' Adodc5.Recordset!LoanName = "null"
                           ' Adodc5.Recordset.Update
                           ' Adodc3.Refresh
                           Label2.Caption = "A section record was deleted."
                           Label2.Visible = True
                           
            Case "Users":
                           ' Set Adodc5.Recordset = Adodc4.Recordset
                           ' Adodc5.Recordset!TermType = "null"
                           ' Adodc5.Recordset.Update
                           ' Adodc4.Refresh
                           Label2.Caption = "A user record was deleted."
                           Label2.Visible = True
                           
            Case "Schedules":
                           ' Set Adodc5.Recordset = Adodc6.Recordset
                           ' Adodc5.Recordset!Name = "null"
                           ' Adodc5.Recordset.Update
                           ' Adodc6.Refresh
                           Label2.Caption = "A schedule record was deleted."
                           Label2.Visible = True
            Case Else: Label2.Caption = ""
        End Select
    Exit Function

ErrFound:

    'Adodc5.Recordset.Cancel
    'Adodc5.Refresh

    MsgBox "Error Number : " & Err.Number & _
    "Error Description : " & _
    Err.Description, vbInformation, Err.Source

Resume

End Function
Private Function goSearch()
    Dim currentTab As String
    Dim strSeek As String
    Dim sqlScript As String

    strSeek = searchStr.Text
    currentTab = Label1.Caption
        
        Select Case currentTab
            Case "Rooms":
                         Dim room As New ModelRoom
                         Set DataGrid1.DataSource = room.search(strSeek)
                         DataGrid1.Refresh
            Case "Sections":
                         Dim section As New ModelSection
                         Set DataGrid4.DataSource = section.search(strSeek)
                         DataGrid4.Refresh
            Case "Subjects":
                         Dim subject As New ModelSubject
                         Set DataGrid2.DataSource = subject.search(strSeek)
                         DataGrid2.Refresh
                         
            Case "Users":
                         Dim user As New ModelUser
                         Set DataGrid3.DataSource = user.search(strSeek)
                         DataGrid3.Refresh
                         
            Case "Schedules":
                         Dim schedule As New ModelSchedule
                         Set DataGrid5.DataSource = schedule.search(strSeek)
                         DataGrid5.Refresh

            Case Else: Label2.Caption = ""
        End Select

End Function

Private Function LoadRoom(args As String)
    Load room
    room.Show vbModal
End Function
Private Function LoadUser(args As String)
    Load user
    user.Show vbModal
End Function
Private Function LoadSubject(args As String)
    Load subject
    subject.Show vbModal
End Function
Private Function LoadSection(args As String)
    Load section
    section.Show vbModal
End Function
Private Function LoadSchedule(args As String)
    Load schedule
    schedule.Show vbModal
End Function





