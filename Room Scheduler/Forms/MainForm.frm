VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
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
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   1560
   End
   Begin MSAdodcLib.Adodc Adodc5 
      Height          =   375
      Left            =   3000
      Top             =   9480
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   $"MainForm.frx":3F5FC
      Caption         =   "Adodc5"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Bindings        =   "MainForm.frx":3F73C
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
      FormatLocked    =   -1  'True
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Day"
         Caption         =   "Day"
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
      BeginProperty Column02 
         DataField       =   "Start Time"
         Caption         =   "Start Time"
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
      BeginProperty Column03 
         DataField       =   "End Time"
         Caption         =   "End Time"
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
      BeginProperty Column04 
         DataField       =   "Room"
         Caption         =   "Room"
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
      BeginProperty Column05 
         DataField       =   "Subject"
         Caption         =   "Subject"
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
      BeginProperty Column06 
         DataField       =   "Section"
         Caption         =   "Section"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1065.26
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc4 
      Height          =   330
      Left            =   3000
      Top             =   9120
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   $"MainForm.frx":3F751
      Caption         =   "Adodc4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Bindings        =   "MainForm.frx":3F7F6
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
      FormatLocked    =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Section Name"
         Caption         =   "Section Name"
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
      BeginProperty Column02 
         DataField       =   "User"
         Caption         =   "User"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   2880
      Top             =   8760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   "SELECT id as ID, last_name as 'Last Name', first_name as 'First Name', middle_name as 'Middle Name', type as Type FROM users"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Bindings        =   "MainForm.frx":3F80B
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
      FormatLocked    =   -1  'True
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Last Name"
         Caption         =   "Last Name"
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
      BeginProperty Column02 
         DataField       =   "First Name"
         Caption         =   "First Name"
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
      BeginProperty Column03 
         DataField       =   "Middle Name"
         Caption         =   "Middle Name"
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
      BeginProperty Column04 
         DataField       =   "Type"
         Caption         =   "Type"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1140.095
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2880
      Top             =   8280
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   "SELECT id as ID, name as 'Subject Name' FROM subjects"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Bindings        =   "MainForm.frx":3F820
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
      FormatLocked    =   -1  'True
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
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Subject Name"
         Caption         =   "Subject Name"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2760
      Top             =   7920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=STI"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "STI"
      OtherAttributes =   ""
      UserName        =   "root"
      Password        =   "amirah@1"
      RecordSource    =   "SELECT id as ID, name as 'Subject Name' FROM rooms"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Bernard MT Condensed"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "MainForm.frx":3F835
      Height          =   6015
      Left            =   5040
      TabIndex        =   0
      Top             =   2880
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   10610
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
         DataField       =   "ID"
         Caption         =   "ID"
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
         DataField       =   "Subject Name"
         Caption         =   "Subject Name"
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
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
            Picture         =   "MainForm.frx":3F84A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":40D7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":422AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":43994
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
      Picture         =   "MainForm.frx":4507A
      Top             =   9360
      Width           =   2190
   End
   Begin VB.Image Image12 
      Height          =   1245
      Left            =   0
      Picture         =   "MainForm.frx":4689E
      Top             =   7080
      Width           =   4455
   End
   Begin VB.Image Image11 
      Height          =   735
      Left            =   15960
      Picture         =   "MainForm.frx":49F55
      Top             =   1920
      Width           =   765
   End
   Begin VB.Image Image10 
      Height          =   735
      Left            =   11280
      Picture         =   "MainForm.frx":4A744
      Top             =   1920
      Width           =   4635
   End
   Begin VB.Image Image9 
      Height          =   735
      Left            =   10080
      Picture         =   "MainForm.frx":4AABC
      Top             =   1905
      Width           =   615
   End
   Begin VB.Image Image8 
      Height          =   735
      Left            =   9360
      Picture         =   "MainForm.frx":4B18C
      Top             =   1900
      Width           =   600
   End
   Begin VB.Image Image7 
      Height          =   690
      Left            =   8640
      Picture         =   "MainForm.frx":4B8F8
      Top             =   1920
      Width           =   630
   End
   Begin VB.Image Image6 
      Height          =   1305
      Left            =   0
      Picture         =   "MainForm.frx":4BFAC
      Top             =   6120
      Width           =   4485
   End
   Begin VB.Image Image5 
      Height          =   1050
      Left            =   0
      Picture         =   "MainForm.frx":4F848
      Top             =   5160
      Width           =   4485
   End
   Begin VB.Image Image4 
      Height          =   990
      Left            =   0
      Picture         =   "MainForm.frx":52B02
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
            Picture         =   "MainForm.frx":559EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":5ADF1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":601F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":655F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6A9F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":6FBA1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":74D4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7A3A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":7F9FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":86445
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":8CE8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":93421
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":999B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9BDD5
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9E1F3
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9EE2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":9FA67
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A0661
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A125B
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A1F19
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A2BD7
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MainForm.frx":A3A1D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":A4863
      Top             =   3240
      Width           =   4485
   End
   Begin VB.Image Image2 
      Height          =   1020
      Left            =   0
      Picture         =   "MainForm.frx":A7633
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

Private Function deployTables()
    Adodc1.ConnectionString = db
    'Adodc1.RecordSource = "SELECT id as ID, name as 'Subject Name' FROM rooms"
    Dim rs As New ADODB.Recordset
    Dim conn As New ADODB.Connection
    conn = db
End Function

Private Sub Form_Load()
    msgFadeout
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
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

End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image4.Picture = ImageList1.ListImages(6).Picture
End Sub

Private Sub Image5_Click()
    Label1.Caption = "Sections"
    
    DataGrid4.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid1.Visible = False
    DataGrid5.Visible = False
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image5.Picture = ImageList1.ListImages(8).Picture
End Sub

Private Sub Image6_Click()
    Label1.Caption = "Schedules"
    
    DataGrid5.Visible = True
    
    DataGrid2.Visible = False
    DataGrid3.Visible = False
    DataGrid4.Visible = False
    DataGrid1.Visible = False
    
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
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    buttonsOut
    Set Image11.Picture = ImageList1.ListImages(22).Picture
End Sub

Private Sub Image7_Click()
    goAdd
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
Private Function LoadRoom(args As String)
    Load Room
    Room.Show vbModal
End Function
Private Function LoadUser(args As String)
    Load User
    User.Show vbModal
End Function
Private Function LoadSubject(args As String)
    Load Subject
    Subject.Show vbModal
End Function
Private Function LoadSection(args As String)
    Load Section
    Section.Show vbModal
End Function
Private Function LoadSchedule(args As String)
    Load Schedule
    Schedule.Show vbModal
End Function



