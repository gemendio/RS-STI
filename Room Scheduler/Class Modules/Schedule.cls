VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelSchedule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id, m_room_id, m_subject_id, m_section_id As Integer
Private m_day, m_start_time, m_end_time As String

Public Property Let Id(value As Integer)
    m_id = value
End Property

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let RoomId(value As Integer)
    m_room_id = value
End Property

Public Property Get RoomId() As Integer
    RoomId = m_room_id
End Property

Public Property Let SubjectId(value As Integer)
    m_subject_id = value
End Property

Public Property Get SubjectId() As Integer
    SubjectId = m_subject_id
End Property

Public Property Let SectionId(value As Integer)
    m_section_id = value
End Property

Public Property Get SectionId() As Integer
    SectionId = m_section_id
End Property

Public Property Let Day(value As String)
    m_day = value
End Property

Public Property Get Day() As String
    Day = m_day
End Property

Public Property Let StartTime(value As String)
    m_start_time = value
End Property

Public Property Get StartTime() As String
    StartTime = m_start_time
End Property

Public Property Let EndTime(value As String)
    m_end_time = value
End Property

Public Property Get EndTime() As String
    EndTime = m_end_time
End Property

Public Sub Delete()
    On Error GoTo ErrHandler
        Dim where As String
        
        where = "WHERE id = " & Id
        
        Call DeleteFromTable("schedules", where)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Load(v_id As Integer)
    Dim row As New ADODB.Recordset
    Dim conn As New ADODB.Connection

    conn = db

    On Error GoTo ErrHandler
        conn.Open

        row.Open "SELECT * FROM schedules WHERE id = " & v_id, conn
                
        Id = row(0).value
        Day = row(1).value
        StartTime = row(2).value
        EndTime = row(3).value
        RoomId = row(4).value
        SubjectId = row(5).value
        SectionId = row(6).value
        
        row.Close
        conn.Close
    Exit Sub
    
ErrHandler:
    If Err.Number = 94 Then
        Resume Next
    Else
        MsgBox Err.Description, vbExclamation
    End If
End Sub

Public Sub Upsert()
    On Error GoTo ErrHandler
        Dim fields(5) As String
        Dim data(5) As Variant
        
        fields(0) = "day"
        fields(1) = "start_time"
        fields(2) = "end_time"
        fields(3) = "room_id"
        fields(4) = "subject_id"
        fields(5) = "section_id"
            
        data(0) = Day
        data(1) = StartTime
        data(2) = EndTime
        data(3) = RoomId
        data(4) = SubjectId
        data(5) = SectionId
        
        Call UpsertTable("schedules", data, fields, Id)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Function GetAll() As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim conn As New ADODB.Connection
    
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        
        rs.Open "SELECT " _
                & "id as user_id, " _
                & "day, " _
                & "start_time, " _
                & "end_time, " _
                & "room_id, " _
                & "subject_id, " _
                & "section_id " _
            & "FROM schedules", conn
        
        Set GetAll = rs
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Function


