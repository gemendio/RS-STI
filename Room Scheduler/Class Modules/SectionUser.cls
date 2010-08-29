VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelSectionUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id, m_section_id, m_user_id As Integer

Public Property Let Id(value As Integer)
    m_id = value
End Property

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let SectionId(value As Integer)
    m_section_id = value
End Property

Public Property Get SectionId() As Integer
    SectionId = m_section_id
End Property

Public Property Let UserId(value As Integer)
    m_user_id = value
End Property

Public Property Get UserId() As Integer
    UserId = m_user_id
End Property

Public Sub Delete()
    On Error GoTo ErrHandler
        Dim where As String
        
        where = "WHERE id = " & Id
        
        Call DeleteFromTable("section_users", where)
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

        row.Open "SELECT * FROM section_users WHERE id = " & v_id, conn
                
        Id = row(0).value
        SectionId = row(1).value
        UserId = row(2).value
        
        row.Close
        conn.Close
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Upsert()
    On Error GoTo ErrHandler
        Dim fields(1) As String
        Dim data(1) As Variant
        
        fields(0) = "section_id"
        fields(1) = "user_id"
        
        data(0) = SectionId
        data(1) = UserId
        
        Call UpsertTable("section_users", data, fields, Id)
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
                & "id as section_user_id, " _
                & "section_id, " _
                & "user_id " _
            & "FROM section_users", conn
        
        Set GetAll = rs
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Function


