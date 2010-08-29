VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id As Integer
Private m_last_name, m_middle_name, m_first_name As String
Private m_user_type As String

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let Id(value As Integer)
    m_id = value
End Property

Public Property Get LastName() As String
    LastName = m_last_name
End Property

Public Property Let LastName(value As String)
    m_last_name = value
End Property

Public Property Get FirstName() As String
    FirstName = m_first_name
End Property

Public Property Let FirstName(value As String)
    m_first_name = value
End Property

Public Property Get MiddleName() As String
    MiddleName = m_middle_name
End Property

Public Property Let MiddleName(value As String)
    m_middle_name = value
End Property

Public Property Get UserType() As String
    UserType = m_user_type
End Property

Public Property Let UserType(value As String)
    m_user_type = value
End Property

Public Sub Load(v_id As Integer)
    Dim row As New ADODB.Recordset
    Dim conn As New ADODB.Connection

    conn = db

    On Error GoTo ErrHandler
        conn.Open

        row.Open "SELECT * FROM users WHERE id = " & v_id, conn
                
        Id = row(0).value
        LastName = row(1).value
        FirstName = row(2).value
        MiddleName = row(3).value
        UserType = row(4).value
        
        row.Close
        conn.Close
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Upsert()
    On Error GoTo ErrHandler
        Dim fields(3) As String
        Dim data(3) As Variant
        
        fields(0) = "last_name"
        fields(1) = "first_name"
        fields(2) = "middle_name"
        fields(3) = "type"
            
        data(0) = LastName
        data(1) = FirstName
        data(2) = MiddleName
        data(3) = UserType
        
        Call UpsertTable("users", data, fields, Id)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Delete()
    On Error GoTo ErrHandler
        Dim where As String
        
        where = "WHERE id = " & Id
        
        Call DeleteFromTable("users", where)
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
                & "last_name, " _
                & "first_name, " _
                & "middle_name, " _
                & "type as user_type " _
            & "FROM users", conn
        
        Set GetAll = rs
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Function



