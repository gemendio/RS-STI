VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id, m_user_id As Integer
Private m_name As String

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let Id(value As Integer)
    m_id = value
End Property
Public Property Get UserId() As Integer
    UserId = m_user_id
End Property

Public Property Let UserId(value As Integer)
    m_userId = value
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(value As String)
    m_name = value
End Property

Public Sub Load(v_id As Integer)
    Dim row As New ADODB.Recordset
    Dim conn As New ADODB.Connection

    conn = db

    On Error GoTo ErrHandler
        conn.Open

        row.Open "SELECT * FROM sections WHERE id = " & v_id, conn
                
        Id = row(0).value
        Name = row(1).value
        UserId = row(3).value
        
        row.Close
        conn.Close
    Exit Sub
    
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Upsert()
    On Error GoTo ErrHandler
        Dim fields(0) As String
        Dim data(0) As Variant

        fields(0) = "name"
        fields(1) = "user_id"
        data(0) = Name
        data(1) = UserId
        Call UpsertTable("sections", data, fields, Id)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Sub Delete()
    On Error GoTo ErrHandler
        Dim where As String
        
        where = "WHERE id = " & Id
        
        Call DeleteFromTable("sections", where)
    Exit Sub
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Sub

Public Function GetAll() As ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim conn As New ADODB.Connection
    Dim sqlScript As String
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
            
        rs.CursorLocation = adUseClient
        
        sqlScript = "SELECT sec.id as ID, sec.name as 'Section Name', CONCAT(usr. last_name , ', ' , usr.first_name) as User "
        sqlScript = sqlScript + "FROM sections sec JOIN users usr ON sec.user_id=usr.id"
        
        rs.Open sqlScript, conn, adOpenStatic, adLockOptimistic
        
        Set GetAll = rs
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Function

Public Function search(strSeek As String)
    Dim rs As New ADODB.Recordset
    Dim conn As New ADODB.Connection
    Dim sqlScript As String
        
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        
        rs.CursorLocation = adUseClient
        
        sqlScript = "SELECT sec.id as ID, sec.name as 'Section Name', CONCAT(usr.last_name , ', ' , usr.first_name) "
        sqlScript = sqlScript + " as User FROM sections sec JOIN users usr ON sec.user_id=usr.id WHERE CONCAT(sec.name, usr.last_name, usr.first_name) LIKE '%" + strSeek + "%'"
        
        rs.Open sqlScript, conn
        
        Set search = rs
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation
    
End Function
