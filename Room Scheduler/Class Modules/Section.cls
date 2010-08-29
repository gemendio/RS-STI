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
    m_user_id = value
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(value As String)
    m_name = value
End Property

Public Sub Load(v_id As Integer)
    Dim rs As New ADODB.Recordset
    Dim field As ADODB.field
    Dim conn As New ADODB.Connection
    
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        rs.Open "SELECT * FROM sections WHERE id = " & v_id, conn
    
        Id = rs(0).value
        Name = rs(1).value
        UserId = rs(2).value
    
        rs.Close
        conn.Close
    Exit Sub
    
ErrHandler:
    'error message here!
End Sub

Public Sub Upsert()
    Dim conn As New ADODB.Connection
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        
        Dim fields(1) As String
        Dim data(1) As Variant

        If Id = 0 Then
            'insert
            fields(0) = "name"
            fields(1) = "user_id"
            
            data(0) = Name
            data(1) = UserId
            
            conn.Execute insertSql("sections", data, fields)
        Else
            'update
            Dim where As String
            
            If Name <> "" Then
                fields(0) = "name"
                data(0) = Name
            End If
            
            If UserId <> 0 Then
                fields(1) = "user_id"
                data(1) = UserId
            End If
        
            where = "WHERE id = " & Id

            conn.Execute updateSql("sections", data, fields, where)
        End If
        
        conn.Close
    Exit Sub
ErrHandler:
    'error message here!
End Sub

Public Sub Delete()
    Dim conn As New ADODB.Connection
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
       
        Dim where As String
        
        where = "WHERE id = " & Id
        
        conn.Execute deleteSql("sections", where)
        conn.Close
    Exit Sub
ErrHandler:
    'error message here
End Sub
