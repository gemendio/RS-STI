VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ModelRoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_id As Integer
Private m_name As String

Public Property Let Id(value As Integer)
    m_id = value
End Property

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let Name(value As String)
    m_name = value
End Property

Public Property Get Name() As String
    Name = m_name
End Property

Public Sub Delete()
    Dim conn As New ADODB.Connection
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
       
        Dim where As String
        
        where = "WHERE id = " & Id
        
        conn.Execute deleteSql("rooms", where)
        conn.Close
    Exit Sub
ErrHandler:
    'error message here
End Sub

Public Sub Load(v_id As Integer)
    Dim rs As New ADODB.Recordset
    Dim field As ADODB.field
    Dim conn As New ADODB.Connection
    
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        rs.Open "SELECT * FROM rooms WHERE id = " & v_id, conn
    
        Id = rs(0).value
        Name = rs(1).value
    
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
        
        Dim fields(0) As String
        Dim data(0) As Variant

        If Id = 0 Then
            'insert
            fields(0) = "name"
            data(0) = Name
            
            conn.Execute insertSql("rooms", data, fields)
        Else
            'update
            Dim where As String
            
            fields(0) = "name"
            data(0) = Name

            where = "WHERE id = " & Id

            conn.Execute updateSql("rooms", data, fields, where)
        End If
        
        conn.Close
    Exit Sub
ErrHandler:
    'error message here!
End Sub

