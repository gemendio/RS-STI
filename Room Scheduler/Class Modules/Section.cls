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

Private m_id As Integer
Private m_name As String

Public Property Get Id() As Integer
    Id = m_id
End Property

Public Property Let Id(value As Integer)
    m_id = value
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
        data(0) = Name

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
    
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        
        rs.Open "SELECT id as section_id, name as section_name FROM sections", conn
        
        Set GetAll = rs
    Exit Function
ErrHandler:
    MsgBox Err.Description, vbExclamation
End Function



