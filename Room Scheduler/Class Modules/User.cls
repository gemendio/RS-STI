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
    Dim rs As New ADODB.Recordset
    Dim field As ADODB.field
    Dim conn As New ADODB.Connection
    
    conn = db
    
    On Error GoTo ErrHandler
        conn.Open
        rs.Open "SELECT * FROM users WHERE id = " & v_id, conn
    
        Id = rs(0).value
        LastName = rs(1).value
        FirstName = rs(2).value
        MiddleName = rs(3).value
        UserType = rs(4).value
    
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
        
        Dim fields(3) As String
        Dim data(3) As Variant

        If Id = 0 Then
            'insert
            fields(0) = "last_name"
            fields(1) = "first_name"
            fields(2) = "middle_name"
            fields(3) = "type"
            
            data(0) = LastName
            data(1) = FirstName
            data(2) = MiddleName
            data(3) = UserType
            
            conn.Execute insertSql("users", data, fields)
        Else
            'update
            Dim where As String
            
            If LastName <> "" Then
                fields(0) = "last_name"
                data(0) = LastName
            End If
            
            If FirstName <> "" Then
                fields(1) = "first_name"
                data(1) = FirstName
            End If
            
            If MiddleName <> "" Then
                fields(2) = "middle_name"
                data(2) = MiddleName
            End If
            
            If UserType <> "" Then
                fields(3) = "type"
                data(3) = UserType
            End If

            where = "WHERE id = " & Id

            conn.Execute updateSql("users", data, fields, where)
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
        
        conn.Execute deleteSql("users", where)
        conn.Close
    Exit Sub
ErrHandler:
    'error message here
End Sub
