Attribute VB_Name = "Global"
Public Function db() As ADODB.Connection
    Dim c As New ADODB.Connection
    c.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
                         & "SERVER=localhost;" _
                         & "DATABASE=room_scheduler;" _
                         & "UID=root;" _
                         & "PWD=p@ssword;" _
                         & "OPTION=3"
    Set db = c
End Function

Public Function insertSql(table_name As String, data() As Variant, fields() As String)
    Dim i As Integer
    Dim values_sql As String
    
    For i = 0 To UBound(data)
        If i <> UBound(data) Then
            If data(i) = 0 And i <> 0 Then
                values_sql = values_sql & "NULL, "
            Else
                values_sql = values_sql & "'" & data(i) & "', "
            End If
        Else
            If data(i) = 0 And i <> 0 Then
                values_sql = values_sql & "NULL"
            Else
                values_sql = values_sql & "'" & data(i) & "' "
            End If
        End If
    Next i

    insertSql = "INSERT INTO " & table_name _
        & " (" & Join(fields, ", ") & ") " _
        & "VALUES(" & values_sql & ")"
End Function

Public Function deleteSql(table_name As String, where As String)
    deleteSql = "DELETE FROM " & table_name & " " & where
End Function

Public Function updateSql(table_name As String, data() As Variant, fields() As String, where As String)
    Dim i As Integer
    Dim set_sql As String

    For i = 0 To UBound(fields)
        If i <> UBound(fields) Then
            If data(i) = 0 And i <> 0 Then
                set_sql = set_sql & fields(i) & " = NULL, "
            Else
                set_sql = set_sql & fields(i) & " = '" & data(i) & "', "
            End If
        Else
            If data(i) = 0 And i <> 0 Then
                set_sql = set_sql & fields(i) & " = NULL "
            Else
                set_sql = set_sql & fields(i) & " = '" & data(i) & "' "
            End If
        End If
    Next i

MsgBox "UPDATE " & table_name _
        & " SET " & set_sql & where
    updateSql = "UPDATE " & table_name _
        & " SET " & set_sql & where
End Function

Public Sub UpsertTable(table_name As String, data() As Variant, fields() As String, v_id As Integer)
    Dim conn As New ADODB.Connection
    conn = db

    conn.Open
        
    If v_id = 0 Then
        conn.Execute insertSql(table_name, data, fields)
    Else
        Dim where As String
            
        where = "WHERE id = " & v_id

        conn.Execute updateSql(table_name, data, fields, where)
    End If
        
    conn.Close
End Sub

Public Sub DeleteFromTable(table_name As String, where As String)
    Dim conn As New ADODB.Connection
    conn = db
    
    conn.Open
    conn.Execute deleteSql(table_name, where)
    conn.Close
End Sub

Public Function RemoveArrayElement(v_array As Variant, element As Integer) As Variant()
    Dim i, n As Integer
    Dim new_array() As Variant
    ReDim new_array(UBound(v_array) - 1)
    
    n = 0
    
    For i = 0 To UBound(v_array)
        If i = element Then
            i = i + 1
        End If
        
        If i > UBound(v_array) Then
            Exit For
        Else
            new_array(n) = v_array(i)
        End If
        
        n = n + 1
    Next
    
    RemoveArrayElement = new_array
End Function


