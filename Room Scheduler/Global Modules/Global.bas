Attribute VB_Name = "Global"
Public Function db() As ADODB.Connection
    Dim c As New ADODB.Connection
    c.ConnectionString = "DRIVER={MySQL ODBC 5.1 Driver};" _
                         & "SERVER=localhost;" _
                         & "DATABASE=room_scheduler;" _
                         & "UID=root;" _
                         & "PWD=amirah@1;" _
                         & "OPTION=3"
    Set db = c
End Function

Public Function insertSql(table_name As String, data() As Variant, fields() As String)
    If (UBound(data) = UBound(fields)) And table_name <> "" Then
        insertSql = "INSERT INTO " & table_name _
            & " (" & Join(fields, ", ") & ") " _
            & "VALUES('" & Join(data, "', '") & "')"
    End If
End Function

Public Function deleteSql(table_name As String, where As String)
    If table_name <> "" And where <> "" Then
        deleteSql = "DELETE FROM " & table_name & " " & where
    End If
End Function

Public Function updateSql(table_name As String, data() As Variant, fields() As String, where As String)
    If (UBound(data) = UBound(fields)) And table_name <> "" Then
        Dim i As Integer
        
        Dim set_sql As String

        For i = 0 To UBound(fields)
            If i <> UBound(fields) Then
                set_sql = set_sql & fields(i) & " = '" & data(i) & "', "
            Else
                set_sql = set_sql & fields(i) & " = '" & data(i) & "' "
            End If
        Next i

        updateSql = "UPDATE " & table_name _
            & " SET " & set_sql & where
    End If
End Function

