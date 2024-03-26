Imports System.Data.SqlClient

Module db_functions
    Private Function CONNECTION_STRING() As String
        If Debugger.IsAttached Then
            Return String.Concat("Server=BOSSLAW-LT\SQLEXPRESS;Database=", My.Settings.DatabaseName, ";User Id=sa;Password = nica@5685647")
        Else
            Return String.Concat("Server=", My.Settings.ServerAddress, ";Database=", My.Settings.DatabaseName, ";User Id=sa;Password = Datanet1")
        End If
    End Function
    Public Function DB_EXECUTE_INSERT(ByVal TableName As String, ByVal InsertField As String, ByVal InsertValue As String) As Integer
        Dim sm_QUERY As String = String.Concat("INSERT INTO ", TableName, " (", InsertField, ") VALUES (", InsertValue, ")")
        Dim sm_SQL_CONNECTION As New SqlConnection(CONNECTION_STRING)
        Dim sm_SQL_COMMAND As New SqlCommand(sm_QUERY)
        Dim sm_RETVAL As Integer = 0
        Try
            sm_SQL_COMMAND.Connection = sm_SQL_CONNECTION
            If sm_SQL_CONNECTION.State = ConnectionState.Open Then sm_SQL_CONNECTION.Close()
            sm_SQL_CONNECTION.Open()
            sm_SQL_COMMAND.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            sm_SQL_COMMAND.Dispose()
            sm_SQL_CONNECTION.Close() : sm_SQL_CONNECTION.Dispose()
            sm_SQL_COMMAND = Nothing : sm_SQL_CONNECTION = Nothing
        End Try
        Return sm_RETVAL
    End Function

    Public Function DB_EXECUTE_UPDATE(ByVal TableName As String, ByVal UpdateField As String, ByVal FilterField As String) As Integer
        Dim sm_QUERY As String = String.Concat("UPDATE ", TableName, " SET ", UpdateField, " WHERE ", FilterField)
        Dim sm_SQL_CONNECTION As New SqlConnection(CONNECTION_STRING)
        Dim sm_SQL_COMMAND As New SqlCommand(sm_QUERY)
        Dim sm_RETVAL As Integer = 0
        Try
            sm_SQL_COMMAND.Connection = sm_SQL_CONNECTION
            If sm_SQL_CONNECTION.State = ConnectionState.Open Then sm_SQL_CONNECTION.Close()
            sm_SQL_CONNECTION.Open()
            sm_SQL_COMMAND.ExecuteNonQuery()
        Catch ex As Exception
        Finally
            sm_SQL_COMMAND.Dispose()
            sm_SQL_CONNECTION.Close() : sm_SQL_CONNECTION.Dispose()
            sm_SQL_COMMAND = Nothing : sm_SQL_CONNECTION = Nothing
        End Try
        Return sm_RETVAL
    End Function

    Public Function DB_EXECUTE_SELECT(ByVal SelectQuery As String) As DataTable
        Dim sm_SQL_CONNECTION As New SqlConnection(CONNECTION_STRING)
        Dim sm_SQL_COMMAND As New SqlCommand(SelectQuery)
        Dim sm_SQL_ADAPTER As New SqlDataAdapter
        Dim sm_SQL_DATASET As New DataSet
        Try
            If sm_SQL_CONNECTION.State = ConnectionState.Open Then sm_SQL_CONNECTION.Close()
            sm_SQL_CONNECTION.Open()
            sm_SQL_ADAPTER.SelectCommand = New SqlCommand(SelectQuery, sm_SQL_CONNECTION)
            sm_SQL_ADAPTER.Fill(sm_SQL_DATASET)
            Return sm_SQL_DATASET.Tables(0)
        Catch ex As Exception
            Return Nothing
        Finally
            sm_SQL_COMMAND.Dispose()
            sm_SQL_CONNECTION.Close() : sm_SQL_CONNECTION.Dispose()
            sm_SQL_COMMAND = Nothing : sm_SQL_CONNECTION = Nothing
        End Try
    End Function

    Public Function DB_EXECUTE_DELETE(ByVal TableName As String, ByVal WhereQuery As String) As Boolean
        Dim sm_QUERY As String = Nothing
        If Trim(WhereQuery).Length = 0 Then
            sm_QUERY = String.Concat("DELETE FROM ", TableName)
        Else
            sm_QUERY = String.Concat("DELETE FROM ", TableName, " WHERE ", WhereQuery)
        End If
        Dim sm_SQL_CONNECTION As New SqlConnection(CONNECTION_STRING)
        Dim sm_SQL_COMMAND As New SqlCommand(sm_QUERY)
        Try
            sm_SQL_COMMAND.Connection = sm_SQL_CONNECTION
            If sm_SQL_CONNECTION.State = ConnectionState.Open Then sm_SQL_CONNECTION.Close()
            sm_SQL_CONNECTION.Open()
            sm_SQL_COMMAND.ExecuteNonQuery()
            Return True
        Catch ex As Exception
        Finally
            sm_SQL_COMMAND.Dispose()
            sm_SQL_CONNECTION.Close() : sm_SQL_CONNECTION.Dispose()
            sm_SQL_COMMAND = Nothing : sm_SQL_CONNECTION = Nothing
        End Try
    End Function

End Module
