Imports System.Data
Imports System.Data.SqlClient

Public Class ConnectionHandler
    Dim conn As SqlConnection
    Dim dataAdapter As SqlDataAdapter
    Dim commandBuilder As SqlCommandBuilder
    Dim data As DataTable

    Public Sub connect()
        If Not conn Is Nothing Then conn.Close()

        Dim hostName As String = ""
        Dim username As String = ""
        Dim password As String = ""
        Dim schemaName As String = "Footprints"
        Dim connStr As String
        connStr = String.Format("server={0};user id={1}; password={2}; database={3}; pooling=false",
                hostName, username, password, schemaName)

        Try
            conn = New SqlConnection(connStr)
            conn.Open()

            'GetDatabases()
            conn.Close()
        Catch ex As SqlException
            MsgBox("Error connecting to the server: " + ex.Message)
        End Try
    End Sub

    Function query(ByVal stringQuery As String) As DataTable
        data = New DataTable
        dataAdapter = New SqlDataAdapter(stringQuery, conn)
        commandBuilder = New SqlCommandBuilder(dataAdapter)

        dataAdapter.Fill(data)

        Return data

    End Function

    Public Sub close()
        Try
            conn.Close()
        Catch ex As SqlException
            MsgBox("Error connecting to the server: " + ex.Message)
        End Try
    End Sub
End Class
