Imports System.Data.OleDb

Module Module1
    Public Conn As OleDbConnection
    Public Da As OleDbDataAdapter
    Public Ds As DataSet
    Public Cmd As OleDbCommand
    Public Dr As OleDbDataReader
    Public Record As New BindingSource
    Public Dt As DataTable

    Public Connect As String =
        "Provider=Microsoft.JET.OLEDB.4.0;Data Source=D:\Document\apsi pak muhsin part 2\apsi_matkul_zheren\apsi_matkul_zheren\bin\Debug\M4_TUBES.mdb;"

    Public Sub Koneksi()
        If Conn Is Nothing Then
            Conn = New OleDbConnection(Connect)
        End If

        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
    End Sub

End Module