Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form4
    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Tampildatabase()
    End Sub
    Sub Tampildatabase()
        Try
            Call Koneksi()
            If Conn Is Nothing Then
                MsgBox("Koneksi masih kosong!")
                Exit Sub
            End If
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Tujuan_Pengiriman", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
    End Sub
    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan kode tujuan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Tujuan_Pengiriman WHERE Kode_Tujuan_Pengiriman LIKE '%" & txtsearch.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView1.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnback_Click(sender As Object, e As EventArgs) Handles btnback.Click
        Form2.Show()
        Me.Hide()
        txtsearch.Clear()
    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        txtsearch.Clear()
        form1.txtnoid.Clear()
    End Sub
    Private Sub btnshowdata_Click(sender As Object, e As EventArgs) Handles btnshowdata.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Tujuan_Pengiriman, No_ID, Tempat_Tujuan_Distribusi, Kontak_Penerima FROM Data_Tujuan_Pengiriman", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Pengiriman")
        DataGridView1.DataSource = Ds.Tables("Pengiriman")
        DataGridView1.ReadOnly = True
        DataGridView1.AllowUserToAddRows = False
        Conn.Close()
    End Sub
End Class