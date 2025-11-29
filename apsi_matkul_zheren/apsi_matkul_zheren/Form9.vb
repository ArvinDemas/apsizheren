Imports System.Data.OleDb
Imports System.Diagnostics

Public Class Form9
    Private Sub btnshowdata_Click(sender As Object, e As EventArgs) Handles btnshowdata.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Pesanan, No_ID, Nama_Produk, Request_Desain, Jumlah_Produk, Deadline_Pesanan, Lokasi_Pengiriman, Status_Pesanan FROM Data_Pesanan", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Pesanan")
        DataGridView1.DataSource = Ds.Tables("Pesanan")
        DataGridView1.ReadOnly = True
        DataGridView1.AllowUserToAddRows = False
        Conn.Close()
    End Sub

    Private Sub btnback_Click(sender As Object, e As EventArgs) Handles btnback.Click
        DataGridView1.DataSource = Nothing
        DataGridView1.Rows.Clear()
        Form5.Show()
        Me.Hide()
    End Sub

    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pesanan WHERE Kode_Pesanan LIKE '%" & txtsearch.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView1.DataSource = Dt
        Else
            MsgBox("Data dengan nama tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If txtsearch.Text.Trim() = "" Then
            Call Koneksi()
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Pesanan, No_ID, Nama_Produk, Request_Desain, Jumlah_Produk, Deadline_Pesanan, Lokasi_Pengiriman, Status_Pesanan FROM Data_Pesanan", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Pesanan")
            DataGridView1.DataSource = Ds.Tables("Pesanan")
            DataGridView1.ReadOnly = True
            DataGridView1.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub
    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
    End Sub


End Class