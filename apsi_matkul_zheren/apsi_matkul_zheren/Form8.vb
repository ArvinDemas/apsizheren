Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form8
    Dim CMD As OleDbCommand
    Dim DR As OleDbDataReader
    Dim tt As New ToolTip()

    Sub ClearInput()
        txtkoderc.Clear()
        txtnoid.Text = form1.LoginNoID
        txtkodep.Clear()
        txtnamap.Clear()
        RichTextBox1.Clear()
        txtsearch.Clear()
    End Sub
    Function CekNama(nama As String) As Boolean
        Call Koneksi()
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pesanan WHERE UCASE(Kode_Pesanan) = UCASE(@kode1)", Conn)
        CMD.Parameters.AddWithValue("@kode1", nama)
        Dim jumlah As Integer = Convert.ToInt32(CMD.ExecuteScalar())
        Return (jumlah > 0)
    End Function
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Request_Customer", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
    End Sub
    Private Sub btninput_Click(sender As Object, e As EventArgs) Handles btninput.Click
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        If txtkoderc.Text = "" Or txtnoid.Text = "" Or txtkodep.Text = "" Or txtnamap.Text = "" Or RichTextBox1.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Request_Customer WHERE Kode_Request_Customer=@kode", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkoderc.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())
        If count > 0 Then
            MsgBox("Kode request sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If
        If Not CekNama(txtkodep.Text.Trim()) Then
            MsgBox("Nama kode pesanan tidak sesuai dengan database!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        CMD = New OleDbCommand("INSERT INTO Data_Request_Customer (Kode_Request_Customer, No_ID, Kode_Pesanan, Nama_Produk, Request_Desain) " & "VALUES (@kode, @id, @kode1, @nama, @request)", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkoderc.Text)
        CMD.Parameters.AddWithValue("@id", txtnoid.Text)
        CMD.Parameters.AddWithValue("@kode1", txtkodep.Text)
        CMD.Parameters.AddWithValue("@nama", txtnamap.Text)
        CMD.Parameters.AddWithValue("@request", RichTextBox1.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        Call ClearInput()
        Conn.Close()
    End Sub


    Private Sub Form8_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        txtnoid.Text = form1.LoginNoID
        tt.ToolTipTitle = "Pengingat"
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.IsBalloon = True
        tt.SetToolTip(txtkoderc, "Isi kode bahan baku dengan lengkap.")
        Tampildatabase()
    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Request_Customer").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As New OleDbCommand("DELETE FROM Data_Request_Customer WHERE Kode_Request_Customer=@kode", Conn)
                CMD.Parameters.AddWithValue("@kode", kode)
                CMD.ExecuteNonQuery()
                Conn.Close()
                Tampildatabase()
                MsgBox("Data berhasil dihapus dari database!", MsgBoxStyle.Information)
            End If
        Else
            MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub btndeleteall_Click(sender As Object, e As EventArgs) Handles btndeleteall.Click
        Call ClearInput()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            txtkoderc.Text = row.Cells("Kode_Request_Customer").Value.ToString()
            txtnoid.Text = row.Cells("No_ID").Value.ToString()
            txtkodep.Text = row.Cells("Kode_Pesanan").Value.ToString()
            txtnamap.Text = row.Cells("Nama_Produk").Value.ToString()
            RichTextBox1.Text = row.Cells("Request_Desain").Value.ToString()
        End If
    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            row.Cells("Kode_Request_Customer").Value = txtkoderc.Text
            row.Cells("No_ID").Value = txtnoid.Text
            row.Cells("Kode_Pesanan").Value = txtkodep.Text
            row.Cells("Nama_Produk").Value = txtnamap.Text
            row.Cells("Request_Desain").Value = RichTextBox1.Text
            MsgBox("Data berhasil diedit !", MsgBoxStyle.Information)
            ClearInput()
        Else
            MsgBox("Pilih data yang ingin diedit dari tabel!", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan kode pesanan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Request_Customer WHERE Kode_Request_Customer LIKE '%" & txtsearch.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView1.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub

    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If txtsearch.Text.Trim() = "" Then
            Call Koneksi()
            Tampildatabase()
            Conn.Close()
        End If
    End Sub
    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        ClearInput()
    End Sub
    Private Sub btnback_Click(sender As Object, e As EventArgs) Handles btnback.Click
        Form5.Show()
        Me.Hide()
        ClearInput()
    End Sub

    Private Sub txtkodep_TextChanged(sender As Object, e As EventArgs) Handles txtkodep.TextChanged

    End Sub
End Class