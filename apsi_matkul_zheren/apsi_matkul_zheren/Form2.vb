Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form2
    Dim CMD As OleDbCommand
    Dim DR As OleDbDataReader
    Dim tt As New ToolTip()

    Sub ClearInput()
        txtnp.Clear()
        txttop.Clear()
        txtnoid.Text = form1.LoginNoID
        txtkodep.Clear()
        RichTextBox1.Clear()
        RichTextBox2.Clear()
        txtstatus.Clear()
        txtsearch.Clear()

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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pesanan", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
    End Sub
    Private Sub btninput_Click(sender As Object, e As EventArgs) Handles btninput.Click
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        Call Koneksi()
        If txtkodep.Text = "" Or txtnoid.Text = "" Or txtnp.Text = "" Or txttop.Text = "" Or RichTextBox1.Text = "" Or RichTextBox2.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pesanan WHERE Kode_Pesanan=@kode", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodep.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())

        If count > 0 Then
            MsgBox("Kode Pesanan sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If

        CMD = New OleDbCommand("INSERT INTO Data_Pesanan (Kode_Pesanan, No_ID, Nama_Produk, Request_Desain, Jumlah_Produk, Deadline_Pesanan, Lokasi_Pengiriman, Status_Pesanan) " & "VALUES (@kode, @id, @nama, @request, @total, @deadline, @lokasi, @status)", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodep.Text)
        CMD.Parameters.AddWithValue("@id", txtnoid.Text)
        CMD.Parameters.AddWithValue("@nama", txtnp.Text)
        CMD.Parameters.AddWithValue("@request", RichTextBox1.Text)
        CMD.Parameters.AddWithValue("@total", txttop.Text)
        CMD.Parameters.AddWithValue("@deadline", DateTimePicker1.Value)
        CMD.Parameters.AddWithValue("@lokasi", RichTextBox2.Text)
        CMD.Parameters.AddWithValue("@status", txtstatus.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        ClearInput()
    End Sub

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        txtnoid.Text = form1.LoginNoID
        tt.ToolTipTitle = "Pengingat"
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.IsBalloon = True
        tt.SetToolTip(txtkodep, "Isi kode bahan baku dengan lengkap.")
        Tampildatabase()
    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Pesanan").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As New OleDbCommand("DELETE FROM Data_Pesanan WHERE Kode_Pesanan=@kode", Conn)
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
            txtkodep.Text = row.Cells("Kode_Pesanan").Value.ToString()
            txtnoid.Text = row.Cells("No_ID").Value.ToString()
            txtnp.Text = row.Cells("Nama_Produk").Value.ToString()
            RichTextBox1.Text = row.Cells("Request_Desain").Value.ToString()
            txttop.Text = row.Cells("Jumlah_Produk").Value.ToString()
            If Not IsDBNull(row.Cells("Deadline_Pesanan").Value) Then
                DateTimePicker1.Value = CDate(row.Cells("Deadline_Pesanan").Value)
            End If
            RichTextBox2.Text = row.Cells("Lokasi_Pengiriman").Value.ToString()
            txtstatus.Text = row.Cells("Status_Pesanan").Value.ToString()
        End If
    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            row.Cells("Kode_Pesanan").Value = txtkodep.Text
            row.Cells("No_ID").Value = txtnoid.Text
            row.Cells("Nama_Produk").Value = txtnp.Text
            row.Cells("Request_Desain").Value = RichTextBox1.Text
            row.Cells("Jumlah_Produk").Value = txttop.Text
            row.Cells("Deadline_Pesanan").Value = DateTimePicker1.Value
            row.Cells("Lokasi_Pengiriman").Value = RichTextBox2.Text
            row.Cells("Status_Pesanan").Value = txtstatus.Text
            MsgBox(" Data berhasil diedit !", MsgBoxStyle.Information)
            ClearInput()
        Else
            MsgBox("Pilih data yang ingin diedit dari tabel!", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan nama pesanan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pesanan WHERE Nama_Produk LIKE '%" & txtsearch.Text & "%'", Conn)
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
            Tampildatabase()
            Conn.Close()
        End If
    End Sub
    Private Sub txttop_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txttop.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        ClearInput()
        form1.txtnoid.Clear()
    End Sub
    Private Sub btncek_Click(sender As Object, e As EventArgs) Handles btncek.Click
        Form4.Show()
        Me.Hide()
        ClearInput()
    End Sub

    Private Sub btnnext_Click(sender As Object, e As EventArgs) Handles btnnext.Click
        Form3.Show()
        Me.Hide()
        ClearInput()
    End Sub

    Private Sub txtnp_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtnp.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> " "c Then
            e.Handled = True
        End If
    End Sub

End Class