Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form6
    Dim CMD As OleDbCommand
    Dim DR As OleDbDataReader
    Dim tt As New ToolTip()

    Sub ClearInput()
        txtkodebb.Clear()
        txtnoid.Text = form1.LoginNoID
        cmbjenisbb.SelectedIndex = -1
        txtnamabb.Clear()
        txttotbb.Clear()
        txtsearch.Clear()
    End Sub
    Function CekNama(nama As String) As Boolean
        Call Koneksi()
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Bahan_Baku WHERE UCASE(Nama_Bahan_Baku) = UCASE(@nama)", Conn)
        CMD.Parameters.AddWithValue("@nama", nama)
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Bahan_Baku", Conn)
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
        If txtkodebb.Text = "" Or txtnoid.Text = "" Or cmbjenisbb.Text = "" Or txttotbb.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Bahan_Baku WHERE Kode_Bahan_Baku=@kode", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodebb.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())

        If count > 0 Then
            MsgBox("Kode bahan baku sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If
        If Not CekNama(txtnamabb.Text.Trim()) Then
            MsgBox("Nama bahan baku tidak sesuai dengan database!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        CMD = New OleDbCommand("INSERT INTO Data_Bahan_Baku (Kode_Bahan_Baku, No_ID, Jenis_Bahan_Baku, Nama_Bahan_Baku, Total_Bahan_Baku) " & "VALUES (@kode, @id, @jenis, @nama, @total)", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodebb.Text)
        CMD.Parameters.AddWithValue("@id", txtnoid.Text)
        CMD.Parameters.AddWithValue("@jenis", cmbjenisbb.Text)
        CMD.Parameters.AddWithValue("@nama", txtnamabb.Text)
        CMD.Parameters.AddWithValue("@total", txttotbb.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        Call ClearInput()
        Conn.Close()
    End Sub

    Private Sub IsiComboJenis()
        Dim Conn As New OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Application.StartupPath & "\M4_TUBES.mdb;")
        Conn.Open()
        Dim CMD As New OleDbCommand("SELECT DISTINCT Jenis_Bahan_Baku FROM Data_Bahan_Baku", Conn)
        Dim DR As OleDbDataReader = CMD.ExecuteReader()
        cmbjenisbb.Items.Clear()
        While DR.Read()
            If Not IsDBNull(DR("Jenis_Bahan_Baku")) Then
                cmbjenisbb.Items.Add(DR("Jenis_Bahan_Baku").ToString)
            End If
        End While
        DR.Close()
        Conn.Close()
    End Sub

    Private Sub Form6_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        txtnoid.Text = form1.LoginNoID
        tt.ToolTipTitle = "Pengingat"
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.IsBalloon = True
        tt.SetToolTip(txtkodebb, "Isi kode bahan baku dengan lengkap.")
        IsiComboJenis()
        Tampildatabase()
    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Bahan_Baku").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As New OleDbCommand("DELETE FROM Data_Bahan_Baku WHERE Kode_Bahan_Baku=@kode", Conn)
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
            txtkodebb.Text = row.Cells("Kode_Bahan_Baku").Value.ToString()
            txtnoid.Text = row.Cells("No_ID").Value.ToString()
            cmbjenisbb.Text = row.Cells("Jenis_Bahan_Baku").Value.ToString()
            txtnamabb.Text = row.Cells("Nama_Bahan_Baku").Value.ToString()
            txttotbb.Text = row.Cells("Total_Bahan_Baku").Value.ToString()
        End If
    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            row.Cells("Kode_Bahan_Baku").Value = txtkodebb.Text
            row.Cells("No_ID").Value = txtnoid.Text
            row.Cells("Jenis_Bahan_Baku").Value = cmbjenisbb.Text
            row.Cells("Nama_Bahan_Baku").Value = txtnamabb.Text
            row.Cells("Total_Bahan_Baku").Value = txttotbb.Text
            MsgBox(" Data berhasil diedit !", MsgBoxStyle.Information)
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
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Bahan_Baku WHERE Kode_Bahan_Baku LIKE '%" & txtsearch.Text & "%'", Conn)
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
    Private Sub txttotbb_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txttotbb.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
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

    Private Sub txtnamabb_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtnamabb.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> " "c Then
            e.Handled = True
        End If
    End Sub


End Class