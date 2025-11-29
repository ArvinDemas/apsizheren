Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form13
    Dim CMD As OleDbCommand
    Dim DR As OleDbDataReader
    Dim tt As New ToolTip()

    Sub ClearInput()
        txtkodegaji.Clear()
        txtnoid.Text = form1.LoginNoID
        txtnama.Clear()
        txtjam.Clear()
        txtpokok.Clear()
        txttot.Clear()
        txtsearch.Clear()
    End Sub
    Function CekNama(nama As String) As Boolean
        Call Koneksi()
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Gaji_Staff WHERE UCASE(Nama_Staff) = UCASE(@nama)", Conn)
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Gaji_Staff", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
    End Sub
    Private Sub HitungTotalGaji()
        If IsNumeric(txtjam.Text) AndAlso IsNumeric(txtpokok.Text) Then
            Dim jam As Double = Val(txtjam.Text)
            Dim pokok As Double = Val(txtpokok.Text)
            txttot.Text = (jam * pokok).ToString()
        Else
            txttot.Text = ""
        End If
    End Sub

    Private Sub btninput_Click(sender As Object, e As EventArgs) Handles btninput.Click
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        Call Koneksi()
        If txtkodegaji.Text = "" Or txtnoid.Text = "" Or txtnama.Text = "" Or txtjam.Text = "" Or txtpokok.Text = "" Or txttot.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Gaji_Staff WHERE Kode_Gaji_Staff=@kode", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodegaji.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())

        If count > 0 Then
            MsgBox("Kode gaji sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If
        If Not CekNama(txtnama.Text.Trim()) Then
            MsgBox("Nama staff tidak sesuai dengan database!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        CMD = New OleDbCommand("INSERT INTO Data_Gaji_Staff(Kode_Gaji_Staff, No_ID, Nama_Staff, Jam_Kerja, Gaji_Pokok, Total) " & "VALUES (@kode, @id, @nama, @jam, @pokok, @total)", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodegaji.Text)
        CMD.Parameters.AddWithValue("@id", txtnoid.Text)
        CMD.Parameters.AddWithValue("@nama", txtnama.Text)
        CMD.Parameters.AddWithValue("@jam", txtjam.Text)
        CMD.Parameters.AddWithValue("@pokok", txtpokok.Text)
        CMD.Parameters.AddWithValue("@total", txttot.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        Call ClearInput()
        Conn.Close()
    End Sub

    Private Sub Form13_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        txtnoid.Text = form1.LoginNoID
        tt.ToolTipTitle = "Pengingat"
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.IsBalloon = True
        tt.SetToolTip(txtkodegaji, "Isi kode bahan baku dengan lengkap.")
        Tampildatabase()
    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Gaji_Staff").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As New OleDbCommand("DELETE FROM Data_Gaji_Staff WHERE Kode_Gaji_Staff=@kode", Conn)
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
            txtkodegaji.Text = row.Cells("Kode_Gaji_Staff").Value.ToString()
            txtnoid.Text = row.Cells("No_ID").Value.ToString()
            txtnama.Text = row.Cells("Nama_Staff").Value.ToString()
            txtjam.Text = row.Cells("Jam_Kerja").Value.ToString()
            txtpokok.Text = row.Cells("Gaji_Pokok").Value.ToString()
            txttot.Text = row.Cells("Total").Value.ToString()
        End If
    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            row.Cells("Kode_Gaji_Staff").Value = txtkodegaji.Text
            row.Cells("No_ID").Value = txtnoid.Text
            row.Cells("Nama_Staff").Value = txtnama.Text
            row.Cells("Jam_Kerja").Value = txtjam.Text
            row.Cells("Gaji_Pokok").Value = txtpokok.Text
            row.Cells("Total").Value = txttot.Text
            MsgBox(" Data berhasil diedit !", MsgBoxStyle.Information)
            ClearInput()
        Else
            MsgBox("Pilih data yang ingin diedit dari tabel!", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan kode gaji yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Gaji_Staff WHERE Kode_Gaji_Staff LIKE '%" & txtsearch.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView1.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub txtpokok_TextChanged(sender As Object, e As EventArgs) Handles txtpokok.TextChanged
        HitungTotalGaji()
    End Sub

    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If txtsearch.Text.Trim() = "" Then
            Call Koneksi()
            Tampildatabase()
            Conn.Close()
        End If
    End Sub
    Private Sub txttot_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txttot.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        Call ClearInput()
    End Sub
    Private Sub btnback_Click(sender As Object, e As EventArgs) Handles btnback.Click
        Form12.Show()
        Me.Hide()
        Call ClearInput()
    End Sub
    Private Sub txtjam_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtjam.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtpokok_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtpokok.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
End Class