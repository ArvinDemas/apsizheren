Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form3
    Dim CMD As OleDbCommand
    Dim DR As OleDbDataReader
    Dim tt As New ToolTip()

    Sub ClearInput()
        txtkodebayar.Clear()
        txtnoid.Text = form1.LoginNoID
        txtkodep.Clear()
        txttotbayar.Clear()
        txtsearch.Clear()
    End Sub
    Function CekNama(nama As String) As Boolean
        Call Koneksi()
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pesanan WHERE UCASE(Kode_Pesanan) = UCASE(@kode)", Conn)
        CMD.Parameters.AddWithValue("@kode", nama)
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pembayaran", Conn)
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
        If txtkodebayar.Text = "" Or txtnoid.Text = "" Or txtkodep.Text = "" Or txttotbayar.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pembayaran WHERE Kode_Pembayaran=@kode1", Conn)
        CMD.Parameters.AddWithValue("@kode1", txtkodebayar.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())

        If count > 0 Then
            MsgBox("Kode Pembayaran sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If
        If Not CekNama(txtkodep.Text.Trim()) Then
            MsgBox("Kode pesananan tidak sesuai dengan database!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        Dim totalBayar As Decimal
        If Not Decimal.TryParse(txttotbayar.Text, totalBayar) Then
            MsgBox("Total Pembayaran harus berupa angka!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("INSERT INTO Data_Pembayaran (Kode_Pembayaran, No_ID, Kode_Pesanan, Tanggal_Pembayaran, Total_Pembayaran) " & "VALUES (@kode1, @id, @kode, @tglbayar, @total)", Conn)
        CMD.Parameters.Add("@kode1", OleDbType.VarChar).Value = txtkodebayar.Text
        CMD.Parameters.Add("@id", OleDbType.VarChar).Value = txtnoid.Text
        CMD.Parameters.Add("@kode", OleDbType.VarChar).Value = txtkodep.Text
        CMD.Parameters.Add("@tglbayar", OleDbType.Date).Value = DateTimePicker1.Value
        CMD.Parameters.Add("@total", OleDbType.Decimal).Value = totalBayar
        CMD.ExecuteNonQuery()
        Tampildatabase()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        Call ClearInput()
        Conn.Close()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        Tampildatabase()
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        txtnoid.Text = form1.LoginNoID
        tt.ToolTipTitle = "Pengingat"
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.IsBalloon = True
        tt.SetToolTip(txtkodep, "Isi kode bahan baku dengan lengkap.")
    End Sub

    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Pembayaran").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                Dim CMD As New OleDbCommand("DELETE FROM Data_Pembayaran WHERE Kode_Pembayaran=@kode1", Conn)
                CMD.Parameters.AddWithValue("@kode1", kode)
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
            txtkodebayar.Text = row.Cells("Kode_Pembayaran").Value.ToString()
            txtnoid.Text = row.Cells("No_ID").Value.ToString()
            txtkodep.Text = row.Cells("Kode_Pesanan").Value.ToString()
            If Not IsDBNull(row.Cells("Tanggal_Pembayaran").Value) Then
                DateTimePicker1.Value = CDate(row.Cells("Tanggal_Pembayaran").Value)
            End If
            txttotbayar.Text = row.Cells("Total_Pembayaran").Value.ToString()
        End If
    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        If DataGridView1.SelectedRows.Count > 0 Then
            Dim row As DataGridViewRow = DataGridView1.SelectedRows(0)
            row.Cells("Kode_Pembayaran").Value = txtkodebayar.Text
            row.Cells("No_ID").Value = txtnoid.Text
            row.Cells("Kode_Pesanan").Value = txtkodep.Text
            row.Cells("Tanggal_Pembayaran").Value = DateTimePicker1.Value
            row.Cells("Total_Pembayaran").Value = txttotbayar.Text
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
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pembayaran WHERE Kode_Pesanan LIKE '%" & txtsearch.Text & "%'", Conn)
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
    Private Sub txttotbayar_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txttotbayar.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        Call ClearInput()
        form1.txtnoid.Clear()
    End Sub
    Private Sub btncek_Click(sender As Object, e As EventArgs) Handles btncek.Click
        Form4.Show()
        Me.Hide()
        Call ClearInput()
    End Sub

    Private Sub btnback_Click(sender As Object, e As EventArgs) Handles btnback.Click
        Form2.Show()
        Me.Hide()
        Call ClearInput()
    End Sub


End Class