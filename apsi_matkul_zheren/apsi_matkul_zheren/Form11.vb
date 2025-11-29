Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form11
    Dim CMD As OleDbCommand
    Dim DR As OleDbDataReader
    Dim tt As New ToolTip()

    Sub ClearInput()
        txtkodetp.Clear()
        txtnoid.Text = form1.LoginNoID
        txtnoid2.Text = form1.LoginNoID
        txtnoid3.Text = form1.LoginNoID
        RichTextBox1.Clear()
        txtkontak.Clear()
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Tujuan_Pengiriman", Conn)
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
        If txtkodetp.Text = "" Or txtnoid.Text = "" Or RichTextBox1.Text = "" Or txtkontak.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Tujuan_Pengiriman WHERE Kode_Tujuan_Pengiriman=@kode", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodetp.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())
        If count > 0 Then
            MsgBox("Kode tujuan pengiriman sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If

        CMD = New OleDbCommand("INSERT INTO Data_Tujuan_Pengiriman (Kode_Tujuan_Pengiriman, No_ID, Tempat_Tujuan_Distribusi, Kontak_Penerima) " & "VALUES (@kode, @id, @tempat, @kontak)", Conn)
        CMD.Parameters.AddWithValue("@kode", txtkodetp.Text)
        CMD.Parameters.AddWithValue("@id", txtnoid.Text)
        CMD.Parameters.AddWithValue("@tempat", RichTextBox1.Text)
        CMD.Parameters.AddWithValue("@kontak", txtkontak.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        ClearInput()
    End Sub

    'Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click
    '    If DataGridView1.SelectedRows.Count > 0 Then
    '        Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Tujuan_Pengiriman").Value.ToString()
    '        Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '        If result = DialogResult.Yes Then
    '            Call Koneksi()
    '            Dim CMD As New OleDbCommand("DELETE FROM Data_Tujuan_Pengiriman WHERE Kode_Tujuan_Pengiriman=@kode", Conn)
    '            CMD.Parameters.AddWithValue("@kode", kode)
    '            CMD.ExecuteNonQuery()
    '            Conn.Close()
    '            Tampildatabase()
    '            MsgBox("Data berhasil dihapus dari database!", MsgBoxStyle.Information)
    '        End If
    '    Else
    '        MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
    '    End If
    '    ClearInput()
    'End Sub
    Private Sub btndelete_Click(sender As Object, e As EventArgs) Handles btndelete.Click

        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Dim kode As String = DataGridView1.SelectedRows(0).Cells("Kode_Tujuan_Pengiriman").Value.ToString()
        Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If result = DialogResult.Yes Then
            Call Koneksi()
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Dim CMD As New OleDbCommand("DELETE FROM Data_Tujuan_Pengiriman WHERE Kode_Tujuan_Pengiriman=@kode", Conn)
            CMD.Parameters.AddWithValue("@kode", kode)
            CMD.ExecuteNonQuery()
            Conn.Close()
            Tampildatabase()
            MsgBox("Data berhasil dihapus dari database!", MsgBoxStyle.Information)
        End If
        ClearInput()
    End Sub


    Private Sub btndeleteall_Click(sender As Object, e As EventArgs) Handles btndeleteall.Click
        ClearInput()
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            txtkodetp.Text = row.Cells("Kode_Tujuan_Pengiriman").Value.ToString()
            txtnoid.Text = row.Cells("No_ID").Value.ToString()
            RichTextBox1.Text = row.Cells("Tempat_Tujuan_Distribusi").Value.ToString()
            txtkontak.Text = row.Cells("Kontak_Penerima").Value.ToString()
        End If
    End Sub

    Private Sub btnedit_Click(sender As Object, e As EventArgs) Handles btnedit.Click
        If DataGridView1.SelectedRows.Count = 0 Then
            MsgBox("Pilih data yang ingin diedit dari tabel!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Dim CMD As New OleDbCommand("UPDATE Data_Tujuan_Pengiriman SET No_ID=@id, Tempat_Tujuan_Distribusi=@tujuan, Kontak_Penerima=@kontak WHERE Kode_Tujuan_Pengiriman=@kode", Conn)
        CMD.Parameters.AddWithValue("@id", txtnoid.Text)
        CMD.Parameters.AddWithValue("@tujuan", RichTextBox1.Text)
        CMD.Parameters.AddWithValue("@kontak", txtkontak.Text)
        CMD.Parameters.AddWithValue("@kode", txtkodetp.Text)
        CMD.ExecuteNonQuery()
        MsgBox("Data berhasil diedit!", MsgBoxStyle.Information)
        Tampildatabase()
        ClearInput()
        Conn.Close()
    End Sub


    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan kode pesanan yang mau dicari!", MsgBoxStyle.Exclamation)
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
        form1.txtnoid.Clear()
    End Sub

    Sub ClearInput1()
        txtkodepg.Clear()
        txtnoid.Text = form1.LoginNoID
        txtnoid2.Text = form1.LoginNoID
        txtnoid3.Text = form1.LoginNoID
        txtkodetp2.Clear()
        txtrute.Clear()
        txtstatus.Clear()
        txtsearch2.Clear()
    End Sub
    Function CekNama1(nama As String) As Boolean
        Call Koneksi()
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Tujuan_Pengiriman WHERE UCASE(Kode_Tujuan_Pengiriman) = UCASE(@kode1)", Conn)
        CMD.Parameters.AddWithValue("@kode1", nama)
        Dim jumlah As Integer = Convert.ToInt32(CMD.ExecuteScalar())
        Return (jumlah > 0)
    End Function
    Sub Tampildatabase1()
        Try
            Call Koneksi()
            If Conn Is Nothing Then
                MsgBox("Koneksi masih kosong!")
                Exit Sub
            End If
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pengiriman", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView2.DataSource = Nothing
            DataGridView2.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        Finally
            If Conn.State = ConnectionState.Open Then
                Conn.Close()
            End If
        End Try
    End Sub

    Private Sub btninput2_Click(sender As Object, e As EventArgs) Handles btninput2.Click
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        If txtkodepg.Text = "" Or txtnoid2.Text = "" Or txtkodetp2.Text = "" Or txtrute.Text = "" Or txtstatus.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pengiriman WHERE Kode_Pengiriman=@kode2", Conn)
        CMD.Parameters.AddWithValue("@kode2", txtkodepg.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())
        If count > 0 Then
            MsgBox("Kode pengiriman sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If
        If Not CekNama1(txtkodetp2.Text.Trim()) Then
            MsgBox("Nama kode tujuan pengiriman tidak sesuai dengan database!", MsgBoxStyle.Critical)
            Exit Sub
        End If
        CMD = New OleDbCommand("INSERT INTO Data_Pengiriman (Kode_Pengiriman, No_ID, Kode_Tujuan_Pengiriman, Rute_Pengiriman, Status_Pengiriman) " & "VALUES (@kode2, @id1, @kode1, @rute, @status)", Conn)
        CMD.Parameters.AddWithValue("@kode2", txtkodepg.Text)
        CMD.Parameters.AddWithValue("@id1", txtnoid2.Text)
        CMD.Parameters.AddWithValue("@kode1", txtkodetp2.Text)
        CMD.Parameters.AddWithValue("@rute", txtrute.Text)
        CMD.Parameters.AddWithValue("@status", txtstatus.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase1()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        ClearInput()
        Call ClearInput1()
    End Sub


    'Private Sub btndelete2_Click(sender As Object, e As EventArgs) Handles btndelete2.Click
    '    If DataGridView2.SelectedRows.Count > 0 Then
    '        Dim kode As String = DataGridView2.SelectedRows(0).Cells("Kode_Pengiriman").Value.ToString()
    '        Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '        If result = DialogResult.Yes Then
    '            Call Koneksi()
    '            Dim CMD As New OleDbCommand("DELETE FROM Data_Pengiriman WHERE Kode_Pengiriman=@kode2", Conn)
    '            CMD.Parameters.AddWithValue("@kode2", kode)
    '            CMD.ExecuteNonQuery()
    '            Conn.Close()
    '            Tampildatabase1()
    '            MsgBox("Data berhasil dihapus dari database!", MsgBoxStyle.Information)
    '        End If
    '    Else
    '        MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
    '    End If
    '    ClearInput1()
    'End Sub
    Private Sub btndelete2_Click(sender As Object, e As EventArgs) Handles btndelete2.Click
        If DataGridView2.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView2.SelectedRows(0).Cells("Kode_Pengiriman").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If
                Dim CMD As New OleDbCommand("DELETE FROM Data_Pengiriman WHERE Kode_Pengiriman=@kode2", Conn)
                CMD.Parameters.AddWithValue("@kode2", kode)
                CMD.ExecuteNonQuery()
                Conn.Close()
                Tampildatabase1()
                MsgBox("Data berhasil dihapus!", MsgBoxStyle.Information)
            End If
        Else
            MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
        End If
        ClearInput1()
    End Sub


    Private Sub btndeleteall2_Click(sender As Object, e As EventArgs) Handles btndeleteall2.Click
        Call ClearInput1()
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView2.Rows(e.RowIndex)
            txtkodepg.Text = row.Cells("Kode_Pengiriman").Value.ToString()
            txtnoid2.Text = row.Cells("No_ID").Value.ToString()
            txtkodetp2.Text = row.Cells("Kode_Tujuan_Pengiriman").Value.ToString()
            txtrute.Text = row.Cells("Rute_Pengiriman").Value.ToString()
            txtstatus.Text = row.Cells("Status_Pengiriman").Value.ToString()
        End If
    End Sub

    Private Sub btnedit2_Click(sender As Object, e As EventArgs) Handles btnedit2.Click
        If DataGridView2.SelectedRows.Count = 0 Then
            MsgBox("Pilih data yang ingin diedit dari tabel!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Dim CMD As New OleDbCommand("UPDATE Data_Pengiriman SET No_ID=@id1, Kode_Tujuan_Pengiriman=@kode1, Rute_Pengiriman=@rute, Status_Pengiriman=@status WHERE Kode_Pengiriman=@kode2", Conn)
        CMD.Parameters.AddWithValue("@id1", txtnoid2.Text)
        CMD.Parameters.AddWithValue("@kode1", txtkodetp2.Text)
        CMD.Parameters.AddWithValue("@rute", txtrute.Text)
        CMD.Parameters.AddWithValue("@status", txtstatus.Text)
        CMD.Parameters.AddWithValue("@kode2", txtkodepg.Text)
        CMD.ExecuteNonQuery()
        MsgBox("Data berhasil diedit!", MsgBoxStyle.Information)
        Tampildatabase1()
        ClearInput()
        Conn.Close()
    End Sub

    Private Sub btnsearch2_Click(sender As Object, e As EventArgs) Handles btnsearch2.Click
        If txtsearch2.Text = "" Then
            MsgBox("Masukkan kode pesanan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pengiriman WHERE Kode_Pengiriman LIKE '%" & txtsearch2.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView2.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub

    Private Sub txtsearch2_TextChanged(sender As Object, e As EventArgs) Handles txtsearch2.TextChanged
        If txtsearch2.Text.Trim() = "" Then
            Call Koneksi()
            Tampildatabase1()
            Conn.Close()
        End If
    End Sub

    Sub ClearInput2()
        txtkodepm.Clear()
        txtnoid.Text = form1.LoginNoID
        txtnoid2.Text = form1.LoginNoID
        txtnoid3.Text = form1.LoginNoID
        cmbjenispm.SelectedIndex = -1
        txtmodel.Clear()
        txtstatuspm.Clear()
        txtsearch2.Clear()
    End Sub
    'Function CekNama2(nama As String) As Boolean
    '    Call Koneksi()
    '    CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pengemasan WHERE UCASE(Kode_Tujuan_Pengiriman) = UCASE(@kode1)", Conn)
    '    CMD.Parameters.AddWithValue("@kode1", nama)
    '    Dim jumlah As Integer = Convert.ToInt32(CMD.ExecuteScalar())
    '    Return (jumlah > 0)
    'End Function
    Sub Tampildatabase2()
        Try
            Call Koneksi()
            If Conn Is Nothing Then
                MsgBox("Koneksi masih kosong!")
                Exit Sub
            End If
            If Conn.State = ConnectionState.Closed Then
                Conn.Open()
            End If
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pengemasan", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView3.DataSource = Nothing
            DataGridView3.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
    End Sub
    Private Sub IsiComboJenis()
        Dim Conn As New OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;Data Source=" & Application.StartupPath & "\M4_TUBES.mdb")
        Conn.Open()
        Dim CMD As New OleDbCommand("SELECT DISTINCT Jenis_Pengemasan FROM Data_Pengemasan", Conn)
        Dim DR As OleDbDataReader = CMD.ExecuteReader()
        cmbjenispm.Items.Clear()
        While DR.Read()
            If Not IsDBNull(DR("Jenis_Pengemasan")) Then
                cmbjenispm.Items.Add(DR("Jenis_Pengemasan").ToString)
            End If
        End While
        DR.Close()
        Conn.Close()
    End Sub

    Private Sub btninput3_Click(sender As Object, e As EventArgs) Handles btninput3.Click
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then
            Conn.Open()
        End If
        If txtkodepm.Text = "" Or txtnoid3.Text = "" Or cmbjenispm.Text = "" Or txtmodel.Text = "" Or txtstatuspm.Text = "" Then
            MsgBox("Lengkapi semua data!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        CMD = New OleDbCommand("SELECT COUNT(*) FROM Data_Pengemasan WHERE Kode_Pengemasan=@kode3", Conn)
        CMD.Parameters.AddWithValue("@kode3", txtkodepm.Text)
        Dim count As Integer = Convert.ToInt32(CMD.ExecuteScalar())
        If count > 0 Then
            MsgBox("Kode pengemasan sudah ada, gunakan kode lain!", MsgBoxStyle.Exclamation)
            Conn.Close()
            Exit Sub
        End If
        'If Not CekNama1(txtkodetp2.Text.Trim()) Then
        '    MsgBox("Nama kode tujuan pengiriman tidak sesuai dengan database!", MsgBoxStyle.Critical)
        '    Exit Sub
        'End If
        CMD = New OleDbCommand("INSERT INTO Data_Pengemasan (Kode_Pengemasan, No_ID, Jenis_Pengemasan, Model_Kemasan, Status_Pengemasan) " & "VALUES (@kode3, @id2, @jenispm, @model, @statuspm)", Conn)
        CMD.Parameters.AddWithValue("@kode3", txtkodepm.Text)
        CMD.Parameters.AddWithValue("@id2", txtnoid3.Text)
        CMD.Parameters.AddWithValue("@jenispm", cmbjenispm.Text)
        CMD.Parameters.AddWithValue("@model", txtmodel.Text)
        CMD.Parameters.AddWithValue("@statuspm", txtstatuspm.Text)
        CMD.ExecuteNonQuery()
        Tampildatabase2()
        MsgBox("Data berhasil ditambahkan!", MsgBoxStyle.Information)
        If Conn.State = ConnectionState.Open Then
            Conn.Close()
        End If
        ClearInput()
        Call ClearInput1()
    End Sub

    Private Sub btndelete3_Click(sender As Object, e As EventArgs) Handles btndelete3.Click
        If DataGridView3.SelectedRows.Count > 0 Then
            Dim kode As String = DataGridView3.SelectedRows(0).Cells("Kode_Pengemasan").Value.ToString()
            Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.Yes Then
                Call Koneksi()
                If Conn.State = ConnectionState.Closed Then
                    Conn.Open()
                End If
                Dim CMD As New OleDbCommand("DELETE FROM Data_Pengemasan WHERE Kode_Pengemasan=@kode3", Conn)
                CMD.Parameters.AddWithValue("@kode3", kode)
                CMD.ExecuteNonQuery()
                Conn.Close()
                Tampildatabase2()
                MsgBox("Data berhasil dihapus dari database!", MsgBoxStyle.Information)
            End If
        Else
            MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
        End If
        ClearInput2()
    End Sub

    'Private Sub btndelete2_Click(sender As Object, e As EventArgs) Handles btndelete2.Click
    '    If DataGridView2.SelectedRows.Count > 0 Then
    '        Dim kode As String = DataGridView2.SelectedRows(0).Cells("Kode_Pengiriman").Value.ToString()
    '        Dim result As DialogResult = MessageBox.Show("Apakah Anda yakin ingin menghapus data dengan Kode: " & kode & "?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
    '        If result = DialogResult.Yes Then
    '            Call Koneksi()
    '            If Conn.State = ConnectionState.Closed Then
    '                Conn.Open()
    '            End If
    '            Dim CMD As New OleDbCommand("DELETE FROM Data_Pengiriman WHERE Kode_Pengiriman=@kode2", Conn)
    '            CMD.Parameters.AddWithValue("@kode2", kode)
    '            CMD.ExecuteNonQuery()
    '            Conn.Close()
    '            Tampildatabase1()
    '            MsgBox("Data berhasil dihapus!", MsgBoxStyle.Information)
    '        End If
    '    Else
    '        MsgBox("Pilih baris yang ingin dihapus!", MsgBoxStyle.Exclamation)
    '    End If
    '    ClearInput1()
    'End Sub

    Private Sub btndeleteall3_Click(sender As Object, e As EventArgs) Handles btndeleteall3.Click
        Call ClearInput2()
    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView3.Rows(e.RowIndex)
            txtkodepm.Text = row.Cells("Kode_Pengemasan").Value.ToString()
            txtnoid3.Text = row.Cells("No_ID").Value.ToString()
            cmbjenispm.Text = row.Cells("Jenis_Pengemasan").Value.ToString()
            txtmodel.Text = row.Cells("Model_Kemasan").Value.ToString()
            txtstatuspm.Text = row.Cells("Status_Pengemasan").Value.ToString()
        End If
    End Sub

    Private Sub btnedit3_Click(sender As Object, e As EventArgs) Handles btnedit3.Click
        If DataGridView3.SelectedRows.Count = 0 Then
            MsgBox("Pilih data yang ingin diedit dari tabel!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Dim CMD As New OleDbCommand("UPDATE Data_Pengemasan SET No_ID=@id, Jenis_Pengemasan=@jenis, Model_Kemasan=@model, Status_Pengemasan=@statuspm WHERE Kode_Pengemasan=@kode3", Conn)
        CMD.Parameters.AddWithValue("@id", txtnoid3.Text)
        CMD.Parameters.AddWithValue("@jenis", cmbjenispm.Text)
        CMD.Parameters.AddWithValue("@model", txtmodel.Text)
        CMD.Parameters.AddWithValue("@statuspm", txtstatuspm.Text)
        CMD.Parameters.AddWithValue("@kode3", txtkodepm.Text)
        CMD.ExecuteNonQuery()
        'If DataGridView3.SelectedRows.Count > 0 Then
        '    Dim row As DataGridViewRow = DataGridView3.SelectedRows(0)
        '    row.Cells("Kode_Pengemasan").Value = txtkodepm.Text
        '    row.Cells("No_ID").Value = txtnoid3.Text
        '    row.Cells("Jenis_Pengemasan").Value = cmbjenispm.Text
        '    row.Cells("Model_Pengemasan").Value = txtmodel.Text
        '    row.Cells("Status_Pengemasan").Value = txtstatuspm.Text
        MsgBox("Data berhasil diedit !", MsgBoxStyle.Information)
        Tampildatabase2()
        ClearInput()
        Conn.Close()
    End Sub



    Private Sub btnsearch3_Click(sender As Object, e As EventArgs) Handles btnsearch3.Click
        If txtsearch3.Text = "" Then
            MsgBox("Masukkan kode pengemasan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pengemasan WHERE Kode_Pengemasan LIKE '%" & txtsearch3.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView3.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub

    Private Sub txtsearch3_TextChanged(sender As Object, e As EventArgs) Handles txtsearch3.TextChanged
        If txtsearch3.Text.Trim() = "" Then
            Call Koneksi()
            Tampildatabase2()
            Conn.Close()
        End If
    End Sub

    Private Sub Form11_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        txtnoid.Text = form1.LoginNoID
        txtnoid2.Text = form1.LoginNoID
        txtnoid3.Text = form1.LoginNoID
        tt.ToolTipTitle = "Pengingat"
        tt.ToolTipIcon = ToolTipIcon.Info
        tt.IsBalloon = True
        txtkontak.Text = "(+62)"
        txtkodetp2.Text = txtkodetp.Text
        tt.SetToolTip(txtkodetp, "Isi kode tujuan pengiriman dengan lengkap.")
        tt.SetToolTip(txtkodepg, "Isi kode pengiriman dengan lengkap.")
        tt.SetToolTip(txtkodepm, "Isi kode pengemasan dengan lengkap.")
        IsiComboJenis()
        Tampildatabase()
        Tampildatabase1()
        Tampildatabase2()
    End Sub

    Private Sub txtkontak_TextChanged(sender As Object, e As EventArgs) Handles txtkontak.TextChanged
        If Not txtkontak.Text.StartsWith("(+62) ") Then
            txtkontak.Text = "(+62) "
            txtkontak.SelectionStart = txtkontak.Text.Length
        End If
    End Sub

    Private Sub txtrute_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtrute.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> " "c Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtstatus_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtstatus.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> " "c Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtkontak_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtkontak.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
End Class