Imports System.Data.OleDb
Imports System.Diagnostics
Public Class Form12
    Private Sub Form12_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Size = New Size(920, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Tampildatabase()
        Tampildatabase1()
        Tampildatabase2()
        Tampildatabase3()
        Tampildatabase4()
        DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView6.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView7.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
        DataGridView8.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        For Each col As DataGridViewColumn In DataGridView1.Columns
            col.MinimumWidth = 100
        Next
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Akun", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView1.DataSource = Nothing
            DataGridView1.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan no id yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Akun WHERE No_ID LIKE '%" & txtsearch.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView1.DataSource = Dt
        Else
            MsgBox("Data dengan id tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        txtsearch.Clear()
    End Sub

    Private Sub btngaji_Click(sender As Object, e As EventArgs) Handles btngaji.Click
        Form13.Show()
        Me.Hide()
        txtsearch.Clear()
    End Sub

    Private Sub btneoq_Click(sender As Object, e As EventArgs) Handles btneoq.Click
        Form10.Show()
        Me.Hide()
        txtsearch.Clear()
    End Sub
    Private Sub btnshowdata_Click(sender As Object, e As EventArgs) Handles btnshowdata.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT No_ID, Username, Posisi, Password FROM Data_Akun", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Akun")
        DataGridView1.DataSource = Ds.Tables("Akun")
        DataGridView1.ReadOnly = True
        DataGridView1.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch_TextChanged(sender As Object, e As EventArgs) Handles txtsearch.TextChanged
        If txtsearch.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT No_ID, Username, Posisi, Password FROM Data_Akun", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Akun")
            DataGridView1.DataSource = Ds.Tables("Akun")
            DataGridView1.ReadOnly = True
            DataGridView1.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If DataGridView1.Rows.Count = 0 Then
            MessageBox.Show("Tidak ada data untuk diverifikasi.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CheckBox1.Checked = False
            Return
        End If
        If CheckBox1.Checked Then
            Dim validCount As Integer = 0
            Dim invalidCount As Integer = 0
            For Each row As DataGridViewRow In DataGridView1.Rows
                If row.IsNewRow Then Continue For
                If row.Cells("No_ID").Value IsNot Nothing AndAlso
                row.Cells("Username").Value IsNot Nothing AndAlso
                row.Cells("Posisi").Value IsNot Nothing AndAlso
                row.Cells("Password").Value IsNot Nothing Then
                    row.DefaultCellStyle.BackColor = Color.LightGreen
                    validCount += 1
                Else
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                    invalidCount += 1
                End If
            Next
            MessageBox.Show($"Verifikasi selesai: {validCount} baris valid, {invalidCount} baris invalid.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            For Each row As DataGridViewRow In DataGridView1.Rows
                row.DefaultCellStyle.BackColor = Color.White
            Next
        End If
    End Sub

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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Perencanaan_Produksi", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView2.DataSource = Nothing
            DataGridView2.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch2_Click(sender As Object, e As EventArgs) Handles btnsearch2.Click
        If txtsearch2.Text = "" Then
            MsgBox("Masukkan kode perencanaan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Perencanaan_Produksi WHERE Kode_Perencanaan_Produksi LIKE '%" & txtsearch2.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView2.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata2_Click(sender As Object, e As EventArgs) Handles btnshowdata2.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Perencanaan_Produksi, No_ID, Kode_Bahan_Baku, Jenis_Bahan_Baku, Nama_Bahan_Baku,Nama_Produk, Total_Produksi FROM Data_Perencanaan_Produksi", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Perencanaan")
        DataGridView2.DataSource = Ds.Tables("Perencanaan")
        DataGridView2.ReadOnly = True
        DataGridView2.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch2_TextChanged(sender As Object, e As EventArgs) Handles txtsearch2.TextChanged
        If txtsearch2.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Perencanaan_Produksi, No_ID, Kode_Bahan_Baku, Jenis_Bahan_Baku, Nama_Bahan_Baku, Nama_Produk, Total_Produksi FROM Data_Perencanaan_Produksi", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Perencanaan")
            DataGridView2.DataSource = Ds.Tables("Perencanaan")
            DataGridView2.ReadOnly = True
            DataGridView2.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        If DataGridView2.Rows.Count = 0 Then
            MessageBox.Show("Tidak ada data untuk diverifikasi.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CheckBox2.Checked = False
            Return
        End If
        If CheckBox2.Checked Then
            Dim validCount As Integer = 0
            Dim invalidCount As Integer = 0
            For Each row As DataGridViewRow In DataGridView2.Rows
                If row.IsNewRow Then Continue For
                If row.Cells("Kode_Perencanaan_Produksi").Value IsNot Nothing AndAlso
                row.Cells("No_ID").Value IsNot Nothing AndAlso
                row.Cells("Kode_Bahan_Baku").Value IsNot Nothing AndAlso
                row.Cells("Jenis_Bahan_Baku").Value IsNot Nothing AndAlso
                row.Cells("Nama_Bahan_Baku").Value IsNot Nothing AndAlso
                row.Cells("Nama_Produk").Value IsNot Nothing AndAlso
                row.Cells("Total_Produksi").Value IsNot Nothing Then
                    row.DefaultCellStyle.BackColor = Color.LightGreen
                    validCount += 1
                Else
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                    invalidCount += 1
                End If
            Next
            MessageBox.Show($"Verifikasi selesai: {validCount} baris valid, {invalidCount} baris invalid.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            For Each row As DataGridViewRow In DataGridView2.Rows
                row.DefaultCellStyle.BackColor = Color.White
            Next
        End If
    End Sub
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
            Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Request_Customer", Conn)
            Dim Dt As New DataTable
            Da.Fill(Dt)
            DataGridView3.DataSource = Nothing
            DataGridView3.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch3_Click(sender As Object, e As EventArgs) Handles btnsearch3.Click
        If txtsearch3.Text = "" Then
            MsgBox("Masukkan kode perencanaan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Request_Customer WHERE Kode_Request_Customer LIKE '%" & txtsearch3.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView3.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata3_Click(sender As Object, e As EventArgs) Handles btnshowdata3.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Request_Customer, No_ID, Kode_Pesanan, Nama_Produk, Request_Desain FROM Data_Request_Customer", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Request")
        DataGridView3.DataSource = Ds.Tables("Request")
        DataGridView3.ReadOnly = True
        DataGridView3.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch3_TextChanged(sender As Object, e As EventArgs) Handles txtsearch3.TextChanged
        If txtsearch3.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Request_Customer, No_ID, Kode_Pesanan, Nama_Produk, Request_Desain FROM Data_Request_Customer", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Request")
            DataGridView3.DataSource = Ds.Tables("Request")
            DataGridView3.ReadOnly = True
            DataGridView3.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        If DataGridView3.Rows.Count = 0 Then
            MessageBox.Show("Tidak ada data untuk diverifikasi.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CheckBox3.Checked = False
            Return
        End If
        If CheckBox3.Checked Then
            Dim validCount As Integer = 0
            Dim invalidCount As Integer = 0
            For Each row As DataGridViewRow In DataGridView3.Rows
                If row.IsNewRow Then Continue For
                If row.Cells("Kode_Request_Customer").Value IsNot Nothing AndAlso
                row.Cells("No_ID").Value IsNot Nothing AndAlso
                row.Cells("Kode_Pesanan").Value IsNot Nothing AndAlso
                row.Cells("Nama_Produk").Value IsNot Nothing AndAlso
                row.Cells("Request_Desain").Value IsNot Nothing Then
                    row.DefaultCellStyle.BackColor = Color.LightGreen
                    validCount += 1
                Else
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                    invalidCount += 1
                End If
            Next
            MessageBox.Show($"Verifikasi selesai: {validCount} baris valid, {invalidCount} baris invalid.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            For Each row As DataGridViewRow In DataGridView3.Rows
                row.DefaultCellStyle.BackColor = Color.White
            Next
        End If
    End Sub
    Sub Tampildatabase3()
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
            DataGridView4.DataSource = Nothing
            DataGridView4.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch4_Click(sender As Object, e As EventArgs) Handles btnsearch4.Click
        If txtsearch4.Text = "" Then
            MsgBox("Masukkan kode pengiriman yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pengiriman WHERE Kode_Pengiriman LIKE '%" & txtsearch4.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView4.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata4_Click(sender As Object, e As EventArgs) Handles btnshowdata4.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Pengiriman, No_ID, Kode_Tujuan_Pengiriman, Rute_Pengiriman, Status_Pengiriman FROM Data_Pengiriman", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Pengiriman")
        DataGridView4.DataSource = Ds.Tables("Pengiriman")
        DataGridView4.ReadOnly = True
        DataGridView4.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch4_TextChanged(sender As Object, e As EventArgs) Handles txtsearch4.TextChanged
        If txtsearch4.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Pengiriman, No_ID, Kode_Tujuan_Pengiriman, Rute_Pengiriman, Status_Pengiriman FROM Data_Pengiriman", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Pengiriman")
            DataGridView4.DataSource = Ds.Tables("Pengiriman")
            DataGridView4.ReadOnly = True
            DataGridView4.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        If DataGridView4.Rows.Count = 0 Then
            MessageBox.Show("Tidak ada data untuk diverifikasi.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CheckBox4.Checked = False
            Return
        End If
        If CheckBox4.Checked Then
            Dim validCount As Integer = 0
            Dim invalidCount As Integer = 0
            For Each row As DataGridViewRow In DataGridView4.Rows
                If row.IsNewRow Then Continue For
                If row.Cells("Kode_Pengiriman").Value IsNot Nothing AndAlso
                row.Cells("No_ID").Value IsNot Nothing AndAlso
                row.Cells("Kode_Tujuan_Pengiriman").Value IsNot Nothing AndAlso
                row.Cells("Rute_Pengiriman").Value IsNot Nothing AndAlso
                row.Cells("Status_Pengiriman").Value IsNot Nothing Then
                    row.DefaultCellStyle.BackColor = Color.LightGreen
                    validCount += 1
                Else
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                    invalidCount += 1
                End If
            Next
            MessageBox.Show($"Verifikasi selesai: {validCount} baris valid, {invalidCount} baris invalid.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            For Each row As DataGridViewRow In DataGridView4.Rows
                row.DefaultCellStyle.BackColor = Color.White
            Next
        End If
    End Sub
    Sub Tampildatabase4()
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
            DataGridView5.DataSource = Nothing
            DataGridView5.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch5_Click(sender As Object, e As EventArgs) Handles btnsearch5.Click
        If txtsearch5.Text = "" Then
            MsgBox("Masukkan kode tujuan pengiriman yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Tujuan_Pengiriman WHERE Kode_Tujuan_Pengiriman LIKE '%" & txtsearch5.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView5.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata5_Click(sender As Object, e As EventArgs) Handles btnshowdata5.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Tujuan_Pengiriman, No_ID, Tempat_Tujuan_Distribusi, Kontak_Penerima FROM Data_Tujuan_Pengiriman", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Tujuan")
        DataGridView5.DataSource = Ds.Tables("Tujuan")
        DataGridView5.ReadOnly = True
        DataGridView5.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch5_TextChanged(sender As Object, e As EventArgs) Handles txtsearch5.TextChanged
        If txtsearch5.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Tujuan_Pengiriman, No_ID, Tempat_Tujuan_Distribusi, Kontak_Penerima FROM Data_Tujuan_Pengiriman", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Tujuan")
            DataGridView5.DataSource = Ds.Tables("Tujuan")
            DataGridView5.ReadOnly = True
            DataGridView5.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        If DataGridView5.Rows.Count = 0 Then
            MessageBox.Show("Tidak ada data untuk diverifikasi.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
            CheckBox5.Checked = False
            Return
        End If
        If CheckBox5.Checked Then
            Dim validCount As Integer = 0
            Dim invalidCount As Integer = 0
            For Each row As DataGridViewRow In DataGridView5.Rows
                If row.IsNewRow Then Continue For
                If row.Cells("Kode_Tujuan_Pengiriman").Value IsNot Nothing AndAlso
                row.Cells("No_ID").Value IsNot Nothing AndAlso
                row.Cells("Tempat_Tujuan_Distribusi").Value IsNot Nothing AndAlso
                row.Cells("Kontak_Penerima").Value IsNot Nothing Then
                    row.DefaultCellStyle.BackColor = Color.LightGreen
                    validCount += 1
                Else
                    row.DefaultCellStyle.BackColor = Color.LightCoral
                    invalidCount += 1
                End If
            Next
            MessageBox.Show($"Verifikasi selesai: {validCount} baris valid, {invalidCount} baris invalid.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            For Each row As DataGridViewRow In DataGridView5.Rows
                row.DefaultCellStyle.BackColor = Color.White
            Next
        End If
    End Sub
    Sub Tampildatabase5()
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
            DataGridView6.DataSource = Nothing
            DataGridView6.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch6_Click(sender As Object, e As EventArgs) Handles btnsearch6.Click
        If txtsearch6.Text = "" Then
            MsgBox("Masukkan kode pembayaran yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pembayaran WHERE Kode_Pembayaran LIKE '%" & txtsearch6.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView6.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata6_Click(sender As Object, e As EventArgs) Handles btnshowdata6.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Pembayaran, No_ID, Kode_Pesanan, Tanggal_Pembayaran, Total_Pembayaran FROM Data_Pembayaran", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Pembayaran")
        DataGridView6.DataSource = Ds.Tables("Pembayaran")
        DataGridView6.ReadOnly = True
        DataGridView6.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch6_TextChanged(sender As Object, e As EventArgs) Handles txtsearch6.TextChanged
        If txtsearch6.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Pembayaran, No_ID, Kode_Pesanan, Tanggal_Pembayaran, Total_Pembayaran FROM Data_Pembayaran", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Pembayaran")
            DataGridView6.DataSource = Ds.Tables("Pembayaran")
            DataGridView6.ReadOnly = True
            DataGridView6.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub
    Sub Tampildatabase6()
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
            DataGridView7.DataSource = Nothing
            DataGridView7.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch7_Click(sender As Object, e As EventArgs) Handles btnsearch7.Click
        If txtsearch7.Text = "" Then
            MsgBox("Masukkan kode bahan baku yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Bahan_Baku WHERE Kode_Bahan_Baku LIKE '%" & txtsearch7.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView7.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata7_Click(sender As Object, e As EventArgs) Handles btnshowdata7.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Bahan_Baku, No_ID, Jenis_Bahan_Baku, Nama_Bahan_Baku, Total_Bahan_Baku FROM Data_Bahan_Baku", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Bahan")
        DataGridView7.DataSource = Ds.Tables("Bahan")
        DataGridView7.ReadOnly = True
        DataGridView7.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch7_TextChanged(sender As Object, e As EventArgs) Handles txtsearch7.TextChanged
        If txtsearch7.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Bahan_Baku, No_ID, Jenis_Bahan_Baku, Nama_Bahan_Baku, Total_Bahan_Baku FROM Data_Bahan_Baku", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Bahan")
            DataGridView7.DataSource = Ds.Tables("Bahan")
            DataGridView7.ReadOnly = True
            DataGridView7.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub
    Sub Tampildatabase7()
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
            DataGridView8.DataSource = Nothing
            DataGridView8.DataSource = Dt
        Catch ex As Exception
            MsgBox("Error di tampilandatabase: " & ex.Message)
        End Try
        Conn.Close()
    End Sub
    Private Sub btnsearch8_Click(sender As Object, e As EventArgs) Handles btnsearch8.Click
        If txtsearch8.Text = "" Then
            MsgBox("Masukkan kode pengemasan yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Pengemasan WHERE Kode_Pengemasan LIKE '%" & txtsearch8.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView8.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnshowdata8_Click(sender As Object, e As EventArgs) Handles btnshowdata8.Click
        Call Koneksi()
        Da = New OleDbDataAdapter("SELECT Kode_Pengemasan, No_ID, Jenis_Pengemasan, Model_Kemasan, Status_Pengemasan FROM Data_Pengemasan", Conn)
        Ds = New DataSet()
        Da.Fill(Ds, "Pengemasan")
        DataGridView8.DataSource = Ds.Tables("Pengemasan")
        DataGridView8.ReadOnly = True
        DataGridView8.AllowUserToAddRows = False
        Conn.Close()
    End Sub
    Private Sub txtsearch8_TextChanged(sender As Object, e As EventArgs) Handles txtsearch8.TextChanged
        If txtsearch8.Text.Trim() = "" Then
            Call Koneksi()
            Da = New OleDbDataAdapter("SELECT Kode_Pengemasan, No_ID, Jenis_Pengemasan, Model_Kemasan, Status_Pengemasan FROM Data_Pengemasan", Conn)
            Ds = New DataSet()
            Da.Fill(Ds, "Pengemasan")
            DataGridView8.DataSource = Ds.Tables("Pengemasan")
            DataGridView8.ReadOnly = True
            DataGridView8.AllowUserToAddRows = False
            Conn.Close()
        End If
    End Sub
End Class