Imports System.Data.OleDb
Imports System.Diagnostics

Public Class Form10
    Private hasilDt As DataTable

    Private Sub Form10_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.StartPosition = FormStartPosition.CenterScreen
        hasilDt = New DataTable()
        hasilDt.Columns.Add("Nama_Bahan_Baku", GetType(String))
        hasilDt.Columns.Add("Biaya_Pesan", GetType(Double))
        hasilDt.Columns.Add("Biaya_Simpan", GetType(Double))
        hasilDt.Columns.Add("Lead_Time", GetType(Double))
        hasilDt.Columns.Add("Demand", GetType(Double))
        hasilDt.Columns.Add("EOQ", GetType(Double))
        hasilDt.Columns.Add("ROP", GetType(Double))
        hasilDt.Columns.Add("Safety_Stock", GetType(Double))
        DataGridView2.DataSource = hasilDt

        Tampildatabase()
    End Sub
    Sub ClearInput()
        txtnamabb.Clear()
        txtbiaya.Clear()
        txtsimpan.Clear()
        txtdemand.Clear()
        txtlead.Clear()
        txteoq.Clear()
        txtrop.Clear()
        txtss.Clear()
    End Sub
    Sub Tampildatabase()
        Try
            Call Koneksi()
            If Conn.State = ConnectionState.Closed Then Conn.Open()
            Dim sql As String = "SELECT Kode_Perencanaan_Produksi, No_ID, Kode_Bahan_Baku, Jenis_Bahan_Baku, Nama_Bahan_Baku, Nama_Produk, Total_Produksi FROM Data_Perencanaan_Produksi"
            Dim da As New OleDbDataAdapter(sql, Conn)
            Dim dt As New DataTable()
            da.Fill(dt)
            DataGridView1.DataSource = dt
            Conn.Close()
        Catch ex As Exception
            MsgBox("Gagal load Data_Perencanaan_Produksi: " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub txtnamabb_Leave(sender As Object, e As EventArgs) Handles txtnamabb.Leave
        If txtnamabb.Text.Trim() = "" Then
            txtdemand.Text = ""
            Return
        End If
        Try
            Call Koneksi()
            If Conn.State = ConnectionState.Closed Then Conn.Open()

            Dim cekCmd As New OleDbCommand("SELECT COUNT(*) FROM Data_Perencanaan_Produksi WHERE UCASE(Nama_Bahan_Baku) = UCASE(?)", Conn)
            cekCmd.Parameters.AddWithValue("?", txtnamabb.Text.Trim())
            Dim ada As Integer = Convert.ToInt32(cekCmd.ExecuteScalar())
            If ada = 0 Then
                MsgBox("Nama bahan baku tidak ditemukan di Data_Perencanaan_Produksi. Mohon masukkan nama yang benar.", MsgBoxStyle.Exclamation)
                txtdemand.Text = ""
                Conn.Close()
                Return
            End If
            Dim cmd As New OleDbCommand("SELECT SUM(Total_Produksi) FROM Data_Perencanaan_Produksi WHERE UCASE(Nama_Bahan_Baku)=UCASE(?)", Conn)
            cmd.Parameters.AddWithValue("?", txtnamabb.Text.Trim())
            Dim obj = cmd.ExecuteScalar()
            Dim totalDemand As Double = 0
            If Not IsDBNull(obj) AndAlso obj IsNot Nothing Then
                Double.TryParse(obj.ToString(), totalDemand)
            End If
            txtdemand.Text = totalDemand.ToString("F0")
            Conn.Close()
        Catch ex As Exception
            MsgBox("Error mengambil demand: " & ex.Message, MsgBoxStyle.Critical)
            Try
                Conn.Close()
            Catch ex2 As Exception
            End Try
        End Try
    End Sub
    Private Sub btnsearch_Click(sender As Object, e As EventArgs) Handles btnsearch.Click
        If txtsearch.Text = "" Then
            MsgBox("Masukkan nama bahan baku yang mau dicari!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If
        Call Koneksi()
        Dim Da As New OleDbDataAdapter("SELECT * FROM Data_Perencanaan_Produksi WHERE Nama_Bahan_Baku LIKE '%" & txtsearch.Text & "%'", Conn)
        Dim Dt As New DataTable
        Da.Fill(Dt)
        If Dt.Rows.Count > 0 Then
            DataGridView2.DataSource = Dt
        Else
            MsgBox("Data dengan kode tersebut tidak ditemukan di grid!", MsgBoxStyle.Information)
        End If
        Conn.Close()
    End Sub
    Private Sub btnHitung_Click(sender As Object, e As EventArgs) Handles btnHitung.Click
        If txtnamabb.Text.Trim() = "" Then
            MsgBox("Masukkan nama bahan baku terlebih dahulu.", MsgBoxStyle.Exclamation)
            txtnamabb.Focus()
            Return
        End If
        Dim nama As String = txtnamabb.Text.Trim()
        Dim biayaPesan_d As Double
        Dim biayaSimpan_d As Double
        Dim demand_d As Double
        Dim leadTime_d As Double
        If Not Double.TryParse(txtbiaya.Text, biayaPesan_d) Then
            MsgBox("Biaya pesan harus berupa angka.", MsgBoxStyle.Exclamation)
            txtbiaya.Focus()
            Return
        End If
        If Not Double.TryParse(txtsimpan.Text, biayaSimpan_d) Then
            MsgBox("Biaya simpan harus berupa angka.", MsgBoxStyle.Exclamation)
            txtsimpan.Focus()
            Return
        End If
        If Not Double.TryParse(txtdemand.Text, demand_d) Then
            MsgBox("Demand tidak valid. Pastikan nama bahan baku sudah dipilih sehingga demand terisi otomatis.", MsgBoxStyle.Exclamation)
            txtnamabb.Focus()
            Return
        End If
        If Not Double.TryParse(txtlead.Text, leadTime_d) Then
            MsgBox("Lead time harus berupa angka (satuan minggu).", MsgBoxStyle.Exclamation)
            txtlead.Focus()
            Return
        End If
        If biayaSimpan_d <= 0 Then
            MsgBox("Biaya simpan (holding cost) harus lebih besar dari 0.", MsgBoxStyle.Exclamation)
            txtsimpan.Focus()
            Return
        End If
        Dim EOQ As Double = Math.Sqrt((2.0 * demand_d * biayaPesan_d) / biayaSimpan_d)
        Dim Z As Double = 1.65
        Dim CV As Double = 0.2
        Try
            If Me.Controls.ContainsKey("txtZ") Then
                Dim tb As TextBox = DirectCast(Me.Controls("txtZ"), TextBox)
                If tb.Text.Trim() <> "" Then Double.TryParse(tb.Text, Z)
            End If
            If Me.Controls.ContainsKey("txtCV") Then
                Dim tb2 As TextBox = DirectCast(Me.Controls("txtCV"), TextBox)
                If tb2.Text.Trim() <> "" Then Double.TryParse(tb2.Text, CV)
            End If
        Catch ex As Exception
        End Try
        Dim sd_per_week As Double = CV * demand_d
        Dim sd_lead As Double = sd_per_week * Math.Sqrt(Math.Max(0.0, leadTime_d))
        Dim SS As Double = Z * sd_lead
        Dim ROP As Double = demand_d * leadTime_d + SS
        txteoq.Text = Math.Round(EOQ, 2).ToString()
        txtrop.Text = Math.Round(ROP, 2).ToString()
        txtss.Text = Math.Round(SS, 2).ToString()

        Dim newRow As DataRow = hasilDt.NewRow()
        newRow("Nama_Bahan_Baku") = nama
        newRow("Biaya_Pesan") = biayaPesan_d
        newRow("Biaya_Simpan") = biayaSimpan_d
        newRow("Lead_Time") = leadTime_d
        newRow("Demand") = demand_d
        newRow("EOQ") = Math.Round(EOQ, 2)
        newRow("ROP") = Math.Round(ROP, 2)
        newRow("Safety_Stock") = Math.Round(SS, 2)
        hasilDt.Rows.Add(newRow)

        MsgBox("Perhitungan selesai dan ditambahkan ke daftar hasil.", MsgBoxStyle.Information)
    End Sub
    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)
            If row.Cells("Nama_Bahan_Baku").Value IsNot Nothing Then
                txtnamabb.Text = row.Cells("Nama_Bahan_Baku").Value.ToString()
                txtnamabb_Leave(Nothing, Nothing)
            End If
        End If
    End Sub

    Private Sub btndeleteall_Click(sender As Object, e As EventArgs) Handles btndeleteall.Click
        If DataGridView2.SelectedRows.Count > 0 Then
            For Each r As DataGridViewRow In DataGridView2.SelectedRows
                If Not r.IsNewRow Then
                    DataGridView2.Rows.Remove(r)
                End If
            Next
        Else
            MsgBox("Pilih baris hasil yang ingin dihapus.", MsgBoxStyle.Exclamation)
        End If
    End Sub

    Private Sub txtnamabb_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtnamabb.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> " "c Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtbiaya_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtbiaya.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtsimpan_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtsimpan.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtlead_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtlead.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.Show()
        Me.Hide()
        txtsearch.Clear()
        form1.txtnoid.Clear()
    End Sub

    Private Sub btnback_Click(sender As Object, e As EventArgs) Handles btnback.Click
        Form12.Show()
        Me.Hide()
        txtsearch.Clear()
    End Sub
End Class