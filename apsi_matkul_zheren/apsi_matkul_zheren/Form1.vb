Imports System.Data.OleDb
Public Class form1
    Dim Maximized As Boolean
    Public Shared LoginNoID As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Posisi() As String = {"Admin", "Staff Designer Produksi", "Distributor", "Pemilik UMKM"}
        For Each nPosisi In Posisi
            cmbposisi.Items.Add(nPosisi)
        Next
        cmbposisi.DropDownStyle = ComboBoxStyle.DropDownList
        txtpw.PasswordChar = "*"
    End Sub
    Private Sub btnlogin_Click(sender As Object, e As EventArgs) Handles btnlogin.Click
        Call Koneksi()
        If Conn.State = ConnectionState.Closed Then Conn.Open()
        Dim cmd As New OleDbCommand
        If txtnoid.Text = "" Or txtpw.Text = "" Then
            MessageBox.Show("Mohon lengkapi data terlebih dahulu!", "LOGIN GAGAL!")
            ErrorProvider.Equals(txtnoid, "Lengkapi data terlebih dahulu!")
            txtnoid.Focus()
            Exit Sub
        Else
            cmd = New OleDbCommand("select * from [Data_Akun] where No_ID ='" & txtnoid.Text & " ' and Username ='" & txtusn.Text & " ' and Posisi ='" & cmbposisi.SelectedItem & " ' and Password ='" & txtpw.Text & "'", Conn)
            Dr = cmd.ExecuteReader
            Dr.Read()
            If Dr.HasRows Then
                LoginNoID = txtnoid.Text
                Select Case cmbposisi.SelectedItem
                    Case "Admin"
                        Form2.Show()
                        Me.Hide()
                        txtusn.Text = ""
                        txtpw.Text = ""
                        cmbposisi.SelectedIndex = -1

                    Case "Staff Designer Produksi"
                        Form5.Show()
                        Me.Hide()
                        txtusn.Text = ""
                        txtpw.Text = ""
                        cmbposisi.SelectedIndex = -1

                    Case "Distributor"
                        Form11.Show()
                        Me.Hide()
                        txtusn.Text = ""
                        txtpw.Text = ""
                        cmbposisi.SelectedIndex = -1

                    Case "Pemilik UMKM"
                        Form12.Show()
                        Me.Hide()
                        txtusn.Text = ""
                        txtpw.Text = ""
                        cmbposisi.SelectedIndex = -1
                End Select
            Else
                MessageBox.Show("Data tidak sesuai!", "Login Gagal!")
            End If
        End If
    End Sub

    Private Sub cbshowpw_CheckedChanged(sender As Object, e As EventArgs) Handles cbshowpw.CheckedChanged
        If cbshowpw.Checked Then
            txtpw.PasswordChar = ""
        Else
            txtpw.PasswordChar = "*"
        End If
    End Sub
    Private Sub txtnoid_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtnoid.KeyPress
        If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) Then
            e.Handled = True
        End If
    End Sub
    Private Sub txtusn_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtusn.KeyPress
        If Not Char.IsLetter(e.KeyChar) AndAlso e.KeyChar <> ChrW(Keys.Back) AndAlso e.KeyChar <> " "c Then
            e.Handled = True
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Application.Exit()
    End Sub
End Class
