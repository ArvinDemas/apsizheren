Imports System.Diagnostics
Public Class Form5
    Private Sub btndatabb_Click(sender As Object, e As EventArgs) Handles btndatabb.Click
        Form6.Show()
        Me.Hide()
    End Sub

    Private Sub btndatapp_Click(sender As Object, e As EventArgs) Handles btndatapp.Click
        Form7.Show()
        Me.Hide()
    End Sub

    Private Sub btndatarc_Click(sender As Object, e As EventArgs) Handles btndatarc.Click
        Form8.Show()
        Me.Hide()
    End Sub

    Private Sub btncekdatap_Click(sender As Object, e As EventArgs) Handles btncekdatap.Click
        Form9.Show()
        Me.Hide()
    End Sub

    Private Sub btnhome_Click(sender As Object, e As EventArgs) Handles btnhome.Click
        form1.txtnoid.Clear()
        form1.txtusn.Clear()
        form1.cmbposisi.SelectedIndex = -1
        form1.txtpw.Clear()
        form1.Show()
        Me.Hide()
    End Sub

End Class