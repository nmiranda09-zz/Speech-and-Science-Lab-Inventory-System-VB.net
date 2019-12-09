Imports System.Data.SqlClient
Public Class Create
    Dim id As Integer
    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If (TextBox3.Text & TextBox4.Text = "") Then
            MessageBox.Show("Please enter a password!")
        ElseIf (TextBox4.Text = TextBox3.Text) Then
            Dim sql As String = "Update tblAccounts set Pass = '" & TextBox4.Text & "' where IDNum = '" & TextBox1.Text & "' "
            query(sql)
            MessageBox.Show("Password successfully created!")
            RefreshTable(sql)
            Me.Close()
        Else
            MessageBox.Show("Password didn't match!")
            TextBox3.Text = ""
            TextBox4.Text = ""
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
        Pass.Show()

    End Sub
End Class