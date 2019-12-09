Imports System.Data.SqlClient

Public Class Pass
    Dim typ As Integer
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Connect.ConnectionString = Connection

        Dim sql As String = "Select* from tblAccounts where Pass ='" & TextBox1.Text & "' and  IDNum='" & Home.TextBox1.Text & "'"
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "Login")
        Dim Pass As String
        Dim count As Integer
        count = ds.Tables("Login").Rows.Count()

        If (count > 0) Then
            Pass = ds.Tables("Login").Rows(0).Item(0).ToString()
            If (ds.Tables("Login").Rows(0).Item(0).ToString() = "Admin") Then
                Admin.AccountsToolStripMenuItem.Enabled = True
            ElseIf (ds.Tables("Login").Rows(0).Item(0).ToString() = "Staff") Then
                Admin.AccountsToolStripMenuItem.Enabled = False

            End If
            Me.Hide()
            Admin.ShowDialog()
            


        ElseIf (TextBox1.Text = "") Then
            MessageBox.Show("Please enter your password!")
        Else
            MessageBox.Show("Password not found or Incorrect password!")
            TextBox1.Text = ""
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Hide()
        Home.ShowDialog()
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Create.TextBox1.Text = Home.TextBox1.Text
        Create.ShowDialog()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
        Home.Show()
        Home.TextBox1.Text = ""
    End Sub
End Class