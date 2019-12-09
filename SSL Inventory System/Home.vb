Imports System.Data.SqlClient
Public Class Home

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Connect.ConnectionString = Connection
        Dim id As Integer


        Dim sql1 As String = "Select* from tblAccounts where IDNum ='" & TextBox1.Text & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql1, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "Login")


        Dim count As Integer
        count = ds.Tables("Login").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("Login").Rows(0).Item(0).ToString()

            Me.Hide()
            Pass.ShowDialog()
            Me.TextBox1.Text = ""
        ElseIf (TextBox1.Text = "") Then
            MessageBox.Show("Please enter your ID Number!")
        Else
            MessageBox.Show("Incorrect ID Number!")
            Me.TextBox1.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Me.Close()
        TextBox1.Text = ""
    End Sub
End Class
