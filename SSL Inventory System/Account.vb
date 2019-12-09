Imports System.Data.SqlClient
Public Class Account
    Dim id As Integer
    Private Sub Account_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "select FName as 'First Name', LName as 'Last Name',IDNum as 'ID Number', Pass as 'Password',Type  from tblAccounts"
        RefreshTable(sql)
        DataGridView1.DataSource = binding
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = "Insert into tblAccounts(FName,LName,IDNum,Pass,Type)values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & ComboBox1.Text & "')"
        query(sql)
        MessageBox.Show("Account succesfully added!")
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""


        Dim sql1 As String = "select FName as 'First Name', LName as 'Last Name',IDNum as 'ID Number', Pass as ' Password ',Type  from tblAccounts"
        RefreshTable(sql1)
        DataGridView1.DataSource = binding
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Dim sql As String = "Select ID from tblAccounts where FName ='" & DataGridView1.CurrentRow.Cells(0).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "acc")

        Dim count As Integer
        count = ds.Tables("acc").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("acc").Rows(0).Item(0)
            Dim sql1 As String = "DELETE FROM tblAccounts where ID = '" & id & "'"
            query(sql1)
            MessageBox.Show("Account succesfully deleted!")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox3.Text = ""
            TextBox4.Text = ""

        End If

        Dim sql2 As String = "select FName as 'First Name',LName as 'Last Name',IDNum as 'ID Number', Pass as ' Password ',Type  from tblAccounts"
        RefreshTable(sql2)
        DataGridView1.DataSource = binding
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Dim sql As String = "select FName as 'First Name',LName as 'Last Name',IDNum as 'ID Number', Pass as ' Password ',Type from tblAccounts where FName like '%" & TextBox5.Text & "%' or LName like '%" & TextBox5.Text & "%'"
        RefreshTable(sql)
        DataGridView1.DataSource = binding

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.Close()
        Admin.Show()
    End Sub


    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim sql As String = "Select ID from tblAccounts where FName ='" & DataGridView1.CurrentRow.Cells(0).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "acc")

        Dim count As Integer
        count = ds.Tables("acc").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("acc").Rows(0).Item(0)
        End If
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView1.Rows(e.RowIndex)


            TextBox1.Text = row.Cells.Item(0).Value.ToString
            TextBox2.Text = row.Cells.Item(1).Value.ToString
            TextBox3.Text = row.Cells.Item(2).Value.ToString
            TextBox4.Text = row.Cells.Item(3).Value.ToString
            ComboBox1.Text = row.Cells.Item(4).Value.ToString



        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sql As String = "Update tblAccounts set FName = '" & TextBox1.Text & "',Lname = '" & TextBox2.Text & "',IDNum = '" & TextBox3.Text & "',Pass = '" & TextBox4.Text & "',Type = '" & ComboBox1.Text & "' where ID =" & id & " "
        query(sql)
        MessageBox.Show("Account succesfully updated!")
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        ComboBox1.Text = ""


        Dim sql1 As String = "select FName as 'First Name',LName as 'Last Name',IDNum as 'ID Number', Pass as ' Password ',Type  from tblAccounts"
        RefreshTable(sql1)
        DataGridView1.DataSource = binding
    End Sub
End Class