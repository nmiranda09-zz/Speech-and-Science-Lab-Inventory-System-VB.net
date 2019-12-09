Imports System.Data.SqlClient
Public Class Details
    Dim id As Integer
    Private Sub Details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim sql As String = "select Name, Type, CourseYear as 'Course and Year',Sanction, Status, Item, Quantity  from tblBroken"
        query(sql)
        RefreshTable(sql)
        DataGridView1.DataSource = binding
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = "Update tblBroken set Status = '" & ComboBox1.Text & "' where ID =" & id & " "
        query(sql)
        MessageBox.Show("Update Successfull!")
        ComboBox1.Text = ""
       

        Dim sql1 As String = "select Name, Type, CourseYear as 'Course and Year', Item, Status  from tblBroken"
        RefreshTable(sql1)
        DataGridView1.DataSource = binding

        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim sql As String = "Select ID from tblBroken where Name ='" & DataGridView1.CurrentRow.Cells(0).Value.ToString & "' "
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
            ComboBox1.Text = row.Cells.Item(4).Value.ToString


        End If
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        Dim sql As String = "select Name, Type, CourseYear as 'Course and Year', Item, Status  from tblBroken where Name like '%" & TextBox2.Text & "%'"
        RefreshTable(sql)
        DataGridView1.DataSource = binding
    End Sub
End Class