Imports System.Data.SqlClient


Module DBConnection1

    Public Connection As String = "Data Source=.\SQLEXPRESS;AttachDbFilename=D:\My Files\Final SAD Documents (for ring bind)\SSL Inventory System\SSL Inventory System\STORAGE.mdf;Integrated Security=True;Connect Timeout=30;User Instance=True"
    Public Connect As New SqlConnection
    Public binding As New BindingSource()

    Public Sub query(ByVal sql As String)
        Connect.Open()
        Dim cmd As SqlCommand
        cmd = New SqlCommand(sql, Connect)
        cmd.ExecuteNonQuery()
        Connect.Close()
    End Sub
    Public Sub RefreshTable(ByVal sql As String)
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim table As New DataTable()
        da.Fill(table)
        binding.DataSource = table
    End Sub
End Module