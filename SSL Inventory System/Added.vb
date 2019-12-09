Imports System.Data
Imports System.Data.SqlClient
Public Class Added
    Public id As Integer


    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If (Admin.reportId2 = 1) Then
            Dim sql1 As String = "select Date, FuName as 'Full Name', Lab as 'Reserved Laboratory', Time1 as 'Start Time', Time2 as 'End Time', Items, Quantity from tblReserve where Date >= '" & DateTimePicker1.Value.Date & "' and Date <= '" & DateTimePicker2.Value.Date & "'"

            query(sql1)
            RefreshTable(sql1)
            DataGridView1.DataSource = binding
        Else
            Dim sql1 As String = "select Date, Cat_Loc as 'Category Location', UnitQuan as 'Unit/Quantity', Description, Brand from tblInventory where Date >= '" & DateTimePicker1.Value.Date & "' and Date <= '" & DateTimePicker2.Value.Date & "'"

            query(sql1)
            RefreshTable(sql1)
            DataGridView1.DataSource = binding
        End If
    End Sub
End Class