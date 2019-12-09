
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting


Public Class Admin
    Public id As Integer
    Public reportId As Integer = 0
    Public reportId2 As Integer = 0
    Public ItemsCol As String
    Public QuanCol As String
    Public ChartArea As String

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        Me.Close()
        Home.Show()
        Home.TextBox1.Text = ""
        Pass.TextBox1.Text = ""
        Create.TextBox1.Text = ""
        Create.TextBox3.Text = ""
        Create.TextBox4.Text = ""

    End Sub



    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Connect.ConnectionString = Connection
        Dim sql As String = "Select * from tblReserve where Lab ='" & ComboBox1.SelectedItem.ToString() & "' and Date ='" & DateTimePicker1.Value.Date & "' and Time1 ='" & DateTimePicker2.Value.ToString("hh:mm tt") & "' and Time2 ='" & DateTimePicker3.Value.ToString("hh:mm tt") & "'"
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "res")

        
        Dim count As Integer
        count = ds.Tables("res").Rows.Count()
        If (TextBox1.Text = "") Then
            MessageBox.Show("Name not specified!")

        ElseIf (count > 0) Then
            MessageBox.Show("Laboratory already reserved!")

        ElseIf (DateTimePicker2.Value.ToString("hh:mm tt") = DateTimePicker3.Value.ToString("hh:mm tt")) Then
            MessageBox.Show("Invalid time setting!")
       
        Else

            Dim sql1 As String = "Insert into tblReserve(FuName,Lab,Date,Time1,Time2)values('" & TextBox1.Text & "','" & ComboBox1.Text & "','" & DateTimePicker1.Value.Date & "','" & DateTimePicker2.Value.ToString("hh:mm tt") & "','" & DateTimePicker3.Value.ToString("hh:mm tt") & "')"
            query(sql1)
            TextBox1.Text = ""
            MessageBox.Show("Transaction Reserved!")
        End If

        Dim sql2 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End'  from tblReserve"
        RefreshTable(sql2)
        DataGridView20.DataSource = binding

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Items.ShowDialog()
        Dim sql As String = "select Quantity, Chemical as 'Chemical Name' from tblChemicals"
        query(sql)
        RefreshTable(sql)
        DataGridView1.DataSource = binding
    End Sub

    Private Sub ToolStripStatusLabel2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Private Sub ToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Connect.Open()

        If (ToolStripTextBox4.Text = ToolStripTextBox5.Text) Then
            Dim sql As String = "Update tblAccounts set Pass = '" & ToolStripTextBox5.Text & "' where ID = '" & ToolStripStatusLabel2.Text & "' "
            query(sql)
            MessageBox.Show("Password successfully created!")
            RefreshTable(sql)

        Else
            MessageBox.Show("Password didn't match!")
            ToolStripTextBox4.Text = ""
            ToolStripTextBox5.Text = ""
        End If

        Connect.Close()
    End Sub

    Private Sub LogoutToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Me.Close()
        Home.Show()
        Home.TextBox1.Text = ""
        Pass.TextBox1.Text = ""
        Create.TextBox1.Text = ""
        Create.TextBox3.Text = ""
        Create.TextBox4.Text = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim sql As String = "Update tblChemicals set Quantity = '" & TextBox2.Text & "',Chemical = '" & TextBox3.Text & "',DisPro = '" & TextBox5.Text & "',HazClass = '" & TextBox4.Text & "',Remarks = '" & TextBox6.Text & "',LoCode = '" & TextBox4.Text & "', Date = '" & DateTimePicker4.Value.Date & "' where ID = " & id & " "
        query(sql)
        MessageBox.Show("Update Successfull!")
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""

        Dim sql1 As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
        RefreshTable(sql1)
        DataGridView1.DataSource = binding

    End Sub

    Private Sub LogoutToolStripMenuItem1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LogoutToolStripMenuItem1.Click
        Me.Close()
        Home.Show()
        Home.TextBox1.Text = ""
        Pass.TextBox1.Text = ""
        ToolStripStatusLabel2.Text = ""

    End Sub

    Private Sub Admin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        DateTimePicker2.Format = DateTimePickerFormat.Custom
        DateTimePicker2.CustomFormat = "hh:mm tt"
        DateTimePicker3.Format = DateTimePickerFormat.Custom
        DateTimePicker3.CustomFormat = "hh:mm tt"

        Dim sql1 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End'  from tblReserve"
        RefreshTable(sql1)
        DataGridView20.DataSource = binding


    End Sub



    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim sql As String = "Insert into tblChemicals(Quantity,Chemical,DisPro,HazClass,Remarks,LoCode,Date)values('" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox5.Text & "','" & TextBox4.Text & "','" & DateTimePicker4.Value.Date & "','" & TextBox7.Text & "','" & TextBox7.Text & "')"
        query(sql)
        MessageBox.Show("Items successfully added!")
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""

        Dim sql1 As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
        RefreshTable(sql1)
        DataGridView1.DataSource = binding
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals where Chemical like '%" & TextBox8.Text & "%'"
        RefreshTable(sql)
        DataGridView1.DataSource = binding


    End Sub

    Private Sub ToolStripMenuItem2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripMenuItem2.Click
        If (ToolStripTextBox4.Text & ToolStripTextBox5.Text = "") Then
            MessageBox.Show("Please enter a password!")
        ElseIf (ToolStripTextBox4.Text = ToolStripTextBox5.Text) Then
            Dim sql As String = "Update tblAccounts set Pass = '" & ToolStripTextBox4.Text & "' where IDnum = '" & ToolStripStatusLabel2.Text & "' "
            query(sql)
            MessageBox.Show("Password successfully changed!")
            RefreshTable(sql)
        Else
            MessageBox.Show("Password didn't match!")
            ToolStripTextBox4.Text = ""
            ToolStripTextBox5.Text = ""
        End If
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick


        Dim sql1 As String = "Select ID from tblChemicals where Chemical ='" & DataGridView1.CurrentRow.Cells(1).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql1, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "adminid")

        Dim count As Integer
        count = ds.Tables("adminid").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("adminid").Rows(0).Item(0)
        End If
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView1.Rows(e.RowIndex)


            TextBox2.Text = row.Cells.Item(0).Value.ToString
            TextBox3.Text = row.Cells.Item(1).Value.ToString
            TextBox5.Text = row.Cells.Item(2).Value.ToString
            TextBox4.Text = row.Cells.Item(3).Value.ToString
            TextBox6.Text = row.Cells.Item(4).Value.ToString
            TextBox7.Text = row.Cells.Item(5).Value.ToString


        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Dim sql1 As String = "Select ID from tblChemicals where Chemical ='" & DataGridView1.CurrentRow.Cells(1).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql1, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "adminid")

        Dim count As Integer
        count = ds.Tables("adminid").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("adminid").Rows(0).Item(0)
            Dim sql2 As String = "DELETE FROM tblChemicals where  ID = '" & id & "'"
            query(sql2)
            MessageBox.Show("Delete Successfull!")
        End If

        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""

        Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
        RefreshTable(sql)
        DataGridView1.DataSource = binding
    End Sub


    Private Sub AddAccountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AddAccountToolStripMenuItem.Click
        Account.ShowDialog()
    End Sub

    Private Sub TabControl1_Selected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl1.Selected
        Dim a As Integer = Me.TabControl1.SelectedIndex.ToString()
        Dim b As Integer = TabControl2.SelectedIndex().ToString()

       
        If (a = 0) Then
            Dim sql As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End' from tblReserve"
            RefreshTable(sql)
            DataGridView20.DataSource = binding

        ElseIf (a = 1) Then

            If (b = 0) Then
                Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
                RefreshTable(sql)
                DataGridView1.DataSource = binding
            ElseIf (b = 1) Then
                Dim sql As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
                RefreshTable(sql)
                DataGridView2.DataSource = binding

                ComboBox6.Text = "Biology Laboratory"
                ComboBox5.Text = "Equipments"
            ElseIf (b = 2) Then
                Dim sql As String = "select ID as '#', Box_Num as 'Box No.', Slides, UnitQuan as 'Quantity', Description as 'Specimen Name' from tblInventory where Cat_Loc = '" & "Mounted Biological Specimen" & "'"
                RefreshTable(sql)
                DataGridView4.DataSource = binding
            ElseIf (b = 3) Then
                Dim sql As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat_Loc = '" & "Microscope Inventory" & "'"
                RefreshTable(sql)
                DataGridView16.DataSource = binding
            ElseIf (b = 4) Then
                Dim sql As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
                RefreshTable(sql)
                DataGridView3.DataSource = binding

                ComboBox6.Text = "Chemistry Laboratory"
                ComboBox5.Text = "Books and References"
            End If
        
        End If


    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
        TextBox7.Text = ""
    End Sub


    Private Sub TabControl16_Selected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl16.Selected
        Dim a As Integer = TabControl16.SelectedIndex()

        If (a = 0) Then
            Dim sql As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End',Items  from tblReserve"
            RefreshTable(sql)
            DataGridView20.DataSource = binding
        ElseIf (a = 1) Then
            Dim sql As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End',Items  from tblReserve"
            RefreshTable(sql)
            DataGridView22.DataSource = binding
        End If
    End Sub
    Private Sub DataGridView22_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView22.CellClick
        Dim sql1 As String = "Select ID from tblReserve where FuName ='" & DataGridView22.CurrentRow.Cells(0).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql1, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "res")

        Dim count As Integer
        count = ds.Tables("res").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("res").Rows(0).Item(0)
        End If
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView22.Rows(e.RowIndex)


            TextBox13.Text = row.Cells.Item(0).Value.ToString
            ComboBox2.Text = row.Cells.Item(1).Value.ToString
            DateTimePicker6.Text = row.Cells.Item(2).Value.ToString
            DateTimePicker7.Text = row.Cells.Item(3).Value.ToString
            DateTimePicker5.Text = row.Cells.Item(4).Value.ToString
           
        End If
    End Sub

    Private Sub TextBox9_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox9.TextChanged
        Dim sql As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End', Items from tblReserve where FuName like '%" & TextBox9.Text & "%' or Lab like '%" & TextBox9.Text & "%' or Date like '%" & TextBox9.Text & "%' or Time1 like '%" & TextBox9.Text & "%' or Time2 like '%" & TextBox9.Text & "%'"
        RefreshTable(sql)
        DataGridView20.DataSource = binding
    End Sub

    Private Sub DataGridView20_CellClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView20.CellClick
        Dim sql1 As String = "Select ID from tblReserve where FuName ='" & DataGridView20.CurrentRow.Cells(0).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql1, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "res")

        Dim count As Integer
        count = ds.Tables("res").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("res").Rows(0).Item(0)
        End If
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView20.Rows(e.RowIndex)


            TextBox1.Text = row.Cells.Item(0).Value.ToString
            ComboBox1.Text = row.Cells.Item(1).Value.ToString
            DateTimePicker1.Text = row.Cells.Item(2).Value.ToString
            DateTimePicker2.Text = row.Cells.Item(3).Value.ToString
            DateTimePicker3.Text = row.Cells.Item(4).Value.ToString
        End If
    End Sub

    Private Sub TextBox12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox12.TextChanged
        Dim sql As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End', Items from tblReserve where FuName like '%" & TextBox9.Text & "%' or Lab like '%" & TextBox9.Text & "%' or Date like '%" & TextBox9.Text & "%' or Time1 like '%" & TextBox9.Text & "%' or Time2 like '%" & TextBox9.Text & "%'"
        RefreshTable(sql)
        DataGridView20.DataSource = binding
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click



        Dim sql As String = "select FuName as 'Full Name', Items, Quantity  from tblReserve"
        RefreshTable(sql)
        DataGridView22.DataSource = binding
        Items.ShowDialog()

    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Details.ShowDialog()

    End Sub

    

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        Connect.ConnectionString = Connection
        Dim sql As String = "Select* from tblReserve"
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "res")

        Dim count As Integer
        count = ds.Tables("res").Rows.Count()
        
            If (TextBox13.Text = "") Then
                MessageBox.Show("Transaction Error! Reserved person not found!")

            Else
                Dim c As Integer = ListBox1.Items().Count()

                For lop As Integer = 0 To c - 1
                ItemsCol = ItemsCol + ListBox1.Items(lop).ToString() + ","
                QuanCol = QuanCol + ListBox2.Items(lop).ToString() + ","
                Next
            Dim sql1 As String = "Update tblReserve set Items = '" & ItemsCol & "', Quantity = '" & QuanCol & "' where ID = " & id & " "
            query(sql)
            query(sql1)
            RefreshTable(sql1)
                MessageBox.Show("Transaction Reserved!")

                TextBox13.Text = ""
                ListBox1.Items.Clear()
                ListBox2.Items.Clear()
            End If
        Dim sql2 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End',Items  from tblReserve"
        RefreshTable(sql2)
        DataGridView20.DataSource = binding
    End Sub


    


    Private Sub Button9_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        Dim sql As String = "Update tblReserve set FuName = '" & TextBox1.Text & "',Lab = '" & ComboBox1.SelectedItem.ToString() & "',Date = '" & DateTimePicker1.Value.Date & "',Time1 = '" & DateTimePicker2.Value.ToString("hh:mm tt") & "',Time2 = '" & DateTimePicker3.Value.ToString("hh:mm tt") & "' where ID = " & id & " "
        query(sql)

        TextBox2.Text = ""

        Dim sql1 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End'  from tblReserve"
        RefreshTable(sql1)
        DataGridView20.DataSource = binding
    End Sub

    Private Sub Button10_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim sql As String = "Select ID from tblReserve where FuName ='" & DataGridView20.CurrentRow.Cells(0).Value.ToString & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "res")

        Dim count As Integer
        count = ds.Tables("res").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("res").Rows(0).Item(0)
            Dim sql1 As String = "DELETE FROM tblReserve where  ID = '" & id & "'"
            query(sql1)
        End If

        Dim sql2 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End'  from tblReserve"
        RefreshTable(sql2)
        DataGridView20.DataSource = binding
    End Sub

   
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim a, b, c, d, f, g, h, i, j, k, l, m, n, o, p, q, r, s, t, u, v, w, x, y, aa, z, bb, cc As String

        a = ComboBox6.Text = "Biology Laboratory" And ComboBox5.Text = "Equipments"
        b = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Equipments"
        c = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Apparatus"
        d = ComboBox6.Text = "Physics Laboratory" And ComboBox5.Text = "Equipments"
        f = ComboBox6.Text = "Dispensing Room" And ComboBox5.Text = "Apparatus"
        g = ComboBox6.Text = "Dispensing Room" And ComboBox5.Text = "Equipments"
        h = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Apparatus"
        i = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Equipments"
        j = ComboBox6.Text = "Speech and Science Laboratories Head Office" And ComboBox5.Text = "Equipments"
        k = ComboBox6.Text = "Third Floor Corridor - College Building" And ComboBox5.Text = "Display Cabinet 1 - Biological Models"
        l = ComboBox6.Text = "Third Floor Corridor - College Building" And ComboBox5.Text = "Display Cabinet 5 - Improvised Physics Project"
        m = ComboBox6.Text = "Speech Laboratory" And ComboBox5.Text = "Equipments"
        n = ComboBox6.Text = "Biology Laboratory" And ComboBox5.Text = "Furnitures"
        o = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Furnitures"
        p = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Consumables"
        q = ComboBox6.Text = "Physics Laboratory" And ComboBox5.Text = "Furnitures"
        r = ComboBox6.Text = "Physics Laboratory" And ComboBox5.Text = "Improvised Materials"
        s = ComboBox6.Text = "Physics Laboratory" And ComboBox5.Text = "Consumables"
        t = ComboBox6.Text = "Dispensing Room" And ComboBox5.Text = "Consumables"
        u = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Furnitures"
        v = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Woodcraft Construction Kit"
        w = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Consumables"
        x = ComboBox6.Text = "Third Floor Corridor - College Building" And ComboBox5.Text = "Furnitures"
        y = ComboBox6.Text = "Speech and Science Laboratories Head Office" And ComboBox5.Text = "Furnitures"
        z = ComboBox6.Text = "Speech Laboratory" And ComboBox5.Text = "Furnitures"
        aa = ComboBox6.Text = "Repair Area" And ComboBox5.Text = "Furnitures"
        bb = ComboBox6.Text = "Office Supplies" And ComboBox5.Text = "Equipments"
        cc = ComboBox6.Text = "College AVR" And ComboBox5.Text = "Equipments"

        If (a = True) Then
        
            Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
      
        ElseIf (b = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
       
        ElseIf (c = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Apparatus" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
            TextBox15.Text = ""
       
        ElseIf (d = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
       
        ElseIf (f = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Dispensing Room" & "' and Cat = '" & "Apparatus" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
       
        ElseIf (g = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Dispensing Room" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""

      
        ElseIf (h = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Apparatus" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
       
        ElseIf (i = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        
        ElseIf (j = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        
        ElseIf (k = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Third Floor Corridor - College Building" & "' and Cat = '" & "Display Cabinet 5 - Biological Models" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        
        ElseIf (l = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Third Floor Corridor - College Building" & "' and Cat = '" & "Display Cabinet 5 - Improvised Physics Project" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        
        ElseIf (m = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        
        ElseIf (n = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (o = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (p = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Consumables" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (q = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (r = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Improvised Materials" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (s = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Consumables" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (t = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & " Dispensing Room" & "' and Cat = '" & "Consumables" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (u = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (v = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Woodcraft Construction Kit" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (w = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Consumbles" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (x = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Third Floor Corridor - College Building" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (y = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (z = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech Laboratory" & "' and Cat = '" & "Furnitures" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (aa = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Repair Area" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        ElseIf (bb = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Office Supplies" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
       
            DataGridView2.DataSource = 0
            ElseIf (cc = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "College AVR" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
        Else
            MessageBox.Show("Sorry not available!")
            ComboBox6.Text = ""
            ComboBox5.Text = ""
            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""
            DataGridView2.DataSource = 0
    
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
        RefreshTable(sql1)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub DataGridView2_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView2.Rows(e.RowIndex)
            TextBox18.Text = row.Cells.Item(1).Value.ToString()
            TextBox17.Text = row.Cells.Item(2).Value.ToString()
            TextBox16.Text = row.Cells.Item(3).Value.ToString()
            TextBox15.Text = row.Cells.Item(4).Value.ToString()
        End If
    End Sub

    Private Sub Button17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button17.Click
        Dim sql As String = "Insert into tblInventory(Cat_Loc,Cat,UnitQuan,Description,Brand,Remarks,Date)values('" & ComboBox6.Text & "','" & ComboBox5.Text & "','" & TextBox18.Text & "','" & TextBox17.Text & "','" & TextBox16.Text & "','" & TextBox15.Text & "','" & DateTimePicker4.Value.Date & "')"
        query(sql)

        MessageBox.Show("Items successfully added!")
        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox16.Text = ""
        TextBox15.Text = ""

        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
        RefreshTable(sql1)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub Button18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button18.Click
        Dim sql As String = "Update tblInventory set UnitQuan = '" & TextBox18.Text & "',Description = '" & TextBox17.Text & "',Brand = '" & TextBox16.Text & "',Remarks = '" & TextBox15.Text & "',Date = '" & DateTimePicker4.Value.Date & "'  where ID = " & DataGridView2.CurrentRow.Cells(0).Value.ToString() & " "
        query(sql)
        MessageBox.Show("Update Successfull!")
        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox16.Text = ""
        TextBox15.Text = ""

        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where  Cat_Loc='" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
        RefreshTable(sql1)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        Dim sql As String = "Select ID from tblInventory where Description ='" & DataGridView2.CurrentRow.Cells(2).Value.ToString() & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "adminid")

        Dim count As Integer
        count = ds.Tables("adminid").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("adminid").Rows(0).Item(0)
            Dim sql1 As String = "DELETE FROM tblInventory where  ID = '" & DataGridView2.CurrentRow.Cells(0).Value.ToString() & "'"
            query(sql1)
            MessageBox.Show("Delete Successfull!")
        End If

        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox16.Text = ""
        TextBox15.Text = ""

        Dim sql2 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where  Cat_Loc='" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
        RefreshTable(sql2)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        TextBox18.Text = ""
        TextBox17.Text = ""
        TextBox16.Text = ""
        TextBox15.Text = ""
    End Sub

    Private Sub TabControl2_Selected(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl2.Selected
        Dim a As Integer = TabControl2.SelectedIndex().ToString()
        If (a = 0) Then
            Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
            RefreshTable(sql)
            DataGridView1.DataSource = binding
        ElseIf (a = 1) Then
            Dim sql As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql)
            DataGridView2.DataSource = binding

            ComboBox6.Text = "Biology Laboratory"
            ComboBox5.Text = "Equipments"
        ElseIf (a = 2) Then
            Dim sql As String = "select ID as '#', Box_Num as 'Box No.', Slides, UnitQuan as 'Quantity', Description as 'Specimen Name' from tblInventory where Cat_Loc = '" & "Mounted Biological Specimen" & "'"
            RefreshTable(sql)
            DataGridView4.DataSource = binding
        ElseIf (a = 3) Then
            Dim sql As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat_Loc = '" & "Microscope Inventory" & "'"
            RefreshTable(sql)
            DataGridView16.DataSource = binding
        ElseIf (a = 4) Then
            Dim sql As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql)
            DataGridView3.DataSource = binding

            ComboBox6.Text = "Chemistry Laboratory"
            ComboBox5.Text = "Books and References"
        End If
    End Sub

    Private Sub Button134_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button134.Click
        Dim sql As String = "Insert into tblInventory(Cat_Loc,Box_Num,Slides,UnitQuan,Description,Date)values('" & "Mounted Biological Specimen" & "','" & TextBox19.Text & "','" & TextBox20.Text & "','" & TextBox21.Text & "','" & TextBox26.Text & "','" & DateTimePicker4.Value.Date & "')"
        query(sql)

        MessageBox.Show("Items successfully added!")
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox26.Text = ""

        Dim sql1 As String = "select ID as '#', Box_Num as 'Box No.', Slides, UnitQuan as 'Quantity', Description as 'Specimen Name' from tblInventory where Cat_Loc = '" & "Mounted Biological Specimen" & "'"
        RefreshTable(sql1)
        DataGridView4.DataSource = binding
    End Sub

    Private Sub Button135_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button135.Click
        Dim sql As String = "Update tblInventory set Cat_Loc = '" & "Mounted Biological Specimen" & "',Box_Num = '" & TextBox19.Text & "',Slides = '" & TextBox20.Text & "',UnitQuan = '" & TextBox21.Text & "',Description = '" & TextBox26.Text & "',Date = '" & DateTimePicker4.Value.Date & "'  where ID = " & DataGridView4.CurrentRow.Cells(0).Value.ToString() & " "
        query(sql)
        MessageBox.Show("Update Successfull!")
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox26.Text = ""

        Dim sql1 As String = "select ID as '#', Box_Num as 'Box No.', Slides, UnitQuan as 'Quantity', Description as 'Specimen Name' from tblInventory where Cat_Loc = '" & "Mounted Biological Specimen" & "'"
        RefreshTable(sql1)
        DataGridView4.DataSource = binding
    End Sub

    Private Sub Button133_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button133.Click
        Dim sql As String = "Select ID from tblInventory where Description ='" & DataGridView4.CurrentRow.Cells(4).Value.ToString() & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "adminid")

        Dim count As Integer
        count = ds.Tables("adminid").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("adminid").Rows(0).Item(0)
            Dim sql1 As String = "DELETE FROM tblInventory where  ID = '" & DataGridView4.CurrentRow.Cells(0).Value.ToString() & "'"
            query(sql1)
            MessageBox.Show("Delete Successfull!")
        End If

        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox26.Text = ""

        Dim sql2 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Mounted Biological Specimen" & "' "
        RefreshTable(sql2)
        DataGridView4.DataSource = binding
    End Sub

    Private Sub Button132_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button132.Click
        TextBox19.Text = ""
        TextBox20.Text = ""
        TextBox21.Text = ""
        TextBox26.Text = ""
    End Sub

    

    Private Sub DataGridView4_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView4.Rows(e.RowIndex)
            TextBox19.Text = row.Cells.Item(1).Value.ToString()
            TextBox20.Text = row.Cells.Item(2).Value.ToString()
            TextBox21.Text = row.Cells.Item(3).Value.ToString()
            TextBox26.Text = row.Cells.Item(4).Value.ToString()
        End If
    End Sub

    Private Sub DataGridView16_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView16.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView16.Rows(e.RowIndex)
            TextBox28.Text = row.Cells.Item(1).Value.ToString()
            TextBox29.Text = row.Cells.Item(2).Value.ToString()
            TextBox32.Text = row.Cells.Item(3).Value.ToString()
            TextBox33.Text = row.Cells.Item(4).Value.ToString()
            TextBox34.Text = row.Cells.Item(4).Value.ToString()
            TextBox35.Text = row.Cells.Item(4).Value.ToString()
        End If
    End Sub

    Private Sub TextBox159_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox159.TextChanged
        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat ='" & "Mounted Biological Specimen" & "' "
        RefreshTable(sql1)
        DataGridView4.DataSource = binding
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        Dim sql As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat = '" & "Microscope Inventory" & "'"
        RefreshTable(sql)
        DataGridView16.DataSource = binding
    End Sub

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView4.DataSource = binding
    End Sub

    Private Sub Button26_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button26.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView16.DataSource = binding
    End Sub

    Private Sub Button139_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button139.Click
        Dim sql As String = "Update tblInventory set Cat_Loc = '" & "Micrscope Inventory" & "',UnitQuan = '" & TextBox28.Text & "',Brand = '" & TextBox29.Text & "', Asset_Num = '" & TextBox32.Text & "',Model_Num = '" & TextBox33.Text & "',Serial_Num = '" & TextBox34.Text & "',Date = '" & DateTimePicker4.Value.Date & "'  where ID = " & DataGridView16.CurrentRow.Cells(0).Value.ToString() & " "
        query(sql)
        MessageBox.Show("Update Successfull!")
        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""
        TextBox35.Text = ""

        Dim sql1 As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat = '" & "Microscope Inventory" & "'"
        RefreshTable(sql1)
        DataGridView16.DataSource = binding
    End Sub
    Private Sub Button138_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button138.Click
        Dim sql As String = "Insert into tblInventory(Cat_Loc,UnitQuan,Brand,Asset_Num,Model_Num,Serial_Num,Remarks,Date)values('" & "Microscope Inventory" & "','" & TextBox28.Text & "','" & TextBox29.Text & "','" & TextBox32.Text & "','" & TextBox33.Text & "','" & TextBox34.Text & "','" & TextBox35.Text & "','" & DateTimePicker4.Value.Date & "')"
        query(sql)

        MessageBox.Show("Items successfully added!")
        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""
        TextBox35.Text = ""

        Dim sql1 As String = "select ID as '#',UnitQuan as 'Quantity', Brand, Asset_Num as 'Asset No.', Model_Num as 'Model No.', Serial_Num as 'Serial No.', Remarks from tblInventory where Cat_Loc = '" & "Microscope Inventory" & "'"
        RefreshTable(sql1)
        DataGridView16.DataSource = binding
    End Sub

    Private Sub Button137_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button137.Click
        Dim sql As String = "Select ID from tblInventory where Brand ='" & DataGridView16.CurrentRow.Cells(2).Value.ToString() & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "adminid")

        Dim count As Integer
        count = ds.Tables("adminid").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("adminid").Rows(0).Item(0)
            Dim sql1 As String = "DELETE FROM tblInventory where  ID = '" & DataGridView16.CurrentRow.Cells(0).Value.ToString() & "'"
            query(sql1)
            MessageBox.Show("Delete Successfull!")
        End If

        TextBox28.Text = ""
        TextBox29.Text = ""
        TextBox32.Text = ""
        TextBox33.Text = ""
        TextBox34.Text = ""
        TextBox35.Text = ""

        Dim sql2 As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat_Loc = '" & "Microscope Inventory" & "'"
        RefreshTable(sql2)
        DataGridView16.DataSource = binding
    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView3.Rows(e.RowIndex)
            TextBox36.Text = row.Cells.Item(1).Value.ToString()
            TextBox25.Text = row.Cells.Item(2).Value.ToString()
            TextBox24.Text = row.Cells.Item(3).Value.ToString()
            TextBox23.Text = row.Cells.Item(4).Value.ToString()
        End If
    End Sub

    Private Sub Button19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button19.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub TextBox22_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox22.TextChanged
        Dim sql As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
        RefreshTable(sql)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click
        Dim sql As String = "Insert into tblInventory(Cat_Loc,Cat,UnitQuan,Description,Author,Remarks,Date)values('" & ComboBox8.Text & "','" & ComboBox7.Text & "','" & TextBox36.Text & "','" & TextBox25.Text & "','" & TextBox24.Text & "','" & TextBox23.Text & "','" & DateTimePicker4.Value.Date & "')"
        query(sql)

        MessageBox.Show("Items successfully added!")
        TextBox36.Text = ""
        TextBox25.Text = ""
        TextBox24.Text = ""
        TextBox23.Text = ""

        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & ComboBox8.SelectedItem().ToString() & "' and Cat = '" & ComboBox7.SelectedItem().ToString() & "'"
        RefreshTable(sql1)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub Button21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21.Click
        Dim sql As String = "Update tblInventory set UnitQuan = '" & TextBox36.Text & "',Description = '" & TextBox25.Text & "',Author = '" & TextBox24.Text & "',Remarks = '" & TextBox23.Text & "',Date = '" & DateTimePicker4.Value.Date & "'  where ID = " & DataGridView3.CurrentRow.Cells(0).Value.ToString() & " "
        query(sql)
        MessageBox.Show("Update Successfull!")
        TextBox36.Text = ""
        TextBox25.Text = ""
        TextBox24.Text = ""
        TextBox23.Text = ""

        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & ComboBox8.SelectedItem().ToString() & "' and Cat = '" & ComboBox7.SelectedItem().ToString() & "'"
        RefreshTable(sql1)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub Button22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button22.Click
        Dim sql As String = "Select ID from tblInventory where Description ='" & DataGridView3.CurrentRow.Cells(2).Value.ToString() & "' "
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "adminid")

        Dim count As Integer
        count = ds.Tables("adminid").Rows.Count()
        If (count > 0) Then
            id = ds.Tables("adminid").Rows(0).Item(0)
            Dim sql1 As String = "DELETE FROM tblInventory where  ID = '" & DataGridView3.CurrentRow.Cells(0).Value.ToString() & "'"
            query(sql1)
            MessageBox.Show("Delete Successfull!")
        End If

        TextBox36.Text = ""
        TextBox25.Text = ""
        TextBox24.Text = ""
        TextBox23.Text = ""

        Dim sql2 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where  Cat_Loc='" & ComboBox8.SelectedItem().ToString() & "' and Cat = '" & ComboBox7.SelectedItem().ToString() & "'"
        RefreshTable(sql2)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub Button20_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button20.Click
        TextBox36.Text = ""
        TextBox25.Text = ""
        TextBox24.Text = ""
        TextBox23.Text = ""
    End Sub

    Private Sub Button27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button27.Click
        Dim a, b, c, d, f, g, h, i, j, k, l, m As String
        a = ComboBox8.Text = "Chemistry Laboratory" And ComboBox7.Text = "Books and References"
        b = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Books and References"
        c = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Books with Interactive CDs"
        d = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "DVDs"
        f = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Audio CDs"
        g = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Books Donated by Miss Tacardon"
        h = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Manuals"
        i = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Assorted Magazines/Illustrations and Readings"
        j = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "Reel Tape with Manual"
        k = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "35 mm. Turn Table Disk"
        l = ComboBox8.Text = "Speech and Science Laboratories Head Office" And ComboBox7.Text = "45 mm. Turn Table Disk"
        m = ComboBox8.Text = "Physics Laboratory" And ComboBox7.Text = "Books and References"

        If (a = True) Then

            Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (b = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (c = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Books with Interactive CDs" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (d = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "DVDs" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (f = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Audio CDs" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (g = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Books Donated by Miss Tacardon" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""


        ElseIf (h = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Manuals" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (i = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Assorted Magazines/Illustrations and Readings" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (j = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Reel Tape with Manual" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

        ElseIf (k = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "35 mm. Turn Table Disk" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox18.Text = ""
            TextBox17.Text = ""
            TextBox16.Text = ""
            TextBox15.Text = ""

        ElseIf (l = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "45 mm. Turn Table Disk" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""

       

        ElseIf (m = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""
        Else
            MessageBox.Show("Sorry not available!")
            ComboBox8.Text = ""
            ComboBox7.Text = ""
            TextBox36.Text = ""
            TextBox25.Text = ""
            TextBox24.Text = ""
            TextBox23.Text = ""
            DataGridView3.DataSource = 0
        End If
    End Sub

    Private Sub Button1_Click_2(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        reportId = 1
        Items.ShowDialog()
    End Sub

    Private Sub Button14_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        Details.ShowDialog()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Connect.ConnectionString = Connection
        Dim sql As String = "Select* from tblBroken"
        Dim da As SqlDataAdapter = New SqlDataAdapter(sql, Connect)
        Dim ds As New DataSet
        da.Fill(ds, "res")

        Dim count As Integer
        count = ds.Tables("res").Rows.Count()

        If (TextBox10.Text = "") Then
            MessageBox.Show("Name not specified!")

        Else
            Dim c As Integer = ListBox4.Items().Count()

            For lop As Integer = 0 To c - 1
                ItemsCol = ItemsCol + ListBox4.Items(lop).ToString() + ","
                QuanCol = QuanCol + ListBox3.Items(lop).ToString() + ","
            Next
            Dim sql1 As String = "Insert into tblBroken(Name,Type,CourseYear,Sanction,Status,Item,Quantity)values('" & TextBox10.Text & "','" & ComboBox3.SelectedItem().ToString() & "','" & TextBox11.Text & "','" & ComboBox9.SelectedItem().ToString() & "','" & ComboBox4.SelectedItem().ToString() & "', '" & ItemsCol & "', '" & QuanCol & "' )"
            query(sql1)
            RefreshTable(sql1)
            MessageBox.Show("Succesfull!")

            TextBox10.Text = ""
            ComboBox3.Text = "Teacher"
            TextBox11.Text = ""
            ComboBox4.Text = "Pending"
            ListBox4.Items.Clear()
            ListBox3.Items.Clear()
        End If
       
    End Sub

    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click

        Added.ShowDialog()
        
    End Sub

    Private Sub Button30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button30.Click
        reportId2 = 1
        Added.ShowDialog()

    End Sub

    

    Private Sub ToolStripStatusLabel4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripStatusLabel4.TextChanged
        ToolStripStatusLabel4.Text = Home.TextBox1.Text
    End Sub

    Private Sub ToolStripStatusLabel3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub
End Class