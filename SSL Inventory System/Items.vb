Imports System.Data.SqlClient
Public Class Items



    Private Sub Items_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim a As Integer = TabControl2.SelectedIndex().ToString()
        If (a = 0) Then
            Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
            RefreshTable(sql)
            DataGridView1.DataSource = binding
        ElseIf (a = 1) Then
            Dim sql As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql)
            DataGridView2.DataSource = binding

            ComboBox6.Text = "Biology Laboratory"
            ComboBox5.Text = "Equipments"
        ElseIf (a = 2) Then
            Dim sql As String = "select UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat = '" & "Microscope Inventory" & "'"
            RefreshTable(sql)
            DataGridView16.DataSource = binding
        ElseIf (a = 3) Then
            Dim sql As String = "select UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql)
            DataGridView3.DataSource = binding

            ComboBox8.Text = "Chemistry Laboratory"
            ComboBox7.Text = "Books and References"

        End If

    End Sub

    Private Sub TabControl2_Selected_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TabControlEventArgs) Handles TabControl2.Selected
        Dim a As Integer = TabControl2.SelectedIndex().ToString()
        If (a = 0) Then
            Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
            RefreshTable(sql)
            DataGridView1.DataSource = binding
        ElseIf (a = 1) Then
            Dim sql As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql)
            DataGridView2.DataSource = binding

            ComboBox6.Text = "Biology Laboratory"
            ComboBox5.Text = "Equipments"
        ElseIf (a = 2) Then
            Dim sql As String = "select UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat_Loc = '" & "Microscope Inventory" & "'"
            RefreshTable(sql)
            DataGridView16.DataSource = binding
        ElseIf (a = 3) Then
            Dim sql As String = "select UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql)
            DataGridView3.DataSource = binding

            ComboBox8.Text = "Chemistry Laboratory"
            ComboBox7.Text = "Books and References"
        End If
    End Sub

   

    Private Sub Button25_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button25.Click
        Dim data1, data2 As String

        'datagrid-----------------------------------------
        data1 = DataGridView1.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "k") Then
                crcter = "kg"
                lop1 = count1
            ElseIf (data1(lop1) = "g") Then
                crcter = "g"
                lop1 = count1
            ElseIf (data1(lop1) = "L") Then
                crcter = "L"
                lop1 = count1
            ElseIf (data1(lop1) = "m") Then
                crcter = "ml"
                lop1 = count1


            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox21.Text.ToString()

        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "k") Then
                lop1 = count1
            ElseIf (data2(lop1) = "g") Then
                lop1 = count1
            ElseIf (data2(lop1) = "L") Then
                lop1 = count1
            ElseIf (data2(lop1) = "m") Then
                lop1 = count1

            Else
                temp2 = temp2 + data2(lop1)
            End If

        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of order!")
        ElseIf (TextBox21.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")

        Else
            'subtract values of converted Strings
            Dim Difference As Integer = convert1 - convert2
            'convert liwat to String
            Dim ans As String = Difference.ToString() + crcter


            Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals"
            RefreshTable(sql)
            DataGridView1.DataSource = binding

            If (Admin.reportId = 1) Then
                Admin.ListBox4.Items.Add(TextBox2.Text)
                Admin.ListBox3.Items.Add(TextBox1.Text)
            Else
                Admin.ListBox1.Items.Add(TextBox22.Text)
                Admin.ListBox2.Items.Add(TextBox21.Text)
            End If


        End If
    End Sub

    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView1.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "k") Then
                crcter = "kg"
                lop1 = count1
            ElseIf (data1(lop1) = "g") Then
                crcter = "g"
                lop1 = count1
            ElseIf (data1(lop1) = "L") Then
                crcter = "L"
                lop1 = count1
            ElseIf (data1(lop1) = "m") Then
                crcter = "ml"
                lop1 = count1
            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox21.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "k") Then
                lop1 = count1
            ElseIf (data2(lop1) = "g") Then
                lop1 = count1
            ElseIf (data2(lop1) = "L") Then
                lop1 = count1
            ElseIf (data2(lop1) = "m") Then
                lop1 = count1
            Else
                temp2 = temp2 + data2(lop1)
            End If
        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings
        Dim Difference As Integer = convert1 - convert2
        'convert liwat to String
        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox21.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        ElseIf (Admin.reportId = 1) Then
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Description = '" & TextBox4.Text & "'"
            query(sql)
        Else
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblChemicals set Quantity = '" & ans & "' where Chemical = '" & TextBox22.Text & "'"
            query(sql)

            TextBox22.Text = ""
            TextBox21.Text = ""
        End If
        Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
        RefreshTable(sql1)
        DataGridView2.DataSource = binding

        Dim sql2 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End', Items  from tblReserve"
        RefreshTable(sql2)
        Admin.DataGridView22.DataSource = binding

        Me.Hide()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        TextBox22.Text = ""
        TextBox21.Text = ""
    End Sub

    Private Sub Button23_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button23.Click
        TextBox22.Text = ""
        TextBox21.Text = ""
        If (Admin.reportId = 1) Then
            Admin.ListBox4.Items.Clear()
            Admin.ListBox3.Items.Clear()
        Else
            Admin.ListBox1.Items.Clear()
            Admin.ListBox2.Items.Clear()
        End If
        Me.Close()
    End Sub

    Private Sub TextBox8_TextChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox8.TextChanged
        Dim sql As String = "select Quantity, Chemical as 'Chemical Name', DisPro as 'Disposal Procedure', HazClass as 'Hazard Classification', Remarks, LoCode as 'Location Code' from tblChemicals where Chemical like '%" & TextBox8.Text & "%'"
        RefreshTable(sql)
        DataGridView1.DataSource = binding
    End Sub


    Private Sub Button97_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub

    Private Sub Button98_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        
    End Sub

    
    Private Sub DataGridView2_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView1.Rows(e.RowIndex)
            TextBox22.Text = row.Cells.Item(1).Value.ToString()
            TextBox21.Text = row.Cells.Item(0).Value.ToString()
            
        End If
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView3.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "c") Then
                crcter = "pcs"
                lop1 = count1
            ElseIf (data1(lop1) = "s") Then
                crcter = "set"
                lop1 = count1
            ElseIf (data1(lop1) = "u") Then
                crcter = "units"
                lop1 = count1
            ElseIf (data1(lop1) = "p") Then
                crcter = "pairs"
                lop1 = count1


            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox3.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "c") Then
                lop1 = count1
            ElseIf (data2(lop1) = "s") Then
                lop1 = count1
            ElseIf (data2(lop1) = "u") Then
                lop1 = count1
            ElseIf (data2(lop1) = "p") Then
                lop1 = count1

            Else
                temp2 = temp2 + data2(lop1)
            End If

        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings



        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox3.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        Else
            Dim Difference As Integer = convert1 - convert2
            'convert liwat to String
            Dim ans As String = Difference.ToString() + crcter

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc = '" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding
            If (Admin.reportId = 1) Then
                Admin.ListBox4.Items.Add(TextBox4.Text)
                Admin.ListBox3.Items.Add(TextBox3.Text)
            Else
                Admin.ListBox1.Items.Add(TextBox4.Text)
                Admin.ListBox2.Items.Add(TextBox3.Text)
            End If


        End If
    End Sub

    Private Sub DataGridView2_CellClick_1(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView2.Rows(e.RowIndex)
            TextBox2.Text = row.Cells.Item(1).Value.ToString()
            TextBox1.Text = row.Cells.Item(0).Value.ToString()
          
        End If
    End Sub

    Private Sub DataGridView16_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView16.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView16.Rows(e.RowIndex)
            TextBox38.Text = row.Cells.Item(1).Value.ToString()
            TextBox36.Text = row.Cells.Item(0).Value.ToString()

        End If
    End Sub

    Private Sub DataGridView3_CellClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        If e.RowIndex >= 0 Then
            Dim row As DataGridViewRow
            row = Me.DataGridView3.Rows(e.RowIndex)
            TextBox4.Text = row.Cells.Item(1).Value.ToString()
            TextBox3.Text = row.Cells.Item(0).Value.ToString()

        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView2.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "c") Then
                crcter = "pcs"
                lop1 = count1
            ElseIf (data1(lop1) = "s") Then
                crcter = "set"
                lop1 = count1
            ElseIf (data1(lop1) = "u") Then
                crcter = "units"
                lop1 = count1
            ElseIf (data1(lop1) = "p") Then
                crcter = "pairs"
                lop1 = count1


            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox1.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "c") Then
                lop1 = count1
            ElseIf (data2(lop1) = "s") Then
                lop1 = count1
            ElseIf (data2(lop1) = "u") Then
                lop1 = count1
            ElseIf (data2(lop1) = "p") Then
                lop1 = count1

            Else
                temp2 = temp2 + data2(lop1)
            End If
            
        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings

       
       
        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox1.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        Else
            Dim Difference As Integer = convert1 - convert2
            'convert liwat to String
            Dim ans As String = Difference.ToString() + crcter

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc = '" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding
            If (Admin.reportId = 1) Then
                Admin.ListBox4.Items.Add(TextBox2.Text)
                Admin.ListBox3.Items.Add(TextBox1.Text)
            Else
                Admin.ListBox1.Items.Add(TextBox2.Text)
                Admin.ListBox2.Items.Add(TextBox1.Text)
            End If


        End If

    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView2.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "c") Then
                crcter = "pcs"
                lop1 = count1
            ElseIf (data1(lop1) = "s") Then
                crcter = "set"
                lop1 = count1
            ElseIf (data1(lop1) = "u") Then
                crcter = "units"
                lop1 = count1
            ElseIf (data1(lop1) = "p") Then
                crcter = "pairs"
                lop1 = count1
            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox1.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "c") Then
                lop1 = count1
            ElseIf (data2(lop1) = "s") Then
                lop1 = count1
            ElseIf (data2(lop1) = "u") Then
                lop1 = count1
            ElseIf (data2(lop1) = "p") Then
                lop1 = count1
            Else
                temp2 = temp2 + data2(lop1)
            End If
        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings
        Dim Difference As Integer = convert1 - convert2
        'convert liwat to String
        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox1.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        ElseIf (Admin.reportId = 1) Then
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Description = '" & TextBox4.Text & "'"
            query(sql)
        Else
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Description = '" & TextBox2.Text & "'"
            query(sql)

            TextBox2.Text = ""
            TextBox1.Text = ""
        End If
        Dim sql1 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End',Items  from tblReserve"
        RefreshTable(sql1)
        Admin.DataGridView22.DataSource = binding

        Me.Hide()

        
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox2.Text = ""
        TextBox1.Text = ""
        If (Admin.reportId = 1) Then
            
            Admin.ListBox3.Items.Clear()
            Admin.ListBox4.Items.Clear()
        Else
            Admin.ListBox1.Items.Clear()
            Admin.ListBox2.Items.Clear()
        End If
        Me.Close()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox2.Text = ""
        TextBox1.Text = ""
    End Sub

    Private Sub TextBox14_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox14.TextChanged
        Dim sql1 As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & ComboBox6.SelectedItem().ToString() & "' and Cat = '" & ComboBox5.SelectedItem().ToString() & "'"
        RefreshTable(sql1)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView2.DataSource = binding
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView16.DataSource = binding
    End Sub

    Private Sub TextBox30_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox30.TextChanged
        Dim sql As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat = '" & "Microscope Inventory" & "'"
        RefreshTable(sql)
        DataGridView16.DataSource = binding
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim sql As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Description like '%" & TextBox14.Text & "%' or Brand like '%" & TextBox14.Text & "%' or Author like '%" & TextBox14.Text & "%' or Remarks like '%" & TextBox14.Text & "%' or Box_Num like '%" & TextBox14.Text & "%' or Slides like '%" & TextBox14.Text & "%' or Asset_Num like '%" & TextBox14.Text & "%' or Model_Num like '%" & TextBox14.Text & "%' or Serial_Num like '%" & TextBox14.Text & "%'"
        RefreshTable(sql)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        Dim sql As String = "select ID as '#', UnitQuan as 'Unit/Quantity', Description, Author, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Books and References" & "'"
        RefreshTable(sql)
        DataGridView3.DataSource = binding
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim a, b, c, d, f, g, h, i, j, k, l, m, n, o, p, q, r As String

        a = ComboBox6.Text = "Biology Laboratory" And ComboBox5.Text = "Equipments"
        b = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Equipments"
        c = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Apparatus"
        d = ComboBox6.Text = "Physics Laboratory" And ComboBox5.Text = "Equipments"
        f = ComboBox6.Text = "Dispensing Room" And ComboBox5.Text = "Apparatus"
        g = ComboBox6.Text = "Dispensing Room" And ComboBox5.Text = "Equipments"
        h = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Apparatus"
        i = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Equipments"
        j = ComboBox6.Text = "Speech and Science Laboratories Head Office" And ComboBox5.Text = "Equipments"
        k = ComboBox6.Text = "Speech Laboratory" And ComboBox5.Text = "Equipments"
        l = ComboBox6.Text = "Chemistry Laboratory" And ComboBox5.Text = "Consumables"
        m = ComboBox6.Text = "Physics Laboratory" And ComboBox5.Text = "Consumables"
        n = ComboBox6.Text = "Dispensing Room" And ComboBox5.Text = "Consumables"
        o = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Woodcraft Construction Kit"
        p = ComboBox6.Text = "Science Laboratory Service Area" And ComboBox5.Text = "Consumables"
        q = ComboBox6.Text = "Office Supplies" And ComboBox5.Text = "Equipments"
        r = ComboBox6.Text = "College AVR" And ComboBox5.Text = "Equipments"

        If (a = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Biology Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""
           

        ElseIf (b = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (c = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Apparatus" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""
        ElseIf (d = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (f = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Dispensing Room" & "' and Cat = '" & "Apparatus" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

           TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (g = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Dispensing Room" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""


        ElseIf (h = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Apparatus" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (i = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (j = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

       
        ElseIf (k = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech Laboratory" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""
        
        ElseIf (l = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Chemistry Laboratory" & "' and Cat = '" & "Consumables" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (m = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Consumables" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (n = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & " Dispensing Room" & "' and Cat = '" & "Consumables" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""
        
        ElseIf (o = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Woodcraft Construction Kit" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        ElseIf (p = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Science Laboratory Service Area" & "' and Cat = '" & "Consumbles" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        
        ElseIf (q = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Office Supplies" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

            DataGridView2.DataSource = 0
        ElseIf (r = True) Then

            Dim sql1 As String = "select UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "College AVR" & "' and Cat = '" & "Equipments" & "'"
            RefreshTable(sql1)
            DataGridView2.DataSource = binding

            TextBox2.Text = ""
            TextBox1.Text = ""

        Else
            MessageBox.Show("Sorry not available!")
            ComboBox6.Text = ""
            ComboBox5.Text = ""
            TextBox2.Text = ""
            TextBox1.Text = ""
            DataGridView2.DataSource = 0

        End If
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
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

            TextBox4.Text = ""
            TextBox3.Text = ""
            
        ElseIf (b = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

           TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (c = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Books with Interactive CDs" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding
            TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (d = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "DVDs" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (f = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Audio CDs" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (g = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Books Donated by Miss Tacardon" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""


        ElseIf (h = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Manuals" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

           TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (i = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Assorted Magazines/Illustrations and Readings" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (j = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "Reel Tape with Manual" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (k = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "35 mm. Turn Table Disk" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""

        ElseIf (l = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Speech and Science Laboratories Head Office" & "' and Cat = '" & "45 mm. Turn Table Disk" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

            TextBox4.Text = ""
            TextBox3.Text = ""



        ElseIf (m = True) Then

            Dim sql1 As String = "select ID as '#',UnitQuan as 'Unit/Quantity', Description, Brand, Remarks from tblInventory where Cat_Loc ='" & "Physics Laboratory" & "' and Cat = '" & "Books and References" & "'"
            RefreshTable(sql1)
            DataGridView3.DataSource = binding

           TextBox4.Text = ""
            TextBox3.Text = ""

        Else
            MessageBox.Show("Sorry not available!")
            ComboBox8.Text = ""
            ComboBox7.Text = ""

            TextBox4.Text = ""
            TextBox3.Text = ""
            DataGridView3.DataSource = 0
        End If
    End Sub

    Private Sub Button54_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button54.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView16.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "c") Then
                crcter = "pcs"
                lop1 = count1
            ElseIf (data1(lop1) = "s") Then
                crcter = "set"
                lop1 = count1
            ElseIf (data1(lop1) = "u") Then
                crcter = "units"
                lop1 = count1
            ElseIf (data1(lop1) = "p") Then
                crcter = "pairs"
                lop1 = count1


            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox36.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "c") Then
                lop1 = count1
            ElseIf (data2(lop1) = "s") Then
                lop1 = count1
            ElseIf (data2(lop1) = "u") Then
                lop1 = count1
            ElseIf (data2(lop1) = "p") Then
                lop1 = count1

            Else
                temp2 = temp2 + data2(lop1)
            End If

        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings



        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox36.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        Else
            Dim Difference As Integer = convert1 - convert2
            'convert liwat to String
            Dim ans As String = Difference.ToString() + crcter

            Dim sql As String = "select ID as '#',UnitQuan as 'Quantity',Brand,Asset_Num as 'Asset No.', Model_Num as 'Model No.',Serial_Num as 'Serial No.',Remarks from tblInventory where Cat_Loc = '" & "Microscope Inventory" & "'"
            RefreshTable(sql)
            DataGridView16.DataSource = binding
            If (Admin.reportId = 1) Then
                Admin.ListBox4.Items.Add(TextBox38.Text)
                Admin.ListBox3.Items.Add(TextBox36.Text)
            Else
                Admin.ListBox1.Items.Add(TextBox38.Text)
                Admin.ListBox2.Items.Add(TextBox36.Text)
            End If


        End If
    End Sub

    Private Sub Button53_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button53.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView16.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "c") Then
                crcter = "pcs"
                lop1 = count1
            ElseIf (data1(lop1) = "s") Then
                crcter = "set"
                lop1 = count1
            ElseIf (data1(lop1) = "u") Then
                crcter = "units"
                lop1 = count1
            ElseIf (data1(lop1) = "p") Then
                crcter = "pairs"
                lop1 = count1
            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox36.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "c") Then
                lop1 = count1
            ElseIf (data2(lop1) = "s") Then
                lop1 = count1
            ElseIf (data2(lop1) = "u") Then
                lop1 = count1
            ElseIf (data2(lop1) = "p") Then
                lop1 = count1
            Else
                temp2 = temp2 + data2(lop1)
            End If
        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings
        Dim Difference As Integer = convert1 - convert2
        'convert liwat to String
        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox36.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        ElseIf (Admin.reportId = 1) Then
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Description = '" & TextBox4.Text & "'"
            query(sql)
        Else
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Serail_Num = '" & DataGridView16.CurrentRow.Cells(5).Value.ToString() & "'"
            query(sql)

            TextBox2.Text = ""
            TextBox1.Text = ""
        End If
        Dim sql1 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End',Items  from tblReserve"
        RefreshTable(sql1)
        Admin.DataGridView22.DataSource = binding

        Me.Hide()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        Dim data1, data2 As String
        'datagrid-----------------------------------------
        data1 = DataGridView3.CurrentRow.Cells(0).Value.ToString()
        Dim temp1, crcter As String
        Dim count1 = data1.Count()
        For lop1 As Integer = 0 To count1
            If (data1(lop1) = "c") Then
                crcter = "pcs"
                lop1 = count1
            ElseIf (data1(lop1) = "s") Then
                crcter = "set"
                lop1 = count1
            ElseIf (data1(lop1) = "u") Then
                crcter = "units"
                lop1 = count1
            ElseIf (data1(lop1) = "p") Then
                crcter = "pairs"
                lop1 = count1
            Else
                temp1 = temp1 + data1(lop1)
            End If
        Next
        'textbox--------------------------------------
        data2 = TextBox3.Text.ToString()
        Dim temp2 As String
        Dim count2 = data2.Count()
        For lop1 As Integer = 0 To count2
            If (data2(lop1) = "c") Then
                lop1 = count1
            ElseIf (data2(lop1) = "s") Then
                lop1 = count1
            ElseIf (data2(lop1) = "u") Then
                lop1 = count1
            ElseIf (data2(lop1) = "p") Then
                lop1 = count1
            Else
                temp2 = temp2 + data2(lop1)
            End If
        Next

        'subtract and Convert------------------------------------------Evidence ( ---MessageBox.Show(temp1 + temp2 + crcter)---) 
        'convert >>
        Dim convert1, convert2 As Integer
        Integer.TryParse(temp1.ToString(), convert1)
        Integer.TryParse(temp2.ToString(), convert2)
        'subtract values of converted Strings
        Dim Difference As Integer = convert1 - convert2
        'convert liwat to String
        If (convert1 < convert2 Or convert1 = 0) Then
            MessageBox.Show("This item is out of stock!")
        ElseIf (TextBox3.Text = "0" + crcter) Then
            MessageBox.Show("Sorry not allowed!")
        ElseIf (Admin.reportId = 1) Then
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Description = '" & TextBox4.Text & "'"
            query(sql)
        Else
            Dim ans As String = Difference.ToString() + crcter
            Dim sql As String = "Update tblInventory set UnitQuan = '" & ans & "' where Description = '" & TextBox4.Text & "'"
            query(sql)

            TextBox4.Text = ""
            TextBox3.Text = ""
        End If
        Dim sql1 As String = "select FuName as 'Full Name', Lab as 'Laboratory', Date, Time1 as 'Time Start', Time2 as 'Time End',Items  from tblReserve"
        RefreshTable(sql1)
        Admin.DataGridView22.DataSource = binding

        Me.Hide()
    End Sub

    Private Sub Button52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button52.Click
        TextBox38.Text = ""
        TextBox36.Text = ""
        If (Admin.reportId = 1) Then
            Admin.ListBox3.Items.Clear()
            Admin.ListBox4.Items.Clear()
            
        Else
            Admin.ListBox1.Items.Clear()
            Admin.ListBox2.Items.Clear()
        End If
        Me.Close()
    End Sub

    Private Sub Button51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button51.Click
        TextBox38.Text = ""
        TextBox36.Text = ""
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        TextBox3.Text = ""
        TextBox4.Text = ""
        If (Admin.reportId = 1) Then
            Admin.ListBox3.Items.Clear()
            Admin.ListBox4.Items.Clear()
            
        Else
            Admin.ListBox1.Items.Clear()
            Admin.ListBox2.Items.Clear()
        End If
        Me.Close()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        TextBox3.Text = ""
        TextBox4.Text = ""
    End Sub
End Class