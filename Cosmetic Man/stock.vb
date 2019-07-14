Imports System.Data.OleDb
Public Class stock

    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ob As Object
    Dim Dadapter As OleDbDataAdapter
    Dim table As DataTableCollection
    Dim DSet As DataSet
    Dim count As Integer
    Private Sub GroupBox1_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub stock_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Developing\bak\Cosmetic Man\Cosmetic.accdb")
        con.Open()

        DSet = New DataSet
        table = DSet.Tables

        Dadapter = New OleDbDataAdapter("select * from product_type", con)
        Dadapter.Fill(DSet, "product_type")

            Dim view As New DataView(table(0))
            With ComboBox1
                .DataSource = DSet.Tables("product_type")
                .DisplayMember = "pt_name"
                .ValueMember = "pt_name"
                .SelectedIndex = -1
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
                .AutoCompleteSource = AutoCompleteSource.ListItems
            End With

        Dadapter = New OleDbDataAdapter("select * from product_type", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "product_type")
        Me.DataGridView1.DataSource = DSet.Tables("product_type")

        Dadapter = New OleDbDataAdapter("select * from stock", con)
        Dadapter.Fill(DSet, "stock")

        Dadapter = New OleDbDataAdapter("select * from stock", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "stock")
        Me.DataGridView2.DataSource = DSet.Tables("stock")

        Label8.Text = Date.Today
        RadioButton1.Checked = False
        RadioButton2.Checked = False

    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        mmenu.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If RadioButton2.Checked = True Then
         
            cmd = New OleDbCommand("insert into product_type values('" & TextBox1.Text & "','" & TextBox2.Text & "')", con)
            cmd.ExecuteNonQuery()
            Dadapter = New OleDbDataAdapter("select * from product_type", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "product_type")
            DataGridView1.DataSource = DSet.Tables("product_type")
            MessageBox.Show("Record Inserted successfully", "Save")
        Else

            cmd = New OleDbCommand("select q_avail ,q_sold from stock where p_id='" & TextBox1.Text & "'", con)
            dr = cmd.ExecuteReader

            Dim q_avail As String
            Dim q_load As Integer
            Dim q_total As Integer
            Dim q_sold As Integer

            While dr.Read()

                q_avail = dr("q_avail").ToString
                q_sold = dr("q_sold").ToString
            End While

            dr.Close()

            q_load = Val(TextBox6.Text)

            q_total = Val(q_avail) + Val(q_load)

            q_sold = Val(q_sold)

            cmd = New OleDbCommand("insert into stock values('" & TextBox3.Text & "','" & ComboBox1.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & q_total & "','" & q_sold & "')", con)
            cmd.ExecuteNonQuery()
            Dadapter = New OleDbDataAdapter("select * from stock", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "stock")
            DataGridView2.DataSource = DSet.Tables("stock")
            MessageBox.Show("Record Inserted successfully", "Save")
        End If
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        If RadioButton2.Checked = True Then
            cmd = New OleDbCommand("select count(*) from product_type ", con)
            dr = cmd.ExecuteReader()
            dr.Read()
            count = dr(0).ToString
            Dim extra As String
            extra = count + 1


            If extra <= 9 Then
                TextBox1.Text = "pt00" + extra
            ElseIf extra >= 99 Then
                TextBox1.Text = "pt" + extra
            Else
                TextBox1.Text = "pt0" + extra
            End If

            TextBox2.Text = ""

        Else
            cmd = New OleDbCommand("select count(*) from stock ", con)
            dr = cmd.ExecuteReader()
            dr.Read()
            count = dr(0).ToString
            Dim extra1 As String
            extra1 = count + 1

            If extra1 <= 9 Then
                TextBox3.Text = "p00" + extra1
            ElseIf extra1 >= 99 Then
                TextBox3.Text = "p" + extra1
            Else
                TextBox3.Text = "p0" + extra1
            End If

            TextBox6.Show()
            Label7.Show()


            ComboBox1.SelectedIndex = Nothing
            TextBox4.Text = ""
            TextBox5.Text = ""
            TextBox6.Text = ""
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If RadioButton2.Checked = True Then
            cmd = New OleDbCommand("update product_type set pt_name='" & TextBox2.Text & " ' where pt_id=' " & TextBox1.Text & " ' ", con)
            cmd.ExecuteNonQuery()
            Dadapter = New OleDbDataAdapter("select * from product_type", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "product_type")
            DataGridView1.DataSource = DSet.Tables("product_type")
            MessageBox.Show("Record updated successfully", "Update")
        Else

            cmd = New OleDbCommand("select q_avail, q_sold from stock where p_id='" & TextBox1.Text & "'", con)
            dr = cmd.ExecuteReader

            Dim q_avail As String
            Dim q_load As Integer
            Dim q_total As Integer
            Dim q_sold As Integer

            While dr.Read()

                q_avail = dr("q_avail").ToString
                q_sold = dr("q_sold").ToString
            End While
            dr.Close()
            q_load = Val(TextBox6.Text)

            q_total = Val(q_avail) + Val(q_load)

            q_sold = Val(q_sold)

            cmd = New OleDbCommand("update stock set pt_name=' " & ComboBox1.Text & " ', p_name='" & TextBox4.Text & " ', p_price='" & TextBox5.Text & " ',q_avail='" & q_total & "',q_sold='" & q_sold & "' where p_id=' " & TextBox3.Text & " ' ", con)
            cmd.ExecuteNonQuery()
            Dadapter = New OleDbDataAdapter("select * from stock", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "stock")
            DataGridView2.DataSource = DSet.Tables("stock")
            MessageBox.Show("Record updated successfully", "Update")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If RadioButton2.Checked = True Then
            cmd = New OleDbCommand("delete from product_type where pt_id=' " & TextBox1.Text & " ' ", con)
            cmd.ExecuteNonQuery()
            Dadapter = New OleDbDataAdapter("select * from product_type", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "product_type")
            Me.DataGridView1.DataSource = DSet.Tables("product_type")
            MessageBox.Show("Record deleted successfully", "Delete")
        Else

            cmd = New OleDbCommand("delete from stock where p_id=' " & TextBox3.Text & " ' ", con)
            cmd.ExecuteNonQuery()
            Dadapter = New OleDbDataAdapter("select * from stock", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "stock")
            Me.DataGridView2.DataSource = DSet.Tables("stock")
            MessageBox.Show("Record deleted successfully", "Delete")
        End If
    End Sub


    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        If RadioButton2.Checked = True Then
            cmd = New OleDbCommand("select * from product_type where pt_id=' " & TextBox1.Text & " '", con)
            dr = cmd.ExecuteReader()

            ComboBox1.Show()
            TextBox6.Hide()

            dr.Read()
            TextBox2.Text = dr("pt_name").ToString

            dr.Close()
            Dadapter = New OleDbDataAdapter("select * from product_type", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "product_type")
            Me.DataGridView1.DataSource = DSet.Tables("product_type")
        Else
            cmd = New OleDbCommand("select * from stock where p_id=' " & TextBox3.Text & " '", con)
            dr = cmd.ExecuteReader()

            ComboBox1.Show()
            'TextBox6.Hide()
            'Label7.Hide()

            While dr.Read
                TextBox4.Text = dr("p_name").ToString
                ComboBox1.SelectedText = dr("pt_name").ToString
                TextBox5.Text = dr("p_price").ToString
                Label10.Text = dr("q_avail").ToString
                Label12.Text = dr("q_sold").ToString
            End While
            dr.Close()
            Dadapter = New OleDbDataAdapter("select * from stock", con)
            DSet = New DataSet
            Dadapter.Fill(DSet, "stock")
            Me.DataGridView2.DataSource = DSet.Tables("stock")
        End If
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Me.Close()
    End Sub

    Private Sub RadioButton1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label8.Click

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub
End Class