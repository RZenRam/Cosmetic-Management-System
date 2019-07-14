Imports System.Data.OleDb
Public Class sales
    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ob As Object
    Dim Dadapter As OleDbDataAdapter
    Dim table As DataTableCollection
    Dim DSet As DataSet
    Public salescnt As Integer
    Dim cnt As Integer
    Public total As Integer
    Public salesid As String
    Public customername As String
    Public mobileno As String
    Dim out As Integer
    Dim avail As String
    Dim tota, totaa As Integer
    Dim sold As String
    Dim temp As Integer
    Dim tempo As Integer
    Private Sub DataGridView1_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bill.Show()
        customername = TextBox1.Text
        mobileno = Val(TextBox2.Text)
        total = Val(LinkLabel2.Text)
        Me.Hide()
    End Sub

    Private Sub sales_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Developing\bak\Cosmetic Man\Cosmetic.accdb")
        con.Open()

        cmd = New OleDbCommand("select count(*) from bill ", con)
        dr = cmd.ExecuteReader()
        dr.Read()
        cnt = dr(0).ToString
        Dim ex As String
        ex = cnt + 1

        If ex <= 9 Then
            Label2.Text = "sa00" + ex
        ElseIf ex >= 99 Then
            Label2.Text = "sa" + ex
        Else
            Label2.Text = "sa0" + ex
        End If

        salesid = Label2.Text

        DSet = New DataSet
        table = DSet.Tables
        Dadapter = New OleDbDataAdapter("select pt_name from product_type", con)
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

        cmd = New OleDbCommand("delete * from tempbill", con)
        cmd.ExecuteNonQuery()

        Dadapter = New OleDbDataAdapter("select * from tempbill", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "tempbill")
        DataGridView1.DataSource = DSet.Tables("tempbill")

        Label14.Text = Date.Today

        ComboBox1.SelectedIndex = -1
        ComboBox2.SelectedIndex = -1

        cmd = New OleDbCommand("select sales_id from sales ")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        
        out = Val(TextBox3.Text)

        cmd = New OleDbCommand("select * from stock where p_name='" & ComboBox2.Text & "'", con)
        dr = cmd.ExecuteReader

        dr.Read()
        avail = dr("q_avail").ToString
        sold = dr("q_sold").ToString
        dr.Close()

        'Label15.Text = avail
        'Label16.Text = sold

        ' avail = Val(temp)
        ' sold = Val(tempo)
        temp = Val(avail)
        tempo = Val(sold)

       

        If avail > out Then

            tota = temp - out

            totaa = tempo + out

            cmd = New OleDbCommand("insert into sales values('" & Label2.Text & "' ,'" & TextBox1.Text & "','" & TextBox2.Text & "','" & Label14.Text & "','" & ComboBox2.Text & "','" & TextBox3.Text & "','" & Label9.Text & "')", con)
            cmd.ExecuteNonQuery()

            cmd = New OleDbCommand("insert into tempbill values('" & Label2.Text & "' ,'" & ComboBox2.Text & "','" & TextBox3.Text & "','" & Label9.Text & "')", con)
            cmd.ExecuteNonQuery()

            cmd = New OleDbCommand("update stock set q_avail='" & tota & "', q_sold='" & totaa & "' where p_name='" & ComboBox2.Text & "'", con)
            cmd.ExecuteNonQuery()

            LinkLabel2.Text = Val(LinkLabel2.Text) + Val(Label9.Text)

            Label9.Text = Val(Label6.Text) * Val(TextBox3.Text)

            If Val(tota) < 5 Then
                MsgBox("Please update stock for further orders")
            End If
        ElseIf avail < out Then
            MsgBox("Item Out of stock")
        End If

        Dadapter = New OleDbDataAdapter("select * from tempbill", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "tempbill")
        DataGridView1.DataSource = DSet.Tables("tempbill")

        '  LinkLabel2.Text = Val(LinkLabel2.Text) + Val(Label9.Text)

        'Label9.Text = Val(Label6.Text) * Val(TextBox3.Text)

        customername = TextBox1.Text
        mobileno = TextBox2.Text
        total = Val(LinkLabel2.Text)

        ComboBox1.SelectedIndex = Nothing
        ComboBox2.SelectedIndex = Nothing
        Label6.Text = ""
        Label9.Text = ""
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        ' Dim selec As String
        'selec = ComboBox2.SelectedText
        cmd = New OleDbCommand("select p_price from stock where p_name= '" & ComboBox2.Text & "' ", con)
        dr = cmd.ExecuteReader()
        While dr.Read
            Label6.Text = dr("p_price").ToString
        End While
        TextBox3.Text = ""
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        DSet = New DataSet
        table = DSet.Tables
        Dadapter = New OleDbDataAdapter("select p_name from stock WHERE pt_name ='" & ComboBox1.Text & "'", con)
        Dadapter.Fill(DSet, "stock")
        Dim view1 As New DataView(table(0))
        With ComboBox2
            .DataSource = DSet.Tables("stock")
            .DisplayMember = "p_name"
            .ValueMember = "p_name"
            .SelectedIndex = -1
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
        End With
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        out = Val(TextBox3.Text)

        cmd = New OleDbCommand("select * from stock where p_name='" & ComboBox2.Text & "'", con)
        dr = cmd.ExecuteReader

        dr.Read()
        avail = dr("q_avail").ToString
        sold = dr("q_sold").ToString
        dr.Close()

        'Label15.Text = avail
        'Label16.Text = sold

        ' avail = Val(temp)
        ' sold = Val(tempo)
        temp = Val(avail)
        tempo = Val(sold)



        If avail > out Then
            Label9.Text = Val(Label6.Text) * Val(TextBox3.Text)
        End If
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        mmenu.Show()
        Me.Hide()
    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class