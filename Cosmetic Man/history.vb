Imports System.Data.OleDb
Public Class history
    Dim con As New OleDbConnection
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ob As Object
    Dim Dadapter As OleDbDataAdapter
    Dim table As DataTableCollection
    Dim DSet As DataSet
    Private Sub history_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        con = New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=G:\Developing\bak\Cosmetic Man\Cosmetic.accdb")
        con.Open()

        Dadapter = New OleDbDataAdapter("select * from stock", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "stock")
        Me.DataGridView1.DataSource = DSet.Tables("stock")

        Dadapter = New OleDbDataAdapter("select * from sales", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "sales")
        Me.DataGridView2.DataSource = DSet.Tables("sales")

        Dadapter = New OleDbDataAdapter("select * from supplier", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "supplier")
        Me.DataGridView3.DataSource = DSet.Tables("supplier")

        Dadapter = New OleDbDataAdapter("select * from bill", con)
        DSet = New DataSet
        Dadapter.Fill(DSet, "bill")
        Me.DataGridView4.DataSource = DSet.Tables("bill")
    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        mmenu.Show()
        Me.Hide()

    End Sub
End Class