Public Class mmenu
    Private Sub StockToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StockToolStripMenuItem.Click
        stock.Show()
        Me.Hide()
    End Sub

    Private Sub SalesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalesToolStripMenuItem.Click
        sales.Show()
        Me.Hide()
    End Sub

    Private Sub SupplierToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SupplierToolStripMenuItem.Click
        supplier.Show()
        Me.Hide()
    End Sub

    Private Sub SearchToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SearchToolStripMenuItem.Click
        salessearch.Show()
        Me.Hide()
    End Sub

    Private Sub HistoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HistoryToolStripMenuItem.Click
        history.Show()
        Me.Hide()
    End Sub
End Class