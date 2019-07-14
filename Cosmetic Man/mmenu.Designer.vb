<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mmenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.StockToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SalesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SupplierToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SearchToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HistoryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(20, 20)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.StockToolStripMenuItem, Me.SalesToolStripMenuItem, Me.SupplierToolStripMenuItem, Me.SearchToolStripMenuItem, Me.HistoryToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1085, 28)
        Me.MenuStrip1.TabIndex = 5
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'StockToolStripMenuItem
        '
        Me.StockToolStripMenuItem.Name = "StockToolStripMenuItem"
        Me.StockToolStripMenuItem.Size = New System.Drawing.Size(57, 24)
        Me.StockToolStripMenuItem.Text = "Stock"
        '
        'SalesToolStripMenuItem
        '
        Me.SalesToolStripMenuItem.Name = "SalesToolStripMenuItem"
        Me.SalesToolStripMenuItem.Size = New System.Drawing.Size(55, 24)
        Me.SalesToolStripMenuItem.Text = "Sales"
        '
        'SupplierToolStripMenuItem
        '
        Me.SupplierToolStripMenuItem.Name = "SupplierToolStripMenuItem"
        Me.SupplierToolStripMenuItem.Size = New System.Drawing.Size(76, 24)
        Me.SupplierToolStripMenuItem.Text = "Supplier"
        '
        'SearchToolStripMenuItem
        '
        Me.SearchToolStripMenuItem.Name = "SearchToolStripMenuItem"
        Me.SearchToolStripMenuItem.Size = New System.Drawing.Size(65, 24)
        Me.SearchToolStripMenuItem.Text = "Search"
        '
        'HistoryToolStripMenuItem
        '
        Me.HistoryToolStripMenuItem.Name = "HistoryToolStripMenuItem"
        Me.HistoryToolStripMenuItem.Size = New System.Drawing.Size(68, 24)
        Me.HistoryToolStripMenuItem.Text = "History"
        '
        'mmenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.BurlyWood
        Me.BackgroundImage = Global.Cosmetic_Man.My.Resources.Resources.Slider_1___Green_Lipstick_Is_A_Neutral____Other_Lessons_From_NYX_Cosmetics
        Me.ClientSize = New System.Drawing.Size(1085, 638)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.Name = "mmenu"
        Me.Text = "Form2"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents StockToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SalesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SupplierToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SearchToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HistoryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
End Class
