<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits AprBase.Controls.Forms.Office

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
        Me.tlpMain = New System.Windows.Forms.TableLayoutPanel()
        Me.tlpButtons = New System.Windows.Forms.TableLayoutPanel()
        Me.btnRefresh = New DevComponents.DotNetBar.ButtonX()
        Me.btnExit = New DevComponents.DotNetBar.ButtonX()
        Me.lrMain = New AprBase.Controls.ListResults()
        Me.wbMain = New System.Windows.Forms.WebBrowser()
        Me.tlpMain.SuspendLayout()
        Me.tlpButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'tlpMain
        '
        Me.tlpMain.ColumnCount = 3
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpMain.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpMain.Controls.Add(Me.tlpButtons, 1, 3)
        Me.tlpMain.Controls.Add(Me.lrMain, 1, 1)
        Me.tlpMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tlpMain.Location = New System.Drawing.Point(0, 0)
        Me.tlpMain.Margin = New System.Windows.Forms.Padding(0)
        Me.tlpMain.Name = "tlpMain"
        Me.tlpMain.RowCount = 5
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23.0!))
        Me.tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpMain.Size = New System.Drawing.Size(800, 450)
        Me.tlpMain.TabIndex = 0
        '
        'tlpButtons
        '
        Me.tlpButtons.ColumnCount = 5
        Me.tlpButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80.0!))
        Me.tlpButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.tlpButtons.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 80.0!))
        Me.tlpButtons.Controls.Add(Me.wbMain, 2, 0)
        Me.tlpButtons.Controls.Add(Me.btnRefresh, 0, 0)
        Me.tlpButtons.Controls.Add(Me.btnExit, 4, 0)
        Me.tlpButtons.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tlpButtons.Location = New System.Drawing.Point(5, 422)
        Me.tlpButtons.Margin = New System.Windows.Forms.Padding(0)
        Me.tlpButtons.Name = "tlpButtons"
        Me.tlpButtons.RowCount = 1
        Me.tlpButtons.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.tlpButtons.Size = New System.Drawing.Size(790, 23)
        Me.tlpButtons.TabIndex = 0
        '
        'btnRefresh
        '
        Me.btnRefresh.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnRefresh.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.btnRefresh.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnRefresh.Location = New System.Drawing.Point(0, 0)
        Me.btnRefresh.Margin = New System.Windows.Forms.Padding(0)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(80, 23)
        Me.btnRefresh.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.btnRefresh.TabIndex = 0
        Me.btnRefresh.Text = "Refresh"
        '
        'btnExit
        '
        Me.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton
        Me.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground
        Me.btnExit.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnExit.Location = New System.Drawing.Point(710, 0)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(0)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(80, 23)
        Me.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled
        Me.btnExit.TabIndex = 1
        Me.btnExit.Text = "Exit"
        '
        'lrMain
        '
        Me.lrMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lrMain.Indent = 0
        Me.lrMain.Location = New System.Drawing.Point(5, 5)
        Me.lrMain.Margin = New System.Windows.Forms.Padding(0)
        Me.lrMain.Name = "lrMain"
        Me.lrMain.Size = New System.Drawing.Size(790, 412)
        Me.lrMain.TabIndex = 0
        '
        'wbMain
        '
        Me.wbMain.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.wbMain.Location = New System.Drawing.Point(85, 0)
        Me.wbMain.Margin = New System.Windows.Forms.Padding(0)
        Me.wbMain.MinimumSize = New System.Drawing.Size(20, 20)
        Me.wbMain.Name = "wbMain"
        Me.wbMain.ScriptErrorsSuppressed = True
        Me.wbMain.Size = New System.Drawing.Size(620, 23)
        Me.wbMain.TabIndex = 4
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.tlpMain)
        Me.DoubleBuffered = True
        Me.Name = "frmMain"
        Me.Opacity = 1.0R
        Me.Text = "frmMain"
        Me.tlpMain.ResumeLayout(False)
        Me.tlpButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents tlpMain As TableLayoutPanel
    Private WithEvents tlpButtons As TableLayoutPanel
    Private WithEvents btnRefresh As DevComponents.DotNetBar.ButtonX
    Private WithEvents btnExit As DevComponents.DotNetBar.ButtonX
    Private lrMain As AprBase.Controls.ListResults
    Private WithEvents wbMain As WebBrowser
End Class
