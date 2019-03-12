<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Friend Class ExceptionPublisherForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.SplitContainer2 = New System.Windows.Forms.SplitContainer
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.PictureBox3 = New System.Windows.Forms.PictureBox
        Me.lblException = New System.Windows.Forms.Label
        Me.ExMsg = New System.Windows.Forms.TextBox
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ctxmSelectAll = New System.Windows.Forms.ToolStripMenuItem
        Me.ctxmCopy = New System.Windows.Forms.ToolStripMenuItem
        Me.cmdReport = New System.Windows.Forms.Button
        Me.btnDetails = New System.Windows.Forms.Button
        Me.PanelBottom = New System.Windows.Forms.Panel
        Me.cmdToClipboard = New System.Windows.Forms.Button
        Me.cmdClose = New System.Windows.Forms.Button
        Me.PanelDetails = New System.Windows.Forms.Panel
        Me.txtDetails = New System.Windows.Forms.TextBox
        Me.SplitContainer1 = New System.Windows.Forms.SplitContainer
        Me.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel1.SuspendLayout()
        Me.SplitContainer2.Panel2.SuspendLayout()
        Me.SplitContainer2.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.PanelBottom.SuspendLayout()
        Me.PanelDetails.SuspendLayout()
        Me.SplitContainer1.Panel1.SuspendLayout()
        Me.SplitContainer1.Panel2.SuspendLayout()
        Me.SplitContainer1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.Window
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.SplitContainer2)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(690, 78)
        Me.Panel1.TabIndex = 3
        '
        'SplitContainer2
        '
        Me.SplitContainer2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer2.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer2.IsSplitterFixed = True
        Me.SplitContainer2.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer2.Name = "SplitContainer2"
        '
        'SplitContainer2.Panel1
        '
        Me.SplitContainer2.Panel1.Controls.Add(Me.PictureBox1)
        '
        'SplitContainer2.Panel2
        '
        Me.SplitContainer2.Panel2.Controls.Add(Me.TableLayoutPanel1)
        Me.SplitContainer2.Size = New System.Drawing.Size(688, 76)
        Me.SplitContainer2.SplitterDistance = 125
        Me.SplitContainer2.TabIndex = 6
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.White
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Image = Global.Amcom.SDC.BaseServices.My.Resources.Resources.amcom_new_logo
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(125, 76)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PictureBox1.TabIndex = 3
        Me.PictureBox1.TabStop = False
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.BackColor = System.Drawing.SystemColors.GradientActiveCaption
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.PictureBox2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.PictureBox3, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.lblException, 1, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.ExMsg, 1, 1)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(559, 76)
        Me.TableLayoutPanel1.TabIndex = 1
        '
        'PictureBox2
        '
        Me.PictureBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox2.Image = Global.Amcom.SDC.BaseServices.My.Resources.Resources.exclamationIcon
        Me.PictureBox2.Location = New System.Drawing.Point(3, 3)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(44, 44)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 0
        Me.PictureBox2.TabStop = False
        '
        'PictureBox3
        '
        Me.PictureBox3.Location = New System.Drawing.Point(3, 53)
        Me.PictureBox3.Name = "PictureBox3"
        Me.PictureBox3.Size = New System.Drawing.Size(34, 20)
        Me.PictureBox3.TabIndex = 1
        Me.PictureBox3.TabStop = False
        '
        'lblException
        '
        Me.lblException.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lblException.Location = New System.Drawing.Point(55, 5)
        Me.lblException.Margin = New System.Windows.Forms.Padding(5)
        Me.lblException.Name = "lblException"
        Me.lblException.Size = New System.Drawing.Size(499, 40)
        Me.lblException.TabIndex = 4
        Me.lblException.Text = "An unexpected error has occurred in this application,  This application is in an " & _
            "unstable state and must close.   The error information below may be logged. "
        Me.lblException.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ExMsg
        '
        Me.ExMsg.BackColor = System.Drawing.SystemColors.Window
        Me.ExMsg.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ExMsg.ContextMenuStrip = Me.ContextMenuStrip1
        Me.ExMsg.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ExMsg.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ExMsg.Location = New System.Drawing.Point(60, 60)
        Me.ExMsg.Margin = New System.Windows.Forms.Padding(10)
        Me.ExMsg.Multiline = True
        Me.ExMsg.Name = "ExMsg"
        Me.ExMsg.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.ExMsg.Size = New System.Drawing.Size(489, 6)
        Me.ExMsg.TabIndex = 5
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ctxmSelectAll, Me.ctxmCopy})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(172, 48)
        '
        'ctxmSelectAll
        '
        Me.ctxmSelectAll.Name = "ctxmSelectAll"
        Me.ctxmSelectAll.Size = New System.Drawing.Size(171, 22)
        Me.ctxmSelectAll.Text = "Select All"
        '
        'ctxmCopy
        '
        Me.ctxmCopy.Name = "ctxmCopy"
        Me.ctxmCopy.Size = New System.Drawing.Size(171, 22)
        Me.ctxmCopy.Text = "Copy to Clipboard"
        '
        'cmdReport
        '
        Me.cmdReport.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdReport.Location = New System.Drawing.Point(365, 11)
        Me.cmdReport.Name = "cmdReport"
        Me.cmdReport.Size = New System.Drawing.Size(93, 24)
        Me.cmdReport.TabIndex = 0
        Me.cmdReport.Text = "Send Report"
        Me.cmdReport.Visible = False
        '
        'btnDetails
        '
        Me.btnDetails.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDetails.Location = New System.Drawing.Point(3, 11)
        Me.btnDetails.Name = "btnDetails"
        Me.btnDetails.Size = New System.Drawing.Size(141, 24)
        Me.btnDetails.TabIndex = 5
        Me.btnDetails.Text = "Show Details"
        '
        'PanelBottom
        '
        Me.PanelBottom.Controls.Add(Me.cmdToClipboard)
        Me.PanelBottom.Controls.Add(Me.cmdClose)
        Me.PanelBottom.Controls.Add(Me.btnDetails)
        Me.PanelBottom.Controls.Add(Me.cmdReport)
        Me.PanelBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.PanelBottom.Location = New System.Drawing.Point(0, 343)
        Me.PanelBottom.Name = "PanelBottom"
        Me.PanelBottom.Size = New System.Drawing.Size(690, 44)
        Me.PanelBottom.TabIndex = 5
        '
        'cmdToClipboard
        '
        Me.cmdToClipboard.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdToClipboard.Location = New System.Drawing.Point(161, 11)
        Me.cmdToClipboard.Name = "cmdToClipboard"
        Me.cmdToClipboard.Size = New System.Drawing.Size(178, 24)
        Me.cmdToClipboard.TabIndex = 7
        Me.cmdToClipboard.Text = "Copy Details to Clipboard"
        '
        'cmdClose
        '
        Me.cmdClose.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.cmdClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdClose.Location = New System.Drawing.Point(601, 11)
        Me.cmdClose.Name = "cmdClose"
        Me.cmdClose.Size = New System.Drawing.Size(86, 24)
        Me.cmdClose.TabIndex = 6
        Me.cmdClose.Text = "Close"
        '
        'PanelDetails
        '
        Me.PanelDetails.Controls.Add(Me.txtDetails)
        Me.PanelDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PanelDetails.Location = New System.Drawing.Point(0, 0)
        Me.PanelDetails.Name = "PanelDetails"
        Me.PanelDetails.Size = New System.Drawing.Size(690, 261)
        Me.PanelDetails.TabIndex = 6
        '
        'txtDetails
        '
        Me.txtDetails.BackColor = System.Drawing.Color.Wheat
        Me.txtDetails.ContextMenuStrip = Me.ContextMenuStrip1
        Me.txtDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtDetails.Location = New System.Drawing.Point(0, 0)
        Me.txtDetails.Multiline = True
        Me.txtDetails.Name = "txtDetails"
        Me.txtDetails.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtDetails.Size = New System.Drawing.Size(690, 261)
        Me.txtDetails.TabIndex = 0
        Me.txtDetails.WordWrap = False
        '
        'SplitContainer1
        '
        Me.SplitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer1.FixedPanel = System.Windows.Forms.FixedPanel.Panel1
        Me.SplitContainer1.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer1.Name = "SplitContainer1"
        Me.SplitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer1.Panel1
        '
        Me.SplitContainer1.Panel1.Controls.Add(Me.Panel1)
        '
        'SplitContainer1.Panel2
        '
        Me.SplitContainer1.Panel2.Controls.Add(Me.PanelDetails)
        Me.SplitContainer1.Size = New System.Drawing.Size(690, 343)
        Me.SplitContainer1.SplitterDistance = 78
        Me.SplitContainer1.TabIndex = 5
        '
        'ExceptionPublisherForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(690, 387)
        Me.Controls.Add(Me.SplitContainer1)
        Me.Controls.Add(Me.PanelBottom)
        Me.Font = New System.Drawing.Font("Verdana", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "ExceptionPublisherForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Application Exception"
        Me.TopMost = True
        Me.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel1.ResumeLayout(False)
        Me.SplitContainer2.Panel2.ResumeLayout(False)
        Me.SplitContainer2.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox3, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.PanelBottom.ResumeLayout(False)
        Me.PanelDetails.ResumeLayout(False)
        Me.PanelDetails.PerformLayout()
        Me.SplitContainer1.Panel1.ResumeLayout(False)
        Me.SplitContainer1.Panel2.ResumeLayout(False)
        Me.SplitContainer1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdReport As System.Windows.Forms.Button
    Friend WithEvents lblException As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents btnDetails As System.Windows.Forms.Button
    Friend WithEvents PanelBottom As System.Windows.Forms.Panel
    Friend WithEvents PanelDetails As System.Windows.Forms.Panel
    Friend WithEvents txtDetails As System.Windows.Forms.TextBox
    Friend WithEvents SplitContainer1 As System.Windows.Forms.SplitContainer
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ctxmSelectAll As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ctxmCopy As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExMsg As System.Windows.Forms.TextBox
    Friend WithEvents cmdClose As System.Windows.Forms.Button
    Friend WithEvents SplitContainer2 As System.Windows.Forms.SplitContainer
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdToClipboard As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox3 As System.Windows.Forms.PictureBox
End Class
