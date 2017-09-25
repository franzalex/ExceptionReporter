<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnBugReport = New System.Windows.Forms.Button()
        Me.btnThrowException = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 3
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.btnBugReport, 1, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.btnThrowException, 1, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(176, 121)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'btnBugReport
        '
        Me.btnBugReport.AutoSize = True
        Me.btnBugReport.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnBugReport.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnBugReport.Location = New System.Drawing.Point(27, 22)
        Me.btnBugReport.MaximumSize = New System.Drawing.Size(160, 0)
        Me.btnBugReport.Name = "btnBugReport"
        Me.btnBugReport.Size = New System.Drawing.Size(121, 25)
        Me.btnBugReport.TabIndex = 0
        Me.btnBugReport.Text = "Report a bug"
        Me.btnBugReport.UseVisualStyleBackColor = True
        '
        'btnThrowException
        '
        Me.btnThrowException.AutoSize = True
        Me.btnThrowException.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.btnThrowException.Dock = System.Windows.Forms.DockStyle.Fill
        Me.btnThrowException.Location = New System.Drawing.Point(27, 73)
        Me.btnThrowException.MaximumSize = New System.Drawing.Size(160, 0)
        Me.btnThrowException.Name = "btnThrowException"
        Me.btnThrowException.Size = New System.Drawing.Size(121, 25)
        Me.btnThrowException.TabIndex = 1
        Me.btnThrowException.Text = "Throw an exception"
        Me.btnThrowException.UseVisualStyleBackColor = True
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(176, 121)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Exception Reporter"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Private WithEvents TableLayoutPanel1 As TableLayoutPanel
    Private WithEvents btnBugReport As Button
    Private WithEvents btnThrowException As Button
End Class
