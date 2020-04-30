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
        Me.ofDlg = New System.Windows.Forms.OpenFileDialog()
        Me.btnOpenExcelFile = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ofOpenExcelFile
        '
        Me.ofDlg.FileName = "Open Excel File"
        '
        'btnOpenExcelFile
        '
        Me.btnOpenExcelFile.Location = New System.Drawing.Point(499, 74)
        Me.btnOpenExcelFile.Name = "btnOpenExcelFile"
        Me.btnOpenExcelFile.Size = New System.Drawing.Size(156, 33)
        Me.btnOpenExcelFile.TabIndex = 0
        Me.btnOpenExcelFile.Text = "Open Excel File"
        Me.btnOpenExcelFile.UseVisualStyleBackColor = True
        '
        'Main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnOpenExcelFile)
        Me.Name = "Main"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents ofDlg As OpenFileDialog
    Friend WithEvents btnOpenExcelFile As Button
End Class
